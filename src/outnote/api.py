"""Optional helpers for running OutNote behind your own API. Standard library only, framework-free,
and nothing here runs unless you import it, so the client itself stays stateless.

    from outnote.api import ApiCache, ClientPool, error_response, respond

    clients = ClientPool(load_token=my_db.load, save_token=my_db.save)
    cache = ApiCache(fresh=30, stale=600)

    entry = cache.get(user, "mail", lambda: clients.get(user).emails(top=25))
    status, headers, body = respond(entry, request_headers.get("If-None-Match"))

See docs/frontend-and-api.md.
"""

import hashlib
import json
import logging
import threading
import time
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any

from entrance import Broker

from ._graph import OutNote

__all__ = ["ApiCache", "ClientPool", "Entry", "error_response", "respond"]

log = logging.getLogger("outnote.api")


@dataclass
class Entry:
    """One cached result: the data, its JSON `body`, an `etag`, and how old it is."""

    body: bytes
    etag: str
    at: float
    data: Any

    @property
    def age(self):
        return time.monotonic() - self.at


class ApiCache:
    """A per-user, in-memory cache of results. Safe to share between threads.

    Up to `fresh` seconds old, a result is answered from memory. Up to `stale` seconds, it's answered
    from memory and refreshed in the background. Older than that, the caller waits for a new one, and
    gets the old one only if the refresh fails. Identical requests share one fetch.

    It holds real mail and chats in memory, keyed by the `user` you pass, so that value must come from
    your own authentication. For several servers, keep this logic but store entries somewhere shared."""

    def __init__(self, fresh=30, stale=600, max_entries=5000, workers=8):
        self.fresh, self.stale, self.max_entries = fresh, stale, max_entries
        self._entries: OrderedDict = OrderedDict()
        self._running: dict = {}  # key -> Future, so identical requests share one fetch
        self._generation: dict = {}  # user -> counter, so a slow old fetch can't overwrite a write
        self._lock = threading.RLock()
        self._pool = ThreadPoolExecutor(workers, thread_name_prefix="outnote-api")

    def get(self, user, name, fetch, fresh=None, stale=None):
        """The `Entry` for this user's `name`, calling `fetch()` if needed. `fetch` must return
        JSON-serializable data. Raises whatever `fetch` raises if there's nothing old to fall back on."""
        fresh = self.fresh if fresh is None else fresh
        stale = self.stale if stale is None else stale
        key = (user, name)
        with self._lock:
            entry = self._entries.get(key)
            if entry and entry.age < fresh:
                self._entries.move_to_end(key)
                return entry
            future = self._refresh(key, fetch)
        if entry and entry.age < stale:
            return entry  # the refresh carries on in the background
        try:
            return future.result()
        except Exception:
            if entry:  # something old beats an error
                log.warning("serving old data for %s after a failed refresh", name, exc_info=True)
                return entry
            raise

    def invalidate(self, user, prefix=""):
        """Forget this user's results whose name starts with `prefix`. Call it after every write that
        could change them, and with no prefix when the user signs out."""
        with self._lock:
            self._generation[user] = self._generation.get(user, 0) + 1
            for key in [k for k in self._entries if k[0] == user and k[1].startswith(prefix)]:
                del self._entries[key]
            for key in [k for k in self._running if k[0] == user]:
                del self._running[key]

    def close(self):
        """Stop the background threads, for a clean shutdown."""
        self._pool.shutdown(wait=False, cancel_futures=True)

    def _refresh(self, key, fetch):
        future = self._running.get(key)
        if not future:
            generation = self._generation.get(key[0], 0)
            future = self._pool.submit(self._load, key, fetch, generation)
            self._running[key] = future
        return future

    def _load(self, key, fetch, generation):
        try:
            data = fetch()
            body = json.dumps(data, separators=(",", ":")).encode()
            etag = '"' + hashlib.sha256(body).hexdigest()[:32] + '"'
            entry = Entry(body, etag, time.monotonic(), data)
            with self._lock:
                if self._generation.get(key[0], 0) == generation:
                    self._entries[key] = entry
                    self._entries.move_to_end(key)
                    while len(self._entries) > self.max_entries:
                        self._entries.popitem(last=False)
            return entry
        finally:
            with self._lock:
                if self._generation.get(key[0], 0) == generation:  # a newer fetch keeps its place
                    self._running.pop(key, None)


class ClientPool:
    """One `OutNote` client per user, built from refresh tokens you store yourself.

    `load_token(user)` returns the user's refresh token. `save_token(user, token)` is called whenever
    it rotates: save it, or that user may eventually have to sign in again. Refresh tokens work like
    passwords, so keep them encrypted. Token requests for one user run one at a time, so threads can't
    use a token that has just been replaced."""

    def __init__(self, load_token, save_token=None, tenant="common", native=True, max_clients=1000):
        self.load_token, self.save_token = load_token, save_token
        self.tenant, self.native, self.max_clients = tenant, native, max_clients
        self._clients: OrderedDict = OrderedDict()
        self._lock = threading.Lock()

    def get(self, user):
        """The client for `user`, created on first use. The least recently used are dropped past
        `max_clients`."""
        with self._lock:
            client = self._clients.get(user)
            if client:
                self._clients.move_to_end(user)
                return client
        token = self.load_token(user)
        if not token:
            raise LookupError(f"no refresh token for {user!r}")
        client = OutNote(self._tokens(user, token), native=self.native)
        with self._lock:
            client = self._clients.setdefault(user, client)
            self._clients.move_to_end(user)
            while len(self._clients) > self.max_clients:
                self._clients.popitem(last=False)
        return client

    def drop(self, user):
        """Forget a user's client and its tokens in memory. Call it when the user signs out, and delete
        the stored refresh token yourself."""
        with self._lock:
            self._clients.pop(user, None)

    def _tokens(self, user, refresh_token):
        broker = Broker(refresh_token, self.tenant)
        lock = threading.Lock()

        def token(client_id, *scope):
            with lock:
                before = broker.refresh_token
                access = broker(client_id, *scope)
                if broker.refresh_token != before and self.save_token:
                    try:
                        self.save_token(user, broker.refresh_token)
                    except Exception:  # the old token usually still works, so keep serving
                        log.exception("could not save the rotated refresh token for a user")
                return access

        return token


def _matches(if_none_match, etag):
    if not if_none_match:
        return False
    if if_none_match.strip() == "*":
        return True
    return any(tag.strip().removeprefix("W/") == etag for tag in if_none_match.split(","))


def respond(entry, if_none_match=None):
    """`(status, headers, body)` for an `Entry`: 304 with no body if the browser's copy still matches
    the request's `If-None-Match`, otherwise 200 with the JSON. `Cache-Control: private` keeps shared
    caches from storing it."""
    headers = {"ETag": entry.etag, "Cache-Control": "private, no-cache", "Age": str(int(entry.age))}
    if _matches(if_none_match, entry.etag):
        return 304, headers, b""
    return 200, {**headers, "Content-Type": "application/json"}, entry.body


def error_response(error):
    """`(status, headers, body)` for a `GraphError`. 403, 404 and 429 pass through, with `Retry-After`
    when Graph asked for a wait; anything else becomes 502, because it's a failure upstream."""
    status = error.status if error.status in (403, 404, 429) else 502
    headers = {"Content-Type": "application/json"}
    if error.retry_after:
        headers["Retry-After"] = str(int(error.retry_after))
    return status, headers, json.dumps({"error": str(error)}).encode()
