# A native-feeling app on top of OutNote

This page shows how to make a web app feel like Outlook or Teams when OutNote runs behind your own API.
It uses only normal calls like `emails()` and `chats()`, not `stream()` or `SQLiteCache`, so the same
ideas work with any backend.

The API layer is shown twice, so pick one:

- **[With `outnote.api`](#with-outnoteapi):** import the cache and client handling, and write only your routes.
- **[Without `outnote.api`](#without-outnoteapi):** own all the code, for example to change how it works or
  to port it to another language.

Both behave the same and passed the same tests, and [the frontend](#the-frontend) works with either.

```
Browser (IndexedDB) -> Your API (short-lived cache) -> OutNote -> Microsoft
```

The two layers do different jobs, and you can use either one alone:

| Layer | What it does | What the user feels |
|---|---|---|
| Browser cache | Draws saved data immediately, then checks for updates | The app opens instantly, even on a slow connection |
| API cache | Shares results, merges duplicate requests, answers quickly | Fewer waits and fewer Microsoft rate limits |

The code below is a working example with FastAPI and plain JavaScript. It ran against a local API with
fake data in a real browser, and token rotation was tested on a real account. The complete app hasn't run
against Microsoft or with real users. Authentication, token storage and your UI are placeholders that you
need to fill in.

## What makes it feel native

These come from watching the traffic of Outlook and Teams on the web:

1. **Draw first, then fetch.** Show saved data straight away and update it quietly when the answer arrives.
2. **One startup request.** Outlook's `startupdata.ashx` returns config, folders and the first 25
   conversations together. Your `/api/bootstrap` can do the same.
3. **Small lists, details on demand.** `emails()` returns list fields, about 8x faster than every field.
   Fetch the full message with `email(id)` when it's needed.
4. **Load ahead, gently.** Outlook fetches conversation details in the background, about 130 ms apart,
   well after the first screen.
5. **Apply actions immediately.** Show the change first, and undo it if the server refuses.
6. **Keep the app itself offline.** Both web apps use a service worker, so the page opens without waiting
   for the network.

## The API layer

### How the cache works

Both versions decide what to do by how old a result is:

| Age of the result | What happens |
|---|---|
| Under `fresh` seconds | Answers from memory |
| Under `stale` seconds | Answers from memory, and refreshes in the background |
| Older, or nothing saved | Waits for Microsoft. If that fails and something old exists, returns that |

Identical requests share one Microsoft call, and a slow old read can't overwrite a newer write.

### With `outnote.api`

`outnote.api` is an optional module with what an API needs. It uses only the standard library and works
with any framework. Nothing changes for the rest of OutNote unless you import it.

`load_token`, `save_token` and `current_user` are yours to write. **`current_user` must come from your own
authentication**, because the cache trusts it to keep users' data apart.

`ClientPool` keeps one client per user, and calls `save_token` whenever a refresh token rotates. Save it, or
the user may eventually have to sign in again.

```python
# api.py
from concurrent.futures import ThreadPoolExecutor

from fastapi import Depends, FastAPI, HTTPException, Request, Response
from outnote import GraphError
from outnote.api import ApiCache, ClientPool, error_response, respond

app = FastAPI()
cache = ApiCache(fresh=30, stale=600)
_startup = ThreadPoolExecutor(8)


def current_user(request: Request) -> str:
    """Your authentication: return a stable id for the signed-in user, or raise 401."""
    raise HTTPException(401)


def load_token(user: str) -> str:
    """Your storage: return this user's refresh token. Keep it encrypted, like a password."""
    raise NotImplementedError


def save_token(user: str, token: str) -> None:
    """Your storage: replace this user's refresh token with the rotated one."""
    raise NotImplementedError


clients = ClientPool(load_token, save_token)


def reply(request: Request, entry) -> Response:
    status, headers, body = respond(entry, request.headers.get("if-none-match"))
    return Response(body, status, headers)


@app.exception_handler(GraphError)
def graph_error(request: Request, error: GraphError):
    """Tell the browser how long Microsoft asked to wait instead of hiding it."""
    status, headers, body = error_response(error)
    return Response(body, status, headers)


@app.get("/api/bootstrap")
def bootstrap(request: Request, user: str = Depends(current_user)):
    """Everything the first screen needs, in one request."""
    client = clients.get(user)

    def load():
        mail = _startup.submit(lambda: cache.get(user, "mail", lambda: client.emails(top=25)).data)
        chats = _startup.submit(lambda: cache.get(user, "chats", lambda: client.chats(25)).data)
        return {"mail": mail.result(), "chats": chats.result()}

    return reply(request, cache.get(user, "bootstrap", load, fresh=5))


@app.get("/api/mail")
def mail(request: Request, user: str = Depends(current_user)):
    client = clients.get(user)
    return reply(request, cache.get(user, "mail", lambda: client.emails(top=25)))


@app.get("/api/mail/{message_id}")
def message(message_id: str, request: Request, user: str = Depends(current_user)):
    client = clients.get(user)
    fetch = lambda: client.email(message_id)  # noqa: E731
    return reply(request, cache.get(user, f"mail/{message_id}", fetch, fresh=300, stale=3600))


@app.post("/api/mail/{message_id}/read")
def mark_read(message_id: str, user: str = Depends(current_user)):
    clients.get(user).mark_read(message_id)
    # Writes always go to Microsoft, then every result they could have changed is forgotten.
    cache.invalidate(user, "mail")
    cache.invalidate(user, "bootstrap")
    return {"ok": True}


@app.post("/api/sign-out")
def sign_out(user: str = Depends(current_user)):
    cache.invalidate(user)
    clients.drop(user)
    # Also delete the stored refresh token here, and end your own session.
    return {"ok": True}
```

### Without `outnote.api`

This version uses only OutNote's normal client, and you own every line. It has two files.

First the cache, which needs only the standard library:

```python
# api_cache.py
import hashlib
import json
import logging
import threading
import time
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any

log = logging.getLogger("api_cache")


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
        self._pool = ThreadPoolExecutor(workers, thread_name_prefix="api-cache")

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
```

Then the routes. Again, `load_token`, `save_token` and `current_user` are yours to write, and **`current_user`
must come from your own authentication**. Unlike `ClientPool`, `client_for` here has no limit on how many
clients it keeps, so add one if you serve many users from one process.

```python
# api_plain.py
import json
import threading
from concurrent.futures import ThreadPoolExecutor

from entrance import Broker
from fastapi import Depends, FastAPI, HTTPException, Request, Response
from outnote import GraphError, OutNote

from api_cache import ApiCache

app = FastAPI()
cache = ApiCache(fresh=30, stale=600)
_startup = ThreadPoolExecutor(8)
_clients, _clients_lock = {}, threading.Lock()


def current_user(request: Request) -> str:
    """Your authentication: return a stable id for the signed-in user, or raise 401."""
    raise HTTPException(401)


def load_token(user: str) -> str:
    """Your storage: return this user's refresh token. Keep it encrypted, like a password."""
    raise NotImplementedError


def save_token(user: str, token: str) -> None:
    """Your storage: replace this user's refresh token with the rotated one."""
    raise NotImplementedError


def client_for(user: str) -> OutNote:
    """One client per user, because it keeps that user's access tokens in memory. Token requests run
    one at a time per user, and a rotated refresh token is saved as soon as it appears."""
    with _clients_lock:
        if user in _clients:
            return _clients[user]
        broker = Broker(load_token(user))
        lock = threading.Lock()

        def token(client_id, *scope):
            with lock:
                before = broker.refresh_token
                access = broker(client_id, *scope)
                if broker.refresh_token != before:
                    save_token(user, broker.refresh_token)
                return access

        _clients[user] = OutNote(token)
        return _clients[user]


def reply(request: Request, entry) -> Response:
    headers = {"ETag": entry.etag, "Cache-Control": "private, no-cache", "Age": str(int(entry.age))}
    if request.headers.get("if-none-match") == entry.etag:
        return Response(status_code=304, headers=headers)
    return Response(entry.body, media_type="application/json", headers=headers)


@app.exception_handler(GraphError)
def graph_error(request: Request, error: GraphError):
    """Tell the browser how long Microsoft asked to wait instead of hiding it."""
    status = error.status if error.status in (403, 404, 429) else 502
    headers = {"Retry-After": str(int(error.retry_after))} if error.retry_after else {}
    return Response(json.dumps({"error": str(error)}), status, headers, "application/json")


@app.get("/api/bootstrap")
def bootstrap(request: Request, user: str = Depends(current_user)):
    """Everything the first screen needs, in one request."""
    client = client_for(user)

    def load():
        mail = _startup.submit(lambda: cache.get(user, "mail", lambda: client.emails(top=25)).data)
        chats = _startup.submit(lambda: cache.get(user, "chats", lambda: client.chats(25)).data)
        return {"mail": mail.result(), "chats": chats.result()}

    return reply(request, cache.get(user, "bootstrap", load, fresh=5))


@app.get("/api/mail")
def mail(request: Request, user: str = Depends(current_user)):
    client = client_for(user)
    return reply(request, cache.get(user, "mail", lambda: client.emails(top=25)))


@app.get("/api/mail/{message_id}")
def message(message_id: str, request: Request, user: str = Depends(current_user)):
    client = client_for(user)
    fetch = lambda: client.email(message_id)  # noqa: E731
    return reply(request, cache.get(user, f"mail/{message_id}", fetch, fresh=300, stale=3600))


@app.post("/api/mail/{message_id}/read")
def mark_read(message_id: str, user: str = Depends(current_user)):
    client_for(user).mark_read(message_id)
    # Writes always go to Microsoft, then every result they could have changed is forgotten.
    cache.invalidate(user, "mail")
    cache.invalidate(user, "bootstrap")
    return {"ok": True}


@app.post("/api/sign-out")
def sign_out(user: str = Depends(current_user)):
    cache.invalidate(user)
    with _clients_lock:
        _clients.pop(user, None)
    # Also delete the stored refresh token here, and end your own session.
    return {"ok": True}
```

### Rules for the API layer

These apply to both versions.

- **Key every result by user.** Never cache by URL alone, or one person can receive another's mail.
- **Send `Cache-Control: private`.** Shared proxies and CDNs must never store these responses.
- **Never log response bodies or tokens,** and keep your API on HTTPS.
- **Write straight through.** Don't cache writes, and invalidate what they change.
- **Rate limit per user.** Your users may share Microsoft's limits for these client ids, so one busy user
  can slow everyone down.
- **Several API servers need a shared cache.** `ApiCache` lives in one process's memory. With more than
  one server, or across restarts, keep the same logic but store entries in something like Redis, encrypted.
- **Clear on sign-out.** Call `cache.invalidate(user)` and `clients.drop(user)`, and delete the stored
  refresh token.
- **Treat the cache as sensitive.** It holds real mail and chats, so keep the times short.

Microsoft can take a moment to show a change you just made, so a read right after a write may still
return the old value. The frontend's immediate update covers that gap.

## The frontend

### Draw first, then update

`swr.js` stores responses in IndexedDB and asks your API whether they changed. If they didn't, the API
answers `304`, with no body and no redraw.

```javascript
// swr.js
let dbPromise;

/** Call once after sign-in. One database per user, so accounts never mix. */
export function openCache(userId) {
  dbPromise = new Promise((resolve, reject) => {
    const open = indexedDB.open(`app-cache-${userId}`, 1);
    open.onupgradeneeded = () => open.result.createObjectStore("responses");
    open.onsuccess = () => resolve(open.result);
    open.onerror = () => reject(open.error);
  });
  return dbPromise;
}

/** Call on sign-out. Close the app's other tabs first, or the delete waits for them. */
export async function clearCache(userId) {
  (await dbPromise)?.close();
  dbPromise = undefined;
  await new Promise((resolve) => {
    const del = indexedDB.deleteDatabase(`app-cache-${userId}`);
    del.onsuccess = del.onerror = del.onblocked = () => resolve();
  });
}

async function store(mode, action) {
  const db = await dbPromise;
  return new Promise((resolve, reject) => {
    const request = action(db.transaction("responses", mode).objectStore("responses"));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export const readCache = (url) => store("readonly", (s) => s.get(url));
export const writeCache = (url, value) => store("readwrite", (s) => s.put(value, url));

/**
 * Draws saved data straight away, then updates it if the API has something newer.
 * `render(data, { stale })` can run twice. `stale` is true while the data is unconfirmed.
 * Resolves to the newest data, or the saved data if the network failed.
 */
export async function swr(url, render, { onError } = {}) {
  const saved = await readCache(url);
  if (saved) render(saved.data, { stale: true });
  try {
    const response = await fetch(url, {
      cache: "no-store", // so this code sees the 304 itself
      headers: saved?.etag ? { "If-None-Match": saved.etag } : {},
    });
    if (response.status === 304) {
      render(saved.data, { stale: false });
      return saved.data;
    }
    if (!response.ok) {
      const error = new Error(`${url}: ${response.status}`);
      error.status = response.status;
      error.retryAfter = Number(response.headers.get("Retry-After")) || null;
      throw error;
    }
    const data = await response.json();
    await writeCache(url, { data, etag: response.headers.get("ETag") });
    render(data, { stale: false });
    return data;
  } catch (error) {
    onError?.(error, Boolean(saved));
    if (!saved) throw error;
    return saved.data;
  }
}

/** Loads details ahead of time, one at a time, like Outlook. Skips anything already saved. */
export async function prefetch(urls, { gapMs = 130 } = {}) {
  if (navigator.connection?.saveData) return;
  for (const url of urls) {
    if (document.hidden) return;
    if (await readCache(url)) continue;
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) return; // a 429 or an error: stop and try another time
      await writeCache(url, { data: await response.json(), etag: response.headers.get("ETag") });
    } catch {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, gapMs));
  }
}

/**
 * Shows a change immediately and undoes it if the server refuses.
 * `change(data)` returns the new version of the data saved for `url`.
 */
export async function optimistic(url, change, send, render) {
  const before = await readCache(url);
  if (!before) return send();
  const after = change(before.data);
  // No etag: the next check must fetch, because the server's copy has changed.
  await writeCache(url, { data: after, etag: null });
  render(after, { stale: true });
  try {
    const response = await send();
    if (!response.ok) throw new Error(`${response.status}`);
    return response;
  } catch (error) {
    await writeCache(url, before);
    render(before.data, { stale: false });
    throw error;
  }
}

/** Refreshes now, when the tab comes back, and on a timer while it's visible. Returns a stop function. */
export function keepFresh(refresh, { everyMs = 30000 } = {}) {
  const tick = () => { if (!document.hidden) refresh(); };
  const onVisible = () => tick();
  document.addEventListener("visibilitychange", onVisible);
  const timer = setInterval(tick, everyMs);
  tick();
  return () => {
    clearInterval(timer);
    document.removeEventListener("visibilitychange", onVisible);
  };
}
```

### Putting it together

```javascript
// app.js
import { keepFresh, openCache, optimistic, prefetch, swr } from "./swr.js";

const list = document.querySelector("#mail");

function draw({ mail }, { stale }) {
  list.dataset.stale = stale;
  list.replaceChildren(...mail.map((m) => {
    const row = document.createElement("li");
    row.dataset.id = m.id;
    row.className = m.isRead ? "read" : "unread";
    row.textContent = m.subject; // textContent, so mail can't inject HTML
    row.onclick = () => markRead(m.id).catch(() => {});
    return row;
  }));
}

const markRead = (id) => optimistic(
  "/api/bootstrap",
  (data) => ({ ...data, mail: data.mail.map((m) => (m.id === id ? { ...m, isRead: true } : m)) }),
  () => fetch(`/api/mail/${encodeURIComponent(id)}/read`, { method: "POST" }),
  draw,
);

export async function start(userId) {
  await openCache(userId);
  keepFresh(async () => {
    const data = await swr("/api/bootstrap", draw, {
      onError: (error, hasSaved) => {
        if (!hasSaved) list.textContent = "Can't load your mail right now.";
      },
    }).catch(() => null);
    if (data) prefetch(data.mail.slice(0, 5).map((m) => `/api/mail/${encodeURIComponent(m.id)}`));
  });
}
```

If you show mail as HTML, **sanitize it first** and never insert `body.content` directly, because it's
written by whoever sent it. `textContent`, as above, is the safe default.

### Open without the network

A service worker keeps the app's own files, so the page opens offline. It only handles your static files,
and leaves `/api/` alone so that private data stays under `swr.js`'s control.

```javascript
// sw.js
const VERSION = "app-v1"; // change on every release
const FILES = ["/", "/app.js", "/swr.js"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(VERSION).then((cache) => cache.addAll(FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== location.origin) return;
  if (url.pathname.startsWith("/api/")) return;
  event.respondWith(caches.match(event.request).then((hit) => hit || fetch(event.request)));
});
```

Register it from your page with `navigator.serviceWorker.register("/sw.js")`.

### Rules for the frontend

- **One database per user, cleared on sign-out.** Call `clearCache(userId)`, or the next person on a shared
  computer could read the previous user's mail. IndexedDB isn't encrypted.
- **Respect `Retry-After`.** If `error.retryAfter` is set, wait that long before trying again.
- **Don't poll too often.** `keepFresh` stops while the tab is hidden, and 30 seconds is a fair start.
- **There's no push here.** Teams and Outlook get updates over their own push connections, so this approach
  has a delay up to your polling interval.

## Choosing the times

| Data | `fresh` | `stale` | Why |
|---|---|---|---|
| Startup bundle | 5 s | 10 min | Reads shared results, so it can be rebuilt often |
| Mail and chat lists | 30 s | 10 min | Changes often, but a short delay is acceptable |
| One message | 5 min | 1 hour | Rarely changes after it arrives |
| Folders, teams, people | 5-15 min | 1 day | Rarely changes |

These are sensible starting points, not measured values. Adjust them to your app.

## See also

- [caching.md](caching.md) covers `SQLiteCache` and `stream()`, which fetch only what changed. They're more
  efficient than re-fetching lists, so consider them once this layout works for you.
