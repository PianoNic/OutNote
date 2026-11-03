"""Graph transport: GET/POST/PATCH/PUT/DELETE + Bearer + JSON over one shared
keep-alive HTTP client (Entrance's), with 429/503 backoff, nextLink paging,
$batch and delta."""

import base64
import io
import json
import logging
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request

from entrance import HTTPError as ClientError
from entrance import TokenError, http_client

log = logging.getLogger("outnote.http")

GRAPH = "https://graph.microsoft.com/v1.0"
_MAX_PAGES = 50  # ponytail: hard stop so follow-the-nextLink can't loop forever
_TIMEOUT = 30  # seconds per request, so a hung connection can't block forever
_RETRIES = 4  # retries on 429/502/503/504, honoring Retry-After
_RETRY_STATUS = (429, 502, 503, 504)
_RETRY_BUDGET = 60  # seconds one call may spend waiting on Retry-After before giving up
_BATCH_MAX = 20  # Graph's per-$batch request limit
_BATCH_WORKERS = 4  # Graph allows 4 concurrent requests per mailbox

# FOCI family clients, each pre-consented for the scopes that service needs.
C_OFFICE = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # mail, calendar, chat, channels, files, teams, contacts, tasks
C_TEAMS = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"  # Microsoft Teams: OneNote (Notes.ReadWrite.All)
C_OFFICEHOME = (
    "4765445b-32c6-49b0-83e6-1d93765276ca"  # OfficeHome: sites (Sites.Read.All), presence
)
C_ONENOTE = C_TEAMS  # older name, for what this client is used for
C_SITES = C_OFFICEHOME  # older name, for what this client is used for

# Who to ask for directory data (/me, /users, /organization, /me/memberOf). Graph throttles these
# per client: Office can get 429s that Teams and OfficeHome don't, for the same user and address.
_DIRECTORY_CLIENTS = (C_OFFICE, C_TEAMS, C_OFFICEHOME)
_ROTATE_ON = (401, 403, 429)

# Characters left as-is when cleaning a caller's URL: everything URL-structural
# plus `%`, so an already-encoded URL (an @odata.nextLink) passes unchanged
# while spaces, quotes and non-ASCII are percent-encoded.
_URL_SAFE = ":/?#[]@!$&'()*+,;=%~-._"


class GraphError(RuntimeError):
    """A failed call. `retry_after` is how many seconds Graph asked to wait, if it was throttling."""

    def __init__(self, status, message, retry_after=None):
        self.status = status
        self.retry_after = retry_after
        super().__init__(f"Graph {status}: {message}")


class CursorExpired(GraphError):
    """The server no longer accepts a `stream()` cursor: start again without one."""


def _retry_after(headers, attempt):
    ra = (headers or {}).get("Retry-After")
    return float(ra) if ra and str(ra).replace(".", "", 1).isdigit() else 2**attempt


class _Response:
    def __init__(self, body):
        self._body = body

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def urlopen(req, timeout=_TIMEOUT):
    """Send a urllib `Request` over the shared client. Mirrors urllib's contract: a response with
    `.read()`, `HTTPError` for 4xx/5xx and `URLError` for connection failures. Connecting is
    bounded to 5 s so an unreachable address fails over quickly instead of eating the read timeout."""
    t0 = time.perf_counter()
    status = "error"
    try:
        r = http_client().request(
            req.get_method(),
            req.full_url,
            content=req.data,
            headers=dict(req.header_items()),
            timeout=timeout,
            follow_redirects=True,
        )
        status = r.status_code
    except ClientError as e:
        raise URLError(e) from None
    finally:
        log.debug(
            "%s %s -> %s in %.0f ms",
            req.get_method(),
            req.full_url.split("?")[0],
            status,
            (time.perf_counter() - t0) * 1000,
        )
    if r.status_code >= 400:
        raise HTTPError(
            req.full_url, r.status_code, r.reason_phrase, r.headers, io.BytesIO(r.content)
        )  # type: ignore[arg-type]
    return _Response(r.content)


def fetch(url, timeout=_TIMEOUT):
    """GET a pre-authenticated URL (a drive item's downloadUrl) as bytes."""
    try:
        with urlopen(Request(url), timeout=timeout) as r:  # noqa: S310 - Graph-issued URL
            return r.read()
    except HTTPError as e:
        raise GraphError(e.code, e.reason) from None


class Transport:
    def __init__(self, token, base=GRAPH, native=True):
        """`token` is either an access-token string (used for every call) or a
        callable `client_id -> access_token` (the broker from `authenticate`,
        which hands each service the right client's token). `native=False` keeps
        reads on Graph; a static token always does, because it can't reach the
        web clients' own APIs (see `_native.py`)."""
        if not token:
            raise ValueError("token is required")
        self._tp = token if callable(token) else (lambda _client: token)
        self._scoped = callable(token)  # a broker can mint non-Graph tokens
        self.native = native and self._scoped
        self.token = token if isinstance(token, str) else None
        self.base = base.rstrip("/")
        self._ids = {}  # cached me/tenant ids

    # --- transport ----------------------------------------------------------
    def _url(self, path, params=None):
        url = path if path.startswith("http") else self.base + "/" + path.lstrip("/")
        url = quote(url, safe=_URL_SAFE)  # inline queries may carry spaces/umlauts
        if params:
            q = urlencode({k: v for k, v in params.items() if v is not None})
            if q:
                url += ("&" if "?" in url else "?") + q
        return url

    def _send(self, req, raw=False, retry=True):
        """Send `req`; returns parsed JSON, or text (`raw=True`), or bytes
        (`raw="bytes"`). Empty bodies (202/204) come back as None."""
        waited = 0.0
        for attempt in range(_RETRIES + 1):
            try:
                with urlopen(req, timeout=_TIMEOUT) as r:  # noqa: S310 - fixed https host
                    body = r.read()
                if raw == "bytes":
                    return body
                if raw:
                    return body.decode("utf-8")
                return json.loads(body) if body.strip() else None
            except HTTPError as e:
                # Graph throttles with 429 (and 503/502/504); back off and retry.
                wait = _retry_after(e.headers, attempt) if e.code in _RETRY_STATUS else None
                may_wait = retry and wait is not None and attempt < _RETRIES
                if may_wait and waited + wait <= _RETRY_BUDGET:
                    time.sleep(wait)
                    waited += wait
                    continue
                detail = e.read().decode("utf-8", "replace")
                try:
                    detail = json.loads(detail)["error"]["message"]
                except (ValueError, KeyError, TypeError):
                    pass
                raise GraphError(e.code, detail, retry_after=wait) from None
        raise AssertionError("unreachable")

    def request(
        self,
        method,
        path,
        body=None,
        params=None,
        raw=False,
        client=C_OFFICE,
        headers=None,
        content_type=None,
        scope=None,
        retry=True,
    ):
        """One request against `client`'s token. A dict/list `body` is sent as
        JSON; bytes/str are sent as-is with `content_type`. `scope` asks the
        broker for a token to a non-Graph resource instead (absolute URLs)."""
        h = {"Authorization": f"Bearer {self._token(client, scope)}", "Accept": "application/json"}
        data = None
        if body is not None:
            if isinstance(body, (bytes, str)):
                data = body.encode() if isinstance(body, str) else body
                h["Content-Type"] = content_type or "application/octet-stream"
            else:
                data = json.dumps(body).encode()
                h["Content-Type"] = "application/json"
        elif method in ("POST", "PATCH", "PUT"):
            data = b""
        h.update(headers or {})
        req = Request(self._url(path, params), data=data, headers=h, method=method)
        return self._send(req, raw, retry)

    def redirect(self, path, params=None, client=C_OFFICE, timeout=_TIMEOUT):
        """GET `path` without following redirects and return the `Location` it
        points to. Graph answers some downloads (format conversions) with a 302
        to a short-lived pre-authenticated URL; that URL needs no Bearer token,
        so the caller can stream it however it likes."""
        h = {"Authorization": f"Bearer {self._token(client)}"}
        url = self._url(path, params)
        waited = 0.0
        for attempt in range(_RETRIES + 1):
            try:
                r = http_client().request(
                    "GET",
                    url,
                    headers=h,
                    timeout=timeout,
                    follow_redirects=False,
                )
            except ClientError as e:
                raise URLError(e) from None
            wait = _retry_after(r.headers, attempt) if r.status_code in _RETRY_STATUS else None
            if wait is not None and attempt < _RETRIES and waited + wait <= _RETRY_BUDGET:
                time.sleep(wait)
                waited += wait
                continue
            location = r.headers.get("Location")
            if 300 <= r.status_code < 400 and location:
                return location
            detail = r.text
            try:
                detail = json.loads(detail)["error"]["message"]
            except (ValueError, KeyError, TypeError):
                pass
            if r.status_code < 400:
                detail = f"expected a redirect, got {r.status_code}"
            raise GraphError(r.status_code, detail, retry_after=wait)
        raise AssertionError("unreachable")

    def _token(self, client, scope=None):
        if scope is None:
            return self._tp(client)
        if not self._scoped:
            raise ValueError(f"a static Graph token can't call {scope}; use authenticate()")
        return self._tp(client, scope)

    def get(
        self, path, params=None, raw=False, client=C_OFFICE, headers=None, scope=None, retry=True
    ):
        """One GET against `client`'s token. `raw=True` returns body text
        (OneNote page content is HTML, not JSON); `raw="bytes"` returns bytes."""
        return self.request(
            "GET",
            path,
            params=params,
            raw=raw,
            client=client,
            headers=headers,
            scope=scope,
            retry=retry,
        )

    def post(self, path, body=None, client=C_OFFICE, headers=None, content_type=None):
        """One POST (JSON body unless bytes/str + `content_type`). Returns the
        response JSON, or None for 202/204."""
        return self.request(
            "POST", path, body, client=client, headers=headers, content_type=content_type
        )

    def patch(self, path, body, client=C_OFFICE, headers=None, content_type=None):
        return self.request(
            "PATCH", path, body, client=client, headers=headers, content_type=content_type
        )

    def put(self, path, body, client=C_OFFICE, headers=None, content_type=None):
        return self.request(
            "PUT", path, body, client=client, headers=headers, content_type=content_type
        )

    def delete(self, path, client=C_OFFICE, headers=None):
        return self.request("DELETE", path, client=client, headers=headers)

    def list(
        self, path, params=None, limit=None, client=C_OFFICE, headers=None, scope=None, retry=True
    ):
        """GET a collection, following @odata.nextLink; returns the combined
        `value` list, capped at `limit` items (and _MAX_PAGES pages)."""
        out = []
        url = self._url(path, params)
        for _ in range(_MAX_PAGES):
            page = self.get(url, client=client, headers=headers, scope=scope, retry=retry)
            out.extend(page.get("value", []))
            if limit and len(out) >= limit:
                return out[:limit]
            url = page.get("@odata.nextLink")
            if not url:
                break
        return out

    # --- $batch & delta -----------------------------------------------------
    def batch(self, requests, client=C_OFFICE):
        """Run many requests in as few round-trips as possible (`$batch`, 20
        per call, up to 4 calls at once). Each request is a path string (a GET)
        or a dict with `url` and optional `method`, `body`, `headers`. Returns
        one entry per request, in order: the response body (None if empty), or a
        `GraphError` instance (not raised) for that request's failure.
        Throttled (429/503) sub-requests are retried with backoff."""
        reqs = []
        for i, r in enumerate(requests):
            r = {"url": r} if isinstance(r, str) else dict(r)
            r["id"] = str(i)
            r.setdefault("method", "GET")
            r["url"] = "/" + r["url"].removeprefix(self.base).lstrip("/")
            if "body" in r:
                r.setdefault("headers", {}).setdefault("Content-Type", "application/json")
            reqs.append(r)
        results = {}
        pending = reqs
        waited = 0.0
        for attempt in range(_RETRIES + 1):
            throttled, wait = [], 0.0
            chunks = [pending[i : i + _BATCH_MAX] for i in range(0, len(pending), _BATCH_MAX)]
            final = attempt == _RETRIES
            for chunk, resp in zip(chunks, self._post_batches(chunks, client), strict=True):
                for r in (resp or {}).get("responses", []):
                    status = r.get("status", 0)
                    if status in _RETRY_STATUS and not final:
                        throttled.append(next(q for q in chunk if q["id"] == r["id"]))
                        wait = max(wait, _retry_after(r.get("headers"), attempt))
                    elif status >= 400:
                        err = (r.get("body") or {}).get("error", {})
                        results[r["id"]] = GraphError(status, err.get("message", err))
                    else:
                        results[r["id"]] = r.get("body")
            if not throttled:
                break
            if waited + wait > _RETRY_BUDGET:
                for q in throttled:
                    results[q["id"]] = GraphError(429, "throttled", retry_after=wait)
                break
            time.sleep(wait)
            waited += wait
            pending = throttled
        return [results.get(r["id"]) for r in reqs]

    def _post_batches(self, chunks, client):
        if len(chunks) < 2:
            return [self.post("$batch", {"requests": c}, client=client) for c in chunks]
        self._token(client)  # mint once up front so the threads share one cached token
        with ThreadPoolExecutor(min(_BATCH_WORKERS, len(chunks))) as pool:
            return list(
                pool.map(lambda c: self.post("$batch", {"requests": c}, client=client), chunks)
            )

    def delta(self, path, delta_link=None, params=None, client=C_OFFICE):
        """Incremental sync of a collection. First call (no `delta_link`) GETs
        `{path}/delta` and returns every item; pass the returned `delta_link`
        next time to get only what changed since. Returns
        `(items, delta_link)`."""
        url = delta_link or self._url(path.rstrip("/") + "/delta", params)
        items = []
        for _ in range(_MAX_PAGES * 4):
            page = self.get(url, client=client)
            items.extend(page.get("value", []))
            if page.get("@odata.deltaLink"):
                return items, page["@odata.deltaLink"]
            url = page.get("@odata.nextLink")
            if not url:
                break
        return items, None

    # --- identity -----------------------------------------------------------
    def _directory(self, call, *args, **kw):
        """Run a directory read, moving to the next family client straight away if one is
        throttled or refused; the last one waits out Retry-After as usual. A static token has only
        one identity to try."""
        clients = _DIRECTORY_CLIENTS if self._scoped else (C_OFFICE,)
        for client in clients[:-1]:
            try:
                return call(*args, client=client, retry=False, **kw)
            except (GraphError, TokenError) as e:
                if isinstance(e, GraphError) and e.status not in _ROTATE_ON:
                    raise
                log.debug("directory read moving on from client %s: %s", client, e)
        return call(*args, client=clients[-1], **kw)

    def me(self):
        return self._directory(self.get, "me")

    def _claims(self):
        """The signed-in user's claims, read from the token without a request. {} if the token
        isn't a readable JWT (only used for ids, so the signature doesn't matter here)."""
        token = self._tp(C_OFFICE)
        try:
            payload = token.split(".")[1]
            return json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
        except (IndexError, ValueError, AttributeError):
            return {}

    def _me_id(self):
        if "me" not in self._ids:
            self._ids["me"] = (
                self._claims().get("oid")
                or self._directory(self.get, "me", {"$select": "id"})["id"]
            )
        return self._ids["me"]

    def _tenant_id(self):
        if "tenant" not in self._ids:
            self._ids["tenant"] = (
                self._claims().get("tid")
                or self._directory(self.list, "organization", {"$select": "id"})[0]["id"]
            )
        return self._ids["tenant"]

    def _upn(self):
        if "upn" not in self._ids:
            claims = self._claims()
            self._ids["upn"] = (
                claims.get("upn")
                or claims.get("preferred_username")
                or self._directory(self.get, "me", {"$select": "userPrincipalName"})[
                    "userPrincipalName"
                ]
            )
        return self._ids["upn"]
