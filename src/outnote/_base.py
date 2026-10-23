"""Graph transport: GET/POST/PATCH/PUT/DELETE + Bearer + JSON over stdlib
urllib, with 429/503 backoff, nextLink paging, $batch and delta."""

import json
import time
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

GRAPH = "https://graph.microsoft.com/v1.0"
_MAX_PAGES = 50  # ponytail: hard stop so follow-the-nextLink can't loop forever
_TIMEOUT = 30  # seconds per request, so a hung connection can't block forever
_RETRIES = 4  # retries on 429/502/503/504, honoring Retry-After
_RETRY_STATUS = (429, 502, 503, 504)
_BATCH_MAX = 20  # Graph's per-$batch request limit

# FOCI family clients, each pre-consented for the scopes that service needs.
C_OFFICE = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # mail, calendar, chat, channels, files, teams, contacts, tasks
C_ONENOTE = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"  # OneNote (Notes.ReadWrite.All)
C_SITES = "4765445b-32c6-49b0-83e6-1d93765276ca"  # SharePoint sites (Sites.Read.All)

# Characters left as-is when cleaning a caller's URL: everything URL-structural
# plus `%`, so an already-encoded URL (an @odata.nextLink) passes unchanged
# while spaces, quotes and non-ASCII are percent-encoded.
_URL_SAFE = ":/?#[]@!$&'()*+,;=%~-._"


class GraphError(RuntimeError):
    def __init__(self, status, message):
        self.status = status
        super().__init__(f"Graph {status}: {message}")


def _retry_after(headers, attempt):
    ra = (headers or {}).get("Retry-After")
    return float(ra) if ra and str(ra).replace(".", "", 1).isdigit() else 2**attempt


def fetch(url, timeout=_TIMEOUT):
    """GET a pre-authenticated URL (a drive item's downloadUrl) as bytes."""
    with urlopen(Request(url), timeout=timeout) as r:  # noqa: S310 - Graph-issued URL
        return r.read()


class Transport:
    def __init__(self, token, base=GRAPH):
        """`token` is either an access-token string (used for every call) or a
        callable `client_id -> access_token` (the broker from `authenticate`,
        which hands each service the right client's token)."""
        if not token:
            raise ValueError("token is required")
        self._tp = token if callable(token) else (lambda _client: token)
        self._scoped = callable(token)  # a broker can mint non-Graph tokens
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

    def _send(self, req, raw=False):
        """Send `req`; returns parsed JSON, or text (`raw=True`), or bytes
        (`raw="bytes"`). Empty bodies (202/204) come back as None."""
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
                if e.code in _RETRY_STATUS and attempt < _RETRIES:
                    time.sleep(_retry_after(e.headers, attempt))
                    continue
                detail = e.read().decode("utf-8", "replace")
                try:
                    detail = json.loads(detail)["error"]["message"]
                except (ValueError, KeyError, TypeError):
                    pass
                raise GraphError(e.code, detail) from None
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
        return self._send(req, raw)

    def _token(self, client, scope=None):
        if scope is None:
            return self._tp(client)
        if not self._scoped:
            raise ValueError(f"a static Graph token can't call {scope}; use authenticate()")
        return self._tp(client, scope)

    def get(self, path, params=None, raw=False, client=C_OFFICE, headers=None, scope=None):
        """One GET against `client`'s token. `raw=True` returns body text
        (OneNote page content is HTML, not JSON); `raw="bytes"` returns bytes."""
        return self.request(
            "GET", path, params=params, raw=raw, client=client, headers=headers, scope=scope
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

    def list(self, path, params=None, limit=None, client=C_OFFICE, headers=None, scope=None):
        """GET a collection, following @odata.nextLink; returns the combined
        `value` list, capped at `limit` items (and _MAX_PAGES pages)."""
        out = []
        url = self._url(path, params)
        for _ in range(_MAX_PAGES):
            page = self.get(url, client=client, headers=headers, scope=scope)
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
        per call). Each request is a path string (a GET) or a dict with
        `url` and optional `method`, `body`, `headers`. Returns one entry per
        request, in order: the response body (None if empty), or a
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
        for attempt in range(_RETRIES + 1):
            throttled, wait = [], 0.0
            for i in range(0, len(pending), _BATCH_MAX):
                chunk = pending[i : i + _BATCH_MAX]
                resp = self.post("$batch", {"requests": chunk}, client=client) or {}
                for r in resp.get("responses", []):
                    status = r.get("status", 0)
                    if status in _RETRY_STATUS and attempt < _RETRIES:
                        throttled.append(next(q for q in chunk if q["id"] == r["id"]))
                        wait = max(wait, _retry_after(r.get("headers"), attempt))
                    elif status >= 400:
                        err = (r.get("body") or {}).get("error", {})
                        results[r["id"]] = GraphError(status, err.get("message", err))
                    else:
                        results[r["id"]] = r.get("body")
            if not throttled:
                break
            time.sleep(wait)
            pending = throttled
        return [results.get(r["id"]) for r in reqs]

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
    def me(self):
        return self.get("me")

    def _me_id(self):
        if "me" not in self._ids:
            self._ids["me"] = self.get("me", {"$select": "id"})["id"]
        return self._ids["me"]

    def _tenant_id(self):
        if "tenant" not in self._ids:
            self._ids["tenant"] = self.list("organization", {"$select": "id"})[0]["id"]
        return self._ids["tenant"]
