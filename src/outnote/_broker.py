"""FOCI token broker: trades one family refresh token for a Graph access token
per first-party client, and (optionally) persists the rotated refresh token so
later runs skip the sign-in entirely."""

import json
import os
import tempfile
import time
from urllib.parse import urlencode
from urllib.request import Request, urlopen

_TOKEN = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"


class TokenError(RuntimeError):
    """The token endpoint refused a grant. `error` is the OAuth error code
    (e.g. `invalid_grant` for an expired or revoked refresh token)."""

    def __init__(self, error, description):
        self.error = error
        super().__init__(f"{error}: {description}")


def _post_token(url, data):
    req = Request(
        url,
        data=urlencode(data).encode(),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urlopen(req, timeout=30) as r:  # noqa: S310 - fixed https host
            return json.loads(r.read())
    except Exception as e:  # noqa: BLE001 - surface the AADSTS error body
        body = getattr(e, "read", lambda: b"")()
        try:
            return json.loads(body)
        except ValueError:
            raise RuntimeError(f"Token endpoint error: {e}") from None


def read_secret(path):
    """The stripped contents of `path`, or None if it is missing/empty."""
    try:
        with open(path, encoding="utf-8") as f:
            return f.read().strip() or None
    except FileNotFoundError:
        return None


def write_secret(path, value):
    """Atomically write `value` to `path` with owner-only permissions: a temp
    file in the same directory is fully written, then `os.replace`d over the
    target, so a crash mid-write never truncates the stored token."""
    d = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=d, prefix=".outnote-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(value)
        os.chmod(tmp, 0o600)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


class Broker:
    """Holds a FOCI family refresh token and mints/caches a Graph access token
    per first-party client on demand. Callable as `broker(client_id) -> token`,
    or `broker(client_id, scope)` for a non-Graph resource (e.g. the Teams
    Assignments API).

    With `store`, the refresh token is written back (atomically, 0600) every
    time it rotates."""

    def __init__(self, refresh_token, tenant="common", store=None):
        self._frt = refresh_token
        self._url = _TOKEN.format(tenant=tenant)
        self._store = store
        self._cache = {}  # (client_id, scope) -> (access_token, expiry_epoch)

    @property
    def refresh_token(self):
        return self._frt

    def __call__(self, client_id, scope=GRAPH_SCOPE):
        tok, exp = self._cache.get((client_id, scope), (None, 0))
        if tok and exp - 60 > time.time():
            return tok
        r = _post_token(
            self._url,
            {
                "grant_type": "refresh_token",
                "refresh_token": self._frt,
                "client_id": client_id,
                "scope": scope,
            },
        )
        if not r.get("access_token"):
            raise TokenError(
                r.get("error", "no_token"),
                f"client {client_id}: {r.get('error_description') or r}",
            )
        new = r.get("refresh_token")
        if new and new != self._frt:  # FRT rotates on use
            self._frt = new
            if self._store:
                write_secret(self._store, new)
        self._cache[(client_id, scope)] = (
            r["access_token"],
            time.time() + r.get("expires_in", 3600),
        )
        return r["access_token"]
