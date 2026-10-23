"""Headless Microsoft 365 sign-in that needs no app registration.

One sign-in (via `entrance`) against the pre-consented OfficeHome client yields
a FOCI **family refresh token**; a broker then mints a Graph token for whichever
first-party client a given call needs — all without further sign-ins. See the
README for the token map and the (accounts-you-own) disclaimer.
"""

import base64
import contextlib
import hashlib
import json
import os
import secrets
import tempfile
import time
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from ._graph import C_OFFICE, OutNote

_AUTHORIZE = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
_TOKEN = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

# The sign-in client: OfficeHome is pre-consented and issues a FOCI (family=1)
# refresh token that the broker trades for the other family clients' tokens.
OFFICEHOME = "4765445b-32c6-49b0-83e6-1d93765276ca"
# Native-client redirect: redeems server-side (no SPA/Origin dance). entrance
# intercepts the code here; it is never actually delivered.
_REDIRECT = "https://login.microsoftonline.com/common/oauth2/nativeclient"
_SCOPES = "https://graph.microsoft.com/.default offline_access openid profile"
_TIMEOUT = 30


class AuthError(RuntimeError):
    """The token endpoint refused a refresh. `error` is the OAuth error code,
    e.g. `invalid_grant` when the refresh token expired or was revoked."""

    def __init__(self, error, description):
        self.error = error
        super().__init__(f"{error}: {description}")


def _write_secret(path, text):
    """Atomically replace `path` with `text`, mode 0600: write a temp file in
    the same directory, fsync, then os.replace - so a crash mid-write can never
    leave a truncated token behind."""
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(os.path.abspath(path)), prefix=".frt-")
    try:
        with os.fdopen(fd, "w") as f:  # mkstemp creates the file 0600
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except BaseException:
        with contextlib.suppress(OSError):
            os.unlink(tmp)
        raise


def _pkce():
    v = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    c = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=").decode()
    return v, c


def _post_token(url, data):
    req = Request(
        url,
        data=urlencode(data).encode(),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urlopen(req, timeout=_TIMEOUT) as r:  # noqa: S310 - fixed https host
            return json.loads(r.read())
    except Exception as e:  # noqa: BLE001 - surface the AADSTS error body
        body = getattr(e, "read", lambda: b"")()
        try:
            return json.loads(body)
        except ValueError:
            raise RuntimeError(f"Token endpoint error: {e}") from None


class _Broker:
    """Holds a FOCI family refresh token and mints/caches a Graph access token
    per first-party client on demand. Callable as `broker(client_id) -> token`.

    With `frt_path`, the token is loaded from that file when `refresh_token` is
    not given, and every rotated token is written back to it atomically."""

    def __init__(self, refresh_token=None, tenant="common", frt_path=None):
        self._path = frt_path
        loaded = not refresh_token and frt_path and os.path.exists(frt_path)
        if loaded:
            with open(frt_path) as f:
                refresh_token = f.read().strip()
        if not refresh_token:
            raise ValueError("refresh_token is required (or a frt_path that holds one)")
        self._frt = refresh_token
        self._url = _TOKEN.format(tenant=tenant)
        self._cache = {}  # client_id -> (access_token, expiry_epoch)
        if not loaded:
            self._save()

    @property
    def refresh_token(self):
        return self._frt

    def _save(self):
        if self._path:
            _write_secret(self._path, self._frt)

    def __call__(self, client_id):
        tok, exp = self._cache.get(client_id, (None, 0))
        if tok and exp - 60 > time.time():
            return tok
        r = _post_token(
            self._url,
            {
                "grant_type": "refresh_token",
                "refresh_token": self._frt,
                "client_id": client_id,
                "scope": "https://graph.microsoft.com/.default",
            },
        )
        if not r.get("access_token"):
            if r.get("error") == "invalid_grant" and self._path:
                # expired or revoked: drop the dead token so the next run logs in
                with contextlib.suppress(OSError):
                    os.unlink(self._path)
            raise AuthError(
                r.get("error", "no_token"),
                f"could not get a token for client {client_id}: {r.get('error_description') or r}",
            )
        if r.get("refresh_token", self._frt) != self._frt:  # FRT rotates on use
            self._frt = r["refresh_token"]
            self._save()
        self._cache[client_id] = (r["access_token"], time.time() + r.get("expires_in", 3600))
        return r["access_token"]


def authenticate(
    username=None,
    password=None,
    *,
    tenant="common",
    totp_secret=None,
    totp_code=None,
    cookies=".ms_session.json",
    frt_store=".ms_frt",
    read_only=False,
    **login_kwargs,
):
    """Sign in and return an `OutNote` client wired to a token broker.

    No app registration or admin consent needed. `username`/`password` are
    optional once `cookies` holds a live session (silent reuse). MFA: pass
    `totp_secret` (Base32) or `totp_code`. Extra kwargs pass through to
    `entrance.login`.

    `frt_store` (a file path, `None` to disable) persists the FOCI refresh
    token: when it holds a working token, no sign-in happens at all; when the
    token is missing, expired or revoked, this falls back to a full sign-in and
    stores the new token. Treat the file like a password. `read_only=True`
    makes every write method raise.
    """
    if frt_store and os.path.exists(frt_store):
        try:
            broker = _Broker(tenant=tenant, frt_path=frt_store)
            broker(C_OFFICE)  # proves the stored token still works
            return OutNote(broker, read_only=read_only)
        except (AuthError, ValueError):
            pass  # unusable store: sign in again below

    from entrance import login  # deferred: only the sign-in path needs entrance

    verifier, challenge = _pkce()
    authorize_url = (
        _AUTHORIZE.format(tenant=tenant)
        + "?"
        + urlencode(
            {
                "client_id": OFFICEHOME,
                "response_type": "code",
                "redirect_uri": _REDIRECT,
                "response_mode": "query",
                "scope": _SCOPES,
                "state": secrets.token_urlsafe(16),
                "code_challenge": challenge,
                "code_challenge_method": "S256",
            }
        )
    )

    res = login(
        authorize_url,
        username=username,
        password=password,
        totp_secret=totp_secret,
        totp_code=totp_code,
        code_verifier=verifier,
        redirect_uri=_REDIRECT,
        cookies=cookies,
        **login_kwargs,
    )
    if not res.get("code"):
        raise RuntimeError("No authorization code intercepted; keys: " + ", ".join(res))

    tokens = _post_token(
        _TOKEN.format(tenant=tenant),
        {
            "grant_type": "authorization_code",
            "code": res["code"],
            "client_id": OFFICEHOME,
            "redirect_uri": _REDIRECT,
            "code_verifier": verifier,
        },
    )
    frt = tokens.get("refresh_token")
    if not frt:
        raise RuntimeError(
            f"No refresh token from sign-in: {tokens.get('error_description') or tokens}"
        )
    return OutNote(_Broker(frt, tenant, frt_path=frt_store), read_only=read_only)
