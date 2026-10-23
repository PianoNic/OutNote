"""Headless Microsoft 365 sign-in that needs no app registration.

One sign-in (via `entrance`) against the pre-consented OfficeHome client yields
a FOCI **family refresh token**; a broker then mints a Graph token for whichever
first-party client a given call needs — all without further sign-ins. The
refresh token is persisted (atomically, 0600) so later runs skip the sign-in.
See the README for the token map and the (accounts-you-own) disclaimer.
"""

import base64
import hashlib
import os
import secrets
from urllib.parse import urlencode

from ._base import C_OFFICE
from ._broker import _TOKEN, Broker, TokenError, _post_token, read_secret, write_secret
from ._graph import OutNote

_AUTHORIZE = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"

# The sign-in client: OfficeHome is pre-consented and issues a FOCI (family=1)
# refresh token that the broker trades for the other family clients' tokens.
OFFICEHOME = "4765445b-32c6-49b0-83e6-1d93765276ca"
# Native-client redirect: redeems server-side (no SPA/Origin dance). entrance
# intercepts the code here; it is never actually delivered.
_REDIRECT = "https://login.microsoftonline.com/common/oauth2/nativeclient"
_SCOPES = "https://graph.microsoft.com/.default offline_access openid profile"

_Broker = Broker  # backwards-compatible private name


def _pkce():
    v = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    c = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=").decode()
    return v, c


def _from_store(frt_store, tenant):
    """An `OutNote` built from the saved refresh token, or None if there is no
    usable one (missing, expired or revoked — a dead token is deleted)."""
    frt = read_secret(frt_store)
    if not frt:
        return None
    broker = Broker(frt, tenant, store=frt_store)
    try:
        broker(C_OFFICE)  # prove the token is alive before skipping sign-in
    except TokenError as e:
        if e.error in ("invalid_grant", "interaction_required"):
            os.remove(frt_store)
            return None
        raise
    return OutNote(broker)


def authenticate(
    username=None,
    password=None,
    *,
    tenant="common",
    totp_secret=None,
    totp_code=None,
    cookies=".ms_session.json",
    frt_store=".ms_frt",
    **login_kwargs,
):
    """Sign in and return an `OutNote` client wired to a token broker.

    No app registration or admin consent needed. If `frt_store` holds a live
    refresh token, no sign-in happens at all; otherwise `entrance` signs in and
    the new token is saved there (pass `frt_store=None` to disable).
    `username`/`password` are optional once `cookies` holds a live session.
    MFA: pass `totp_secret` (Base32) or `totp_code`. Extra kwargs pass through
    to `entrance.login`.
    """
    if frt_store:
        note = _from_store(frt_store, tenant)
        if note:
            return note

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
    if frt_store:
        write_secret(frt_store, frt)
    return OutNote(Broker(frt, tenant, store=frt_store))
