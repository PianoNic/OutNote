"""Headless Microsoft 365 sign-in - a thin wrapper over `entrance.foci_login`.

All the auth mechanics (PKCE sign-in, the FOCI family refresh token, the token
broker and its persistence) live in `entrance`; OutNote only supplies the
Microsoft 365 specifics - the OfficeHome sign-in client (entrance's default)
and the Office client used to validate a stored token.
"""

from entrance import OFFICEHOME, foci_login

from ._base import C_OFFICE
from ._graph import OutNote

__all__ = ["authenticate", "OFFICEHOME"]


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
    MFA: pass `totp_secret` (Base32) or `totp_code`; or `interactive=True` to
    finish in a browser. Extra kwargs pass through to `entrance.login`.
    """
    broker = foci_login(
        username,
        password,
        tenant=tenant,
        totp_secret=totp_secret,
        totp_code=totp_code,
        cookies=cookies,
        frt_store=frt_store,
        probe_client=C_OFFICE,
        **login_kwargs,
    )
    return OutNote(broker)
