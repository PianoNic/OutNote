"""Headless Microsoft 365 sign-in - a thin wrapper over `entrance.foci_login`.

All the auth mechanics (PKCE sign-in, device code, the FOCI family refresh
token, the token broker and its persistence) live in `entrance`; OutNote only
supplies the Microsoft 365 specifics - the sign-in client and the Office client
used to validate a stored token.
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
    device=False,
    open_browser=False,
    on_prompt=None,
    push=False,
    on_push=None,
    push_timeout=120,
    push_interval=3,
    **login_kwargs,
):
    """Sign in and return an `OutNote` client wired to a token broker.

    No app registration or admin consent needed. If `frt_store` holds a live
    refresh token, no sign-in happens at all. Otherwise pick how to sign in:
      * default - headless (password + TOTP, or a seeded cookie jar).
      * `device=True` - the device code flow: a short code shown on the terminal
        that you enter at a URL, no browser redirect or copy-paste (the Office
        client is used, which supports it and brokers to every M365 service).
      * `interactive=True` / `"always"` - finish in a browser.
      * `push=True` - headless password sign-in finished with a number-matching
        Authenticator push: `on_push(PushEvent)` receives the two-digit number to
        tap and the pending/approved/denied/timed-out progress, so your own UI
        can show it. `push_timeout` bounds the wait, `push_interval` paces the
        poll. Raises `outnote.PushError` if rejected or unanswered.
    MFA: pass `totp_secret` (Base32) or `totp_code`. Extra kwargs pass through
    to `entrance.login`.
    """
    broker = foci_login(
        username,
        password,
        client_id=C_OFFICE if device else OFFICEHOME,
        tenant=tenant,
        totp_secret=totp_secret,
        totp_code=totp_code,
        cookies=cookies,
        frt_store=frt_store,
        probe_client=C_OFFICE,
        device=device,
        open_browser=open_browser,
        on_prompt=on_prompt,
        push=push,
        on_push=on_push,
        push_timeout=push_timeout,
        push_interval=push_interval,
        **login_kwargs,
    )
    return OutNote(broker)
