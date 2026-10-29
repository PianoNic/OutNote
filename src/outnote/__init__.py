"""OutNote — Microsoft 365 over Graph, headless, no app registration.

    from outnote import authenticate
    note = authenticate("you@org.com", "password", totp_secret="BASE32")

    note.emails(); note.send_mail(to, subject, body)        # mail
    note.events(); note.free_busy([...], start, end)        # calendar
    note.chats(); note.chat_messages(id); note.send_to(user, "hi")   # Teams chats
    note.teams(); note.channel_messages(t, c)               # Teams channels
    note.files(); note.download(id); note.upload(path, "dir/name")   # OneDrive
    note.sites()                                            # SharePoint
    note.notebooks(); note.pages(sec); note.page_text(id)   # OneNote
    note.search("budget", entity="driveItem")               # unified search

One sign-in yields a FOCI family refresh token (saved to `.ms_frt`, so later
runs skip the sign-in); a broker mints the right first-party client's token per
call. Already have a token? `OutNote(token)`.
"""

from entrance import (
    APPROVED,
    DENIED,
    NUMBER_MATCHING,
    OTP,
    PENDING,
    PUSH,
    TIMED_OUT,
    Blocked,
    FederatedAccount,
    InvalidCredentials,
    MfaCancelled,
    MfaPrompt,
    MfaSwitch,
    OtpRejected,
    PushError,
    PushEvent,
    Secret,
    Unsupported,
    UnsupportedMfa,
    push_login,
)

from ._auth import OFFICEHOME, authenticate
from ._broker import TokenError
from ._graph import C_EDU, C_OFFICE, C_ONENOTE, C_SITES, GraphError, OutNote
from ._text import extract_text, html_to_text, message_text

__all__ = [
    "OutNote",
    "authenticate",
    "GraphError",
    "TokenError",
    "extract_text",
    "html_to_text",
    "message_text",
    "OFFICEHOME",
    "C_OFFICE",
    "C_ONENOTE",
    "C_SITES",
    "C_EDU",
    "push_login",
    "PushEvent",
    "PushError",
    "NUMBER_MATCHING",
    "PENDING",
    "APPROVED",
    "DENIED",
    "TIMED_OUT",
    "MfaPrompt",
    "MfaSwitch",
    "MfaCancelled",
    "Secret",
    "PUSH",
    "OTP",
    "InvalidCredentials",
    "Unsupported",
    "FederatedAccount",
    "UnsupportedMfa",
    "Blocked",
    "OtpRejected",
]
