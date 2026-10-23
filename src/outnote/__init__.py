"""OutNote — read Microsoft 365 over Graph, headless, no app registration.

    from outnote import authenticate
    note = authenticate("you@org.com", "password", totp_secret="BASE32")

    note.emails()                      # mail
    note.events()                      # calendar
    note.chats(); note.chat_messages(id)        # Teams chats
    note.teams(); note.channel_messages(t, c)   # Teams channels
    note.channel_files(t, c)                    # files posted in a channel
    note.files()                       # OneDrive
    note.sites()                       # SharePoint
    note.notebooks(); note.pages(sec); note.page_content(id)   # OneNote

One sign-in yields a FOCI family refresh token (saved to `.ms_frt`, so later
runs skip the sign-in); a broker mints the right first-party client's token per
call. Already have a token? `OutNote(token)` or `OutNote.from_refresh_token(frt)`.
"""

from ._auth import OFFICEHOME, AuthError, authenticate
from ._extract import extract_text, html_to_text, vtt_to_text
from ._graph import C_OFFICE, C_ONENOTE, C_SITES, GraphError, OutNote

__all__ = [
    "OutNote",
    "authenticate",
    "GraphError",
    "AuthError",
    "OFFICEHOME",
    "C_OFFICE",
    "C_ONENOTE",
    "C_SITES",
    "extract_text",
    "html_to_text",
    "vtt_to_text",
]
