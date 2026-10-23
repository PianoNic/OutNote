"""Microsoft Graph client for mail, calendar, Teams, files, SharePoint,
OneNote, people, tasks, Planner and Education.

Each method routes to the pre-consented first-party client whose token carries
the scope it needs (see `authenticate` / the README token map). Graph is just
HTTP + Bearer + JSON, so this leans on stdlib urllib. Methods that change the
account (send, create, update, delete) say so in their docstrings.
"""

from ._base import (
    _MAX_PAGES,
    _RETRIES,
    _TIMEOUT,
    C_OFFICE,
    C_ONENOTE,
    C_SITES,
    GRAPH,
    GraphError,
)
from ._calendar import CalendarMixin
from ._files import FilesMixin
from ._mail import MailMixin
from ._people import ProfileMixin
from ._teams import TeamsMixin
from ._text import extract_text, html_to_text, message_text
from ._work import _SEARCH_CLIENT, C_EDU, EducationMixin, SearchMixin, TasksMixin

__all__ = [
    "GRAPH",
    "C_OFFICE",
    "C_ONENOTE",
    "C_SITES",
    "C_EDU",
    "GraphError",
    "OutNote",
    "_MAX_PAGES",
    "_RETRIES",
    "_TIMEOUT",
    "_SEARCH_CLIENT",
]


class OutNote(
    MailMixin,
    CalendarMixin,
    TeamsMixin,
    FilesMixin,
    ProfileMixin,
    TasksMixin,
    EducationMixin,
    SearchMixin,
):
    """The whole client. Build it with `authenticate(...)`, from a saved
    refresh token with `OutNote.from_refresh_token(...)`, or from a single
    access token with `OutNote(token)`."""

    @classmethod
    def from_refresh_token(cls, refresh_token, tenant="common", store=None):
        """A client brokered from a FOCI family refresh token you already
        have — no sign-in. With `store`, the rotated token is saved there."""
        from ._broker import Broker

        return cls(Broker(refresh_token, tenant, store=store))

    # plain-text helpers, also importable from `outnote`
    message_text = staticmethod(message_text)
    html_to_text = staticmethod(html_to_text)
    extract_text = staticmethod(extract_text)
