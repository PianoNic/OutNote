"""Microsoft Graph client for mail, calendar, chats, channels, files,
SharePoint and OneNote.

Each method routes to the pre-consented first-party client whose token carries
the scope it needs (see `authenticate` / the README token map). Graph is just
GET + Bearer + JSON, so this leans on stdlib urllib.
"""

import json
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

GRAPH = "https://graph.microsoft.com/v1.0"
_MAX_PAGES = 50  # ponytail: hard stop so follow-the-nextLink can't loop forever

# FOCI family clients, each pre-consented for the scopes that service needs.
C_OFFICE = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # mail, calendar, chat, channels, files, teams, contacts, tasks
C_ONENOTE = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"  # OneNote (Notes.ReadWrite.All)
C_SITES = "4765445b-32c6-49b0-83e6-1d93765276ca"  # SharePoint sites (Sites.Read.All)


class GraphError(RuntimeError):
    def __init__(self, status, message):
        self.status = status
        super().__init__(f"Graph {status}: {message}")


class OutNote:
    def __init__(self, token, base=GRAPH):
        """`token` is either an access-token string (used for every call) or a
        callable `client_id -> access_token` (the broker from `authenticate`,
        which hands each service the right client's token)."""
        if not token:
            raise ValueError("token is required")
        self._tp = token if callable(token) else (lambda _client: token)
        self.token = token if isinstance(token, str) else None
        self.base = base.rstrip("/")

    # --- transport ----------------------------------------------------------
    def _url(self, path, params=None):
        url = path if path.startswith("http") else self.base + "/" + path.lstrip("/")
        if params:
            q = urlencode({k: v for k, v in params.items() if v is not None})
            if q:
                url += ("&" if "?" in url else "?") + q
        return url

    def get(self, path, params=None, raw=False, client=C_OFFICE):
        """One GET against `client`'s token. `raw=True` returns body text
        (OneNote page content is HTML, not JSON)."""
        req = Request(
            self._url(path, params),
            headers={"Authorization": f"Bearer {self._tp(client)}", "Accept": "application/json"},
        )
        try:
            with urlopen(req) as r:  # noqa: S310 - fixed https host
                body = r.read()
        except HTTPError as e:
            detail = e.read().decode("utf-8", "replace")
            try:
                detail = json.loads(detail)["error"]["message"]
            except (ValueError, KeyError, TypeError):
                pass
            raise GraphError(e.code, detail) from None
        return body.decode("utf-8") if raw else json.loads(body)

    def list(self, path, params=None, limit=None, client=C_OFFICE):
        """GET a collection, following @odata.nextLink; returns the combined
        `value` list, capped at `limit` items (and _MAX_PAGES pages)."""
        out = []
        url = self._url(path, params)
        for _ in range(_MAX_PAGES):
            page = self.get(url, client=client)
            out.extend(page.get("value", []))
            if limit and len(out) >= limit:
                return out[:limit]
            url = page.get("@odata.nextLink")
            if not url:
                break
        return out

    # --- mail ---------------------------------------------------------------
    def emails(self, folder="inbox", top=25, search=None, select=None):
        params = {"$top": top, "$select": select}
        if search:
            params["$search"] = f'"{search}"'
        else:
            params["$orderby"] = "receivedDateTime desc"
        return self.list(f"me/mailFolders/{folder}/messages", params, limit=top)

    def email(self, message_id):
        return self.get(f"me/messages/{message_id}")

    # --- calendar -----------------------------------------------------------
    def events(self, start=None, end=None, top=25):
        if start and end:
            params = {
                "startDateTime": start,
                "endDateTime": end,
                "$top": top,
                "$orderby": "start/dateTime",
            }
            return self.list("me/calendarView", params, limit=top)
        return self.list("me/events", {"$top": top, "$orderby": "start/dateTime"}, limit=top)

    # --- teams: chats -------------------------------------------------------
    def chats(self, top=25):
        """1:1 and group chats."""
        return self.list("me/chats", {"$top": top}, limit=top)

    def chat_messages(self, chat_id, top=30):
        return self.list(f"me/chats/{chat_id}/messages", {"$top": top}, limit=top)

    # --- teams: teams & channels --------------------------------------------
    def teams(self):
        return self.list("me/joinedTeams")

    def channels(self, team_id):
        return self.list(f"teams/{team_id}/channels")

    def channel_messages(self, team_id, channel_id, top=30):
        return self.list(
            f"teams/{team_id}/channels/{channel_id}/messages", {"$top": top}, limit=top
        )

    def channel_files(self, team_id, channel_id):
        """Files posted in a channel (the channel's SharePoint folder)."""
        ff = self.get(f"teams/{team_id}/channels/{channel_id}/filesFolder")
        drive, item = ff["parentReference"]["driveId"], ff["id"]
        return self.list(f"drives/{drive}/items/{item}/children")

    # --- files / onedrive / sharepoint --------------------------------------
    def files(self, path="root", top=100):
        """OneDrive items. `path` is 'root' or 'root:/Folder/Sub:'."""
        return self.list(f"me/drive/{path}/children", {"$top": top}, limit=top)

    def download_url(self, item_id, drive_id=None):
        """A direct, pre-authed download URL for a drive item (works for OneDrive
        and channel files — pass the item's drive_id for the latter)."""
        base = f"drives/{drive_id}/items/{item_id}" if drive_id else f"me/drive/items/{item_id}"
        return self.get(base).get("@microsoft.graph.downloadUrl")

    def sites(self, search="*"):
        return self.list(f"sites?search={search}", client=C_SITES)

    def site_drive_items(self, site_id, path="root", top=100):
        return self.list(
            f"sites/{site_id}/drive/{path}/children", {"$top": top}, limit=top, client=C_SITES
        )

    # --- onenote ------------------------------------------------------------
    def notebooks(self):
        return self.list("me/onenote/notebooks", client=C_ONENOTE)

    def sections(self, notebook_id=None):
        path = (
            f"me/onenote/notebooks/{notebook_id}/sections" if notebook_id else "me/onenote/sections"
        )
        return self.list(path, client=C_ONENOTE)

    def pages(self, section_id, top=50):
        """Pages in a section. (List per-section — the all-pages endpoint is
        flaky.)"""
        return self.list(
            f"me/onenote/sections/{section_id}/pages", {"$top": top}, limit=top, client=C_ONENOTE
        )

    def page_content(self, page_id):
        """A OneNote page's body as HTML (string)."""
        return self.get(f"me/onenote/pages/{page_id}/content", raw=True, client=C_ONENOTE)

    # --- people / tasks -----------------------------------------------------
    def contacts(self, top=50):
        return self.list("me/contacts", {"$top": top}, limit=top)

    def task_lists(self):
        return self.list("me/todo/lists")

    def tasks(self, list_id, top=50):
        return self.list(f"me/todo/lists/{list_id}/tasks", {"$top": top}, limit=top)

    def me(self):
        return self.get("me")
