"""Microsoft Graph client for mail, calendar, chats, channels, files,
SharePoint and OneNote.

Each method routes to the pre-consented first-party client whose token carries
the scope it needs (see `authenticate` / the README token map). Graph is just
GET + Bearer + JSON, so this leans on stdlib urllib.
"""

import base64
import json
import os
import re
import time
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from ._extract import html_images, html_to_text

GRAPH = "https://graph.microsoft.com/v1.0"
_MAX_PAGES = 50  # ponytail: hard stop so follow-the-nextLink can't loop forever
_TIMEOUT = 30  # seconds per request, so a hung connection can't block forever
_RETRIES = 4  # retries on 429/503, honoring Retry-After
# characters left alone when encoding a URL: reserved delimiters plus `%` (so
# already-encoded nextLinks pass through untouched); a stray `%` that is not
# an escape is handled separately in _url
_URL_SAFE = ":/?#[]@!$&'()*+,;=%~"
_STRAY_PCT = re.compile(r"%(?![0-9A-Fa-f]{2})")

# FOCI family clients, each pre-consented for the scopes that service needs.
C_OFFICE = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # mail, calendar, chat, channels, files, teams, contacts, tasks
C_ONENOTE = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"  # OneNote (Notes.ReadWrite.All)
C_SITES = "4765445b-32c6-49b0-83e6-1d93765276ca"  # SharePoint sites (Sites.Read.All)
# 1fec8e78 is the Microsoft Teams client: the fallback for Teams-only scopes
# (presence, meetings, education) that the Office client may lack in a tenant
C_TEAMS = C_ONENOTE

# which client's token each unified-search entity type needs
_SEARCH_CLIENT = {
    "message": C_OFFICE,
    "event": C_OFFICE,
    "driveItem": C_OFFICE,
    "chatMessage": C_OFFICE,
    "person": C_OFFICE,
    "site": C_SITES,
    "listItem": C_SITES,
}


def _retry_after(header, attempt):
    return float(header) if header and header.replace(".", "", 1).isdigit() else 2**attempt


def _save(data, dest, name):
    """Return `data`, or write it to `dest` (a file, or a directory to put
    `name` in) and return the written path."""
    if dest is None:
        return data
    path = os.path.join(dest, name) if os.path.isdir(dest) else dest
    with open(path, "wb") as f:
        f.write(data)
    return path


class GraphError(RuntimeError):
    def __init__(self, status, message):
        self.status = status
        super().__init__(f"Graph {status}: {message}")


class OutNote:
    def __init__(self, token, base=GRAPH, read_only=False):
        """`token` is either an access-token string (used for every call) or a
        callable `client_id -> access_token` (the broker from `authenticate`,
        which hands each service the right client's token). `read_only=True`
        makes every method that changes the account raise."""
        if not token:
            raise ValueError("token is required")
        self._tp = token if callable(token) else (lambda _client: token)
        self.token = token if isinstance(token, str) else None
        self.base = base.rstrip("/")
        self.read_only = read_only
        self._routes = {}  # area -> client that worked (see _try_clients)
        self._me = None

    @classmethod
    def from_refresh_token(cls, refresh_token=None, tenant="common", frt_path=None, **kw):
        """Build a client from a saved FOCI family refresh token - no sign-in.
        Pass the token, or `frt_path` to load it from (and keep rotating it
        into) a file written by `authenticate(frt_store=...)`."""
        from ._auth import _Broker  # deferred: _auth imports this module

        return cls(_Broker(refresh_token, tenant, frt_path=frt_path), **kw)

    # --- transport ----------------------------------------------------------
    def _url(self, path, params=None):
        url = path if path.startswith("http") else self.base + "/" + path.lstrip("/")
        # a query written into the path (`me/messages?$search="x y ö"`) may carry
        # spaces, quotes or non-ASCII; percent-encode those, leave escapes alone
        url = quote(_STRAY_PCT.sub("%25", url), safe=_URL_SAFE)
        if params:
            q = urlencode({k: v for k, v in params.items() if v is not None})
            if q:
                url += ("&" if "?" in url else "?") + q
        return url

    def _send(self, req, raw=False):
        for attempt in range(_RETRIES + 1):
            try:
                with urlopen(req, timeout=_TIMEOUT) as r:  # noqa: S310 - fixed https host
                    body = r.read()
                if raw == "bytes":
                    return body
                if raw:
                    return body.decode("utf-8")
                # 202/204 (sendMail, PATCH, DELETE, ...) come back empty
                return json.loads(body) if body.strip() else None
            except HTTPError as e:
                # Graph throttles with 429 (and 503); back off and retry.
                if e.code in (429, 503) and attempt < _RETRIES:
                    time.sleep(_retry_after(e.headers.get("Retry-After"), attempt))
                    continue
                detail = e.read().decode("utf-8", "replace")
                try:
                    detail = json.loads(detail)["error"]["message"]
                except (ValueError, KeyError, TypeError):
                    pass
                raise GraphError(e.code, detail) from None

    def request(
        self,
        method,
        path,
        body=None,
        *,
        params=None,
        raw=False,
        client=C_OFFICE,
        data=None,
        content_type="application/json",
        headers=None,
        auth=True,
    ):
        """One Graph request. `body` is sent as JSON; or pass raw `data` bytes
        with their `content_type`. Returns parsed JSON (None for an empty
        202/204 reply); `raw=True` returns text, `raw="bytes"` bytes.
        `auth=False` skips the bearer token (pre-authed download/upload URLs)."""
        if body is not None:
            data = json.dumps(body).encode()
        if data is None and method in ("POST", "PUT", "PATCH"):
            data = b""  # Graph wants a Content-Length even on empty POSTs
        h = {"Accept": "application/json", **(headers or {})}
        if auth:
            h["Authorization"] = f"Bearer {self._tp(client)}"
        if data is not None:
            h.setdefault("Content-Type", content_type)
        req = Request(self._url(path, params), data=data, headers=h, method=method)
        return self._send(req, raw)

    def get(self, path, params=None, raw=False, client=C_OFFICE):
        """One GET against `client`'s token. `raw=True` returns body text
        (OneNote page content is HTML, not JSON), `raw="bytes"` the bytes."""
        return self.request("GET", path, params=params, raw=raw, client=client)

    def post(self, path, body=None, client=C_OFFICE, **kw):
        """One POST of a JSON body against `client`'s token."""
        return self.request("POST", path, body, client=client, **kw)

    def patch(self, path, body=None, client=C_OFFICE, **kw):
        return self.request("PATCH", path, body, client=client, **kw)

    def put(self, path, body=None, client=C_OFFICE, **kw):
        return self.request("PUT", path, body, client=client, **kw)

    def delete(self, path, client=C_OFFICE, **kw):
        return self.request("DELETE", path, client=client, **kw)

    def _fetch(self, url):
        """GET a pre-authenticated URL (e.g. @microsoft.graph.downloadUrl) as
        bytes, without sending our bearer token to a foreign host."""
        return self.request("GET", url, raw="bytes", auth=False)

    def _try_clients(self, area, call, clients=(C_OFFICE, C_TEAMS)):
        """Run `call(client)` with each client in turn until one is allowed
        (401/403 -> try the next), remembering the winner for `area`. Which
        first-party client carries a scope is tenant-dependent."""
        if self.token:  # one static token: no other client to fall back to
            clients = clients[:1]
        known = self._routes.get(area)
        order = [known, *(c for c in clients if c != known)] if known else list(clients)
        for i, client in enumerate(order):
            try:
                out = call(client)
            except GraphError as e:
                if e.status in (401, 403) and i < len(order) - 1:
                    continue
                raise
            self._routes[area] = client
            return out

    def _writable(self):
        if self.read_only:
            raise PermissionError("this OutNote client is read_only; writes are disabled")

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
    def mail_folders(self, parent_id=None):
        """Mail folders (top level, or the child folders of `parent_id`). A
        folder's id - or a well-known name such as 'inbox', 'sentitems',
        'drafts', 'deleteditems', 'archive', 'junkemail' - works as
        `emails(folder=...)`."""
        path = f"me/mailFolders/{parent_id}/childFolders" if parent_id else "me/mailFolders"
        return self.list(path, {"$top": 100})

    def emails(self, folder="inbox", top=25, search=None, select=None):
        """Messages in `folder` (a well-known name or a folder id), newest
        first."""
        params = {"$top": top, "$select": select}
        if search:
            params["$search"] = f'"{search}"'
        else:
            params["$orderby"] = "receivedDateTime desc"
        return self.list(f"me/mailFolders/{folder}/messages", params, limit=top)

    def email(self, message_id):
        return self.get(f"me/messages/{message_id}")

    def attachments(self, message_id):
        """Attachment metadata of a message (name, contentType, size, id,
        isInline, @odata.type); an empty list when it has none."""
        return self.list(
            f"me/messages/{message_id}/attachments",
            {"$select": "id,name,contentType,size,isInline,lastModifiedDateTime"},
        )

    def download_attachment(self, message_id, attachment_id, dest=None):
        """An attachment's bytes, or the path written when `dest` (a file or a
        directory) is given. File attachments are decoded from contentBytes;
        attached Outlook items come back as MIME (.eml). Link ("reference")
        attachments have no bytes and raise ValueError."""
        base = f"me/messages/{message_id}/attachments/{attachment_id}"
        att = self.get(base)
        kind = att.get("@odata.type", "")
        name = att.get("name") or attachment_id
        if att.get("contentBytes") is not None:
            data = base64.b64decode(att["contentBytes"])
        elif kind.endswith("itemAttachment"):
            data = self.get(base + "/$value", raw="bytes")
            name += ".eml"
        else:
            raise ValueError(f"{kind or 'attachment'} has no downloadable content")
        return _save(data, dest, name)

    # --- calendar -----------------------------------------------------------
    def calendars(self):
        return self.list("me/calendars")

    def events(self, start=None, end=None, top=25, calendar_id=None):
        """Events of the default calendar (or `calendar_id`). With `start` and
        `end` (ISO 8601), the expanded occurrences in that window."""
        base = f"me/calendars/{calendar_id}" if calendar_id else "me"
        if start and end:
            params = {
                "startDateTime": start,
                "endDateTime": end,
                "$top": top,
                "$orderby": "start/dateTime",
            }
            return self.list(f"{base}/calendarView", params, limit=top)
        return self.list(f"{base}/events", {"$top": top, "$orderby": "start/dateTime"}, limit=top)

    def event(self, event_id):
        return self.get(f"me/events/{event_id}")

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

    def channel_message_replies(self, team_id, channel_id, message_id, top=50):
        """The replies under a channel post (channel messages are threaded;
        `channel_messages` returns only the top-level posts)."""
        return self.list(
            f"teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            {"$top": min(top, 50)},
            limit=top,
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

    def _item(self, item_id, drive_id=None):
        base = f"drives/{drive_id}/items/{item_id}" if drive_id else f"me/drive/items/{item_id}"
        return self.get(base)

    def download_url(self, item_id, drive_id=None):
        """A direct, pre-authed download URL for a drive item (works for OneDrive
        and channel files — pass the item's drive_id for the latter)."""
        return self._item(item_id, drive_id).get("@microsoft.graph.downloadUrl")

    def download(self, item_id, dest=None, drive_id=None):
        """A drive item's bytes (OneDrive, SharePoint or channel files - pass
        the item's drive_id for the latter two), or the path written when
        `dest` (a file or a directory) is given."""
        item = self._item(item_id, drive_id)
        url = item.get("@microsoft.graph.downloadUrl")
        if not url:
            raise ValueError(f"{item.get('name', item_id)} is not a downloadable file")
        return _save(self._fetch(url), dest, item.get("name") or item_id)

    def sites(self, search="*"):
        return self.list(f"sites?search={search}", client=C_SITES)

    def site_drive_items(self, site_id, path="root", top=100):
        return self.list(
            f"sites/{site_id}/drive/{path}/children", {"$top": top}, limit=top, client=C_SITES
        )

    # --- onenote ------------------------------------------------------------
    # Every method takes an optional `group_id` (a team id) to read that team's
    # or class's notebooks instead of your own.
    def _onenote(self, group_id=None):
        return f"groups/{group_id}/onenote" if group_id else "me/onenote"

    def notebooks(self, group_id=None):
        return self.list(f"{self._onenote(group_id)}/notebooks", client=C_ONENOTE)

    def sections(self, notebook_id=None, group_id=None):
        base = self._onenote(group_id)
        path = f"{base}/notebooks/{notebook_id}/sections" if notebook_id else f"{base}/sections"
        return self.list(path, client=C_ONENOTE)

    def pages(self, section_id, top=50, group_id=None):
        """Pages in a section, most recently modified first. (Listed per
        section and sorted here: the all-pages endpoint and `$select` +
        `$orderby` combinations on pages are flaky.)"""
        pages = self.list(
            f"{self._onenote(group_id)}/sections/{section_id}/pages",
            {"$top": min(top, 100)},
            limit=top,
            client=C_ONENOTE,
        )
        return sorted(pages, key=lambda p: p.get("lastModifiedDateTime") or "", reverse=True)

    def page_content(self, page_id, group_id=None):
        """A OneNote page's body as HTML (string)."""
        return self.get(
            f"{self._onenote(group_id)}/pages/{page_id}/content", raw=True, client=C_ONENOTE
        )

    def page_text(self, page_id, group_id=None, markdown=False):
        """A OneNote page as plain text (or Markdown)."""
        return html_to_text(self.page_content(page_id, group_id), markdown=markdown)

    def page_images(self, page_id, group_id=None):
        """The images embedded in a page (handwriting scans, screenshots of
        Lernziele, ...) as `{"alt", "content_type", "url", "data"}` dicts, in
        page order, at full resolution where available."""
        out = []
        for img in html_images(self.page_content(page_id, group_id)):
            url = img.get("data-fullres-src") or img.get("src")
            if not url or url.startswith("data:"):
                continue
            out.append(
                {
                    "alt": img.get("alt") or "",
                    "content_type": img.get("data-fullres-src-type") or img.get("data-src-type"),
                    "url": url,
                    "data": self.get(url, raw="bytes", client=C_ONENOTE),
                }
            )
        return out

    # --- people / tasks -----------------------------------------------------
    def contacts(self, top=50):
        return self.list("me/contacts", {"$top": top}, limit=top)

    def task_lists(self):
        return self.list("me/todo/lists")

    def tasks(self, list_id, top=50):
        return self.list(f"me/todo/lists/{list_id}/tasks", {"$top": top}, limit=top)

    # --- search (unified Graph /search/query) -------------------------------
    def search(self, query, entity="message", top=25):
        """Search one entity type across the account and return the matched
        resources, each with its `id` set to the hit id (usable with `email(id)`,
        `download(id)`, ...). `entity` is one of: 'message' (email), 'driveItem' (files),
        'chatMessage' (Teams chat/channel messages), 'event' (calendar),
        'person', 'site', 'listItem'. (Graph only allows one of these per
        request, so search one kind at a time.)"""
        client = _SEARCH_CLIENT.get(entity, C_OFFICE)
        body = {
            "requests": [{"entityTypes": [entity], "query": {"queryString": query}, "size": top}]
        }
        data = self.post("search/query", body, client=client)
        containers = (data.get("value") or [{}])[0].get("hitsContainers") or []
        hits = containers[0].get("hits", []) if containers else []
        # the hit's `hitId` is the item id (a message `resource` has none of its
        # own), so callers can follow up with email(id) etc.
        return [{**h["resource"], "id": h.get("hitId") or h["resource"].get("id")} for h in hits]

    def search_emails(self, query, top=25):
        return self.search(query, "message", top)

    def search_files(self, query, top=25):
        return self.search(query, "driveItem", top)

    def search_messages(self, query, top=25):
        """Search Teams chat and channel messages."""
        return self.search(query, "chatMessage", top)

    def search_people(self, query, top=25):
        return self.search(query, "person", top)

    # --- people & profile ---------------------------------------------------
    def me(self):
        self._me = self.get("me")
        return self._me

    def _my_id(self):
        return (self._me or self.me())["id"]

    def user(self, id_or_upn):
        """Another user's profile by id or UPN/email."""
        return self.get(f"users/{id_or_upn}")

    def member_of(self):
        """Groups, teams and directory roles the signed-in user belongs to."""
        return self.list("me/memberOf")

    def photo(self, user=None):
        """Profile photo bytes (yours, or `user`'s), or None if there is none."""
        try:
            return self.get(
                f"users/{user}/photo/$value" if user else "me/photo/$value", raw="bytes"
            )
        except GraphError as e:
            if e.status == 404:
                return None
            raise

    def presence(self, user_id=None):
        """Teams presence (availability + activity), yours or `user_id`'s."""
        path = f"users/{user_id}/presence" if user_id else "me/presence"
        return self._try_clients("presence", lambda c: self.get(path, client=c))

    def presences(self, user_ids):
        """Teams presence of several users (by id) in one call."""
        body = {"ids": list(user_ids)}
        return self._try_clients(
            "presence",
            lambda c: self.post("communications/getPresencesByUserId", body, client=c)["value"],
        )
