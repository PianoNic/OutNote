"""Changes as a stream: OutNote stores nothing, you keep one opaque cursor.

    cursor = load()                                  # None the first time: everything
    for change in note.stream("mail", cursor):
        if change.op == "delete":
            my_db.remove(change.id)
        elif change.op == "upsert":
            my_db.save(change.id, change.item)
        save(change.cursor)                          # safe after every change

`change.cursor` always resumes from a point that includes this change or an earlier one, so a crash
replays a few changes and never loses one: handle changes idempotently. See docs/caching.md.
"""

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from typing import Any
from urllib.parse import urlsplit

from ._base import _MAX_PAGES, CursorExpired, GraphError, _check_link
from ._mail import LIST_SELECT
from ._native import _VIEW, NativeMixin, chat_from_native, message_from_native

_EXPIRED = (400, 404, 409, 410)
_GRAPH_PAGE = 999  # 1150 messages: 0.7 s at 999 per page, 2.7 s at 50
_CHAT_PAGE = 200
_CHAT_HOST_SUFFIXES = (".teams.microsoft.com",)


@dataclass
class Change:
    """One change. `op` is `upsert`, `delete`, or `sync` (the last one of a run: nothing left, and
    `cursor` is the one to use next time)."""

    kind: str
    op: str
    id: Any
    item: Any
    cursor: str


def _graph_item(item):
    return ("delete", None) if "@removed" in item else ("upsert", item)


def _stamp(value):
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


class StreamMixin(NativeMixin):
    def stream(self, kind, cursor=None, **scope):
        """Yield `Change`s since `cursor` (None = everything). Stateless: nothing is kept between
        calls but the cursor you pass back. Raises `CursorExpired` when the server no longer
        accepts the cursor; start again from None. Raises `ValueError` for a cursor that doesn't
        point to the expected Microsoft service, because cursors are called with the user's token.

        `kind` and its options:
          * `mail` - `folder="inbox"`, one folder's messages, with `select=` as in `emails`
          * `mail_folders` - the mail folders
          * `events` - `start=`, `end=` (default: a month back to a year ahead, fixed by the cursor)
          * `contacts` - contacts
          * `files` - OneDrive
          * `chats` - Teams chats (native only, Graph has no equivalent)
          * `messages` - `chat_id=` one Teams chat's messages (native only)
        """
        if kind in ("chats", "messages"):
            return self._stream_chatsvc(kind, cursor, **scope)
        return self._stream_graph(kind, cursor, **scope)

    # --- Graph delta --------------------------------------------------------
    def _delta_start(self, kind, folder="inbox", select=LIST_SELECT, start=None, end=None):
        if kind == "mail":
            return f"me/mailFolders/{folder}/messages/delta", {"$select": select}
        if kind == "mail_folders":
            return "me/mailFolders/delta", None
        if kind == "contacts":
            return "me/contacts/delta", None
        if kind == "files":
            return "me/drive/root/delta", None
        if kind == "events":
            now = datetime.now(UTC)
            return "me/calendarView/delta", {
                "startDateTime": start or _stamp(now - timedelta(days=30)),
                "endDateTime": end or _stamp(now + timedelta(days=365)),
            }
        raise ValueError(f"unknown stream kind: {kind!r}")

    def _stream_graph(self, kind, cursor, **scope):
        if cursor:
            self._check_graph_link(cursor)
            url = cursor
        else:
            path, params = self._delta_start(kind, **scope)
            url = self._url(path, params)
        headers = {"Prefer": f"odata.maxpagesize={_GRAPH_PAGE}"}
        for _ in range(_MAX_PAGES * 20):
            try:
                page = self.get(url, headers=headers)
            except GraphError as e:
                if cursor and e.status in _EXPIRED:
                    raise CursorExpired(e.status, f"{kind} cursor no longer valid") from None
                raise
            done = page.get("@odata.deltaLink")
            following = done or page.get("@odata.nextLink")
            items = page.get("value", [])
            for i, item in enumerate(items):
                op, body = _graph_item(item)
                yield Change(
                    kind, op, item.get("id"), body, following if i == len(items) - 1 else url
                )
            if done or not following:
                yield Change(kind, "sync", None, None, following or url)
                return
            url = following
        raise GraphError(508, f"{kind} delta did not finish")

    # --- chatsvc sync -------------------------------------------------------
    def _stream_chatsvc(self, kind, cursor, chat_id=None):
        if not self.native:
            raise ValueError(f"stream({kind!r}) needs the native Teams API: a brokered client")
        if kind == "messages" and not chat_id:
            raise ValueError("stream('messages') needs chat_id=")
        if cursor:
            service = urlsplit(self._chatsvc()["url"]).hostname
            _check_link(cursor, "https", hosts=(service,), suffixes=_CHAT_HOST_SUFFIXES)
            url, params = cursor, None
        elif kind == "chats":
            url, params = "conversations", {"view": _VIEW, "pageSize": _CHAT_PAGE}
        else:
            url = f"conversations/{chat_id}/messages"
            params = {
                "view": f"{_VIEW}|supportsMessageProperties",
                "pageSize": _CHAT_PAGE,
                "startTime": 1,
            }
        key = "conversations" if kind == "chats" else "messages"
        current = cursor
        for _ in range(_MAX_PAGES * 20):
            try:
                page = self._chatsvc_get(url, params)
            except GraphError as e:
                if cursor and e.status in _EXPIRED:
                    raise CursorExpired(e.status, f"{kind} cursor no longer valid") from None
                raise
            following = (page.get("_metadata") or {}).get("syncState") or current
            changes = [self._chatsvc_change(kind, raw, chat_id) for raw in page.get(key, [])]
            changes = [c for c in changes if c]
            for i, (op, cid, body) in enumerate(changes):
                at = following if i == len(changes) - 1 else (current or following)
                yield Change(kind, op, cid, body, at)
            if not page.get(key) or following == url:
                yield Change(kind, "sync", None, None, following)
                return
            url, params, current = following, None, following
        raise GraphError(508, f"{kind} sync did not finish")

    @staticmethod
    def _chatsvc_change(kind, raw, chat_id):
        if kind == "chats":
            chat = chat_from_native(raw)
            return ("upsert", chat["id"], chat) if chat else None
        msg = message_from_native(raw, chat_id)
        if msg["deletedDateTime"]:
            return ("delete", msg["id"], None)
        return ("upsert", msg["id"], msg)
