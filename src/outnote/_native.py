"""The APIs Outlook and Teams on the web call themselves, reached with the same FOCI tokens.

Only used where measured to win: chat messages (a little faster), `unread_chats` and `find_chat`
(every chat, no 50-per-page expansion), and what Graph can't do at all (one-call mail startup, sync
cursors for chats). The plain chat list stays on Graph, which is as fast once sorted and has names.
They are undocumented and can change without notice, so every read that has a Graph equivalent falls
back to Graph when a native call fails, and results are converted to Graph's shape.
`OutNote(..., native=False)` turns them off.
"""

import json
import logging
import time
import uuid
from datetime import UTC, datetime

from ._base import GraphError, Transport

log = logging.getLogger("outnote.native")

SKYPE_SCOPE = "https://api.spaces.skype.com/.default"
PRESENCE_SCOPE = "https://presence.teams.microsoft.com/.default"
OUTLOOK_SCOPE = "https://outlook.office.com/.default"
_AUTHZ = "https://teams.microsoft.com/api/authsvc/v1.0/authz"
_OWA_STARTUP = "https://outlook.office.com/owa/startupdata.ashx?app=Mail&n=0"
_VIEW = "msnp24Equivalent"
_PAGE_MAX = 200
_CHAT_TYPES = {"OneToOneChat": "oneOnOne", "Chat": "group", "Meeting": "meeting"}
# A failed native call of any of these kinds means "use Graph instead".
_FALLBACK = (GraphError, OSError, KeyError, ValueError, TypeError, AttributeError)


def _iso_ms(ms):
    """Epoch milliseconds as Graph's timestamp format."""
    t = datetime.fromtimestamp(int(ms) / 1000, tz=UTC)
    return t.strftime("%Y-%m-%dT%H:%M:%S.") + f"{t.microsecond // 1000:03d}Z"


def _sender(raw_from, name):
    """`.../contacts/8:orgid:<id>` or `.../contacts/28:<bot>` as a Graph `from`."""
    mri = (raw_from or "").rsplit("/", 1)[-1]
    if not mri:
        return None
    if mri.startswith("8:orgid:"):
        return {"user": {"id": mri.removeprefix("8:orgid:"), "displayName": name}}
    return {"application": {"id": mri.split(":", 1)[-1], "displayName": name}}


def _body(m):
    kind = m.get("messagetype") or ""
    html = kind.startswith("RichText") or m.get("contenttype") == "text/html"
    return {"contentType": "html" if html else "text", "content": m.get("content") or ""}


def _files(m):
    """`properties.files` is a JSON string; returns Graph's `reference` attachments."""
    raw = (m.get("properties") or {}).get("files")
    try:
        files = json.loads(raw) if raw else []
    except (TypeError, ValueError):
        return []
    return [
        {
            "id": f.get("itemid") or f.get("id"),
            "name": f.get("fileName") or f.get("title"),
            "contentType": "reference",
            "contentUrl": f.get("objectUrl") or (f.get("fileInfo") or {}).get("shareUrl"),
        }
        for f in files
    ]


def message_from_native(m, chat_id=None):
    """A chatsvc message in Graph's chatMessage shape; the raw one stays under `_native`."""
    props = m.get("properties") or {}
    kind = m.get("messagetype") or ""
    edited = props.get("edittime")
    deleted = props.get("deletetime")
    created = m.get("originalarrivaltime") or m.get("composetime")
    return {
        "id": m.get("id"),
        "chatId": chat_id or m.get("conversationid"),
        "messageType": "message" if kind.startswith(("Text", "RichText")) else "systemEventMessage",
        "createdDateTime": created,
        "lastModifiedDateTime": _iso_ms(edited) if edited else created,
        "deletedDateTime": _iso_ms(deleted) if deleted else None,
        "from": _sender(m.get("from"), m.get("imdisplayname")),
        "body": _body(m),
        "attachments": _files(m),
        "importance": props.get("importance") or "normal",
        "_native": m,
    }


def chat_from_native(c):
    """A chatsvc conversation in Graph's chat shape, or None for what Graph doesn't call a chat
    (channels, teams, notification streams). 1:1 members come from the id; a name is only known for
    whoever sent the last message, so use `chats()` when you need every name."""
    tp = c.get("threadProperties") or {}
    chat_type = _CHAT_TYPES.get(tp.get("productThreadType"))
    if not chat_type:
        return None
    last = c.get("lastMessage") or {}
    preview = None
    if last.get("id"):
        preview = {
            "id": last["id"],
            "createdDateTime": last.get("originalarrivaltime") or last.get("composetime"),
            "from": _sender(last.get("from"), last.get("imdisplayname")),
            "body": _body(last),
            "messageType": "message",
        }
    horizon = ((c.get("properties") or {}).get("consumptionhorizon") or "").split(";")
    seen = horizon[0] if horizon[0].isdigit() else None  # matches Graph's viewpoint exactly
    members = []
    if chat_type == "oneOnOne":
        sender = ((preview or {}).get("from") or {}).get("user") or {}
        for uid in c["id"].removeprefix("19:").split("@")[0].split("_"):
            name = sender.get("displayName") if sender.get("id") == uid else None
            members.append({"userId": uid, "displayName": name})
    return {
        "id": c["id"],
        "chatType": chat_type,
        "topic": tp.get("topic"),
        "createdDateTime": tp.get("createdat"),
        "lastMessagePreview": preview,
        "viewpoint": {"lastMessageReadDateTime": _iso_ms(seen) if seen else None},
        "members": members,
        "_native": c,
    }


def _last_activity(chat):
    return (chat.get("lastMessagePreview") or {}).get("createdDateTime") or ""


class NativeMixin(Transport):
    # --- Teams chat service ---------------------------------------------------
    def _chatsvc(self):
        """Region URL and session header, from Teams' own sign-in. Kept in memory until it expires."""
        s = self._ids.get("chatsvc")
        if s and s["exp"] - 60 > time.time():
            return s
        j = self.request("POST", _AUTHZ, scope=SKYPE_SCOPE)
        tok = j["tokens"]
        s = {
            "url": j["regionGtms"]["chatService"].rstrip("/"),
            "presence": j["regionGtms"].get("unifiedPresence", "").rstrip("/"),
            "headers": {"Authentication": f"skypetoken={tok['skypeToken']}"},
            "exp": time.time() + int(tok.get("expiresIn", 3600)),
        }
        self._ids["chatsvc"] = s
        return s

    def _chatsvc_get(self, url, params=None):
        s = self._chatsvc()
        if not url.startswith("http"):
            url = f"{s['url']}/v1/users/ME/{url}"
        return self.get(url, params, headers=s["headers"], scope=SKYPE_SCOPE)

    def _native_chats(self, top, chat_type=None):
        """Newest chats through chatsvc, or None if Graph should answer instead."""
        if not self.native:
            return None
        try:
            out = []
            page = self._chatsvc_get("conversations", {"view": _VIEW, "pageSize": _PAGE_MAX})
            for _ in range(10):
                for c in page.get("conversations", []):
                    chat = chat_from_native(c)
                    if chat and (not chat_type or chat["chatType"] == chat_type):
                        out.append(chat)
                # Results aren't sorted by activity, so every page has to be read before cutting.
                link = (page.get("_metadata") or {}).get("backwardLink")
                if not link:
                    break
                page = self._chatsvc_get(link)
            out.sort(key=_last_activity, reverse=True)
            return out[:top]
        except _FALLBACK as e:
            log.warning("native chats failed, using Graph: %s", e)
            return None

    def _native_messages(self, chat_id, top, since=None):
        """Newest messages of a chat through chatsvc, or None if Graph should answer instead."""
        if not self.native:
            return None
        try:
            return self._read_messages(chat_id, top, since)
        except _FALLBACK as e:
            log.warning("native messages failed, using Graph: %s", e)
            return None

    def _read_messages(self, chat_id, top, since=None):
        out = []
        page = self._chatsvc_get(
            f"conversations/{chat_id}/messages",
            {
                "view": f"{_VIEW}|supportsMessageProperties",
                "pageSize": min(top, _PAGE_MAX),
                "startTime": 1,
            },
        )
        for _ in range(50):
            batch = [message_from_native(m, chat_id) for m in page.get("messages", [])]
            out.extend(batch)
            link = (page.get("_metadata") or {}).get("backwardLink")
            older = since and batch and batch[-1]["createdDateTime"] < since
            if len(out) >= top or not link or not batch or older:
                break
            page = self._chatsvc_get(link)
        if since:
            out = [m for m in out if m["lastModifiedDateTime"] > since]
        return out[:top]

    def _stream_messages(self, stream, top):
        if not self.native:
            raise ValueError("this needs the native Teams API: a brokered client")
        return self._read_messages(stream, top)

    def activity(self, top=30):
        """Your Teams activity feed (reactions, replies, mentions and so on), newest first, in
        Graph's message shape. Teams' own API, undocumented; Graph has no equivalent, so there is
        no fallback. Reading it doesn't mark anything as seen."""
        return self._stream_messages("48:notifications", top)

    def mentions(self, top=30):
        """Messages that @mentioned you, newest first. Same API and caveats as `activity`."""
        return self._stream_messages("48:mentions", top)

    def call_history(self, top=30):
        """Your Teams calls, newest first; details are in `_native`. Same API and caveats as
        `activity`."""
        return self._stream_messages("48:calllogs", top)

    def teams_presence(self, user_ids):
        """Availability of Teams users through Teams' own presence service, for tenants where
        Graph's presence permission is missing. Undocumented; returns its raw entries, each with
        `mri` and `presence`."""
        s = self._chatsvc()
        body = [{"mri": f"8:orgid:{uid}"} for uid in user_ids]
        return self.request(
            "POST", f"{s['presence']}/v1/presence/getpresence/", body, scope=PRESENCE_SCOPE
        )

    # --- typing and live events -----------------------------------------------
    def _control(self, chat_id, kind):
        if not self.native:
            raise ValueError("this needs the native Teams API: a brokered client")
        s = self._chatsvc()
        body = {"content": "", "messagetype": kind, "contenttype": "Application/Message"}
        return self.request(
            "POST",
            f"{s['url']}/v1/users/ME/conversations/{chat_id}/messages",
            body,
            headers=s["headers"],
            scope=SKYPE_SCOPE,
        )

    def send_typing(self, chat_id):
        """Show "typing" to the other people in a chat. The signal fades by itself, so while the user
        keeps typing, call this again about every 6 seconds, as Teams does. Teams' own API,
        undocumented, no Graph fallback."""
        self._control(chat_id, "Control/Typing")

    def clear_typing(self, chat_id):
        """Stop showing "typing" straight away, for when the user clears what they wrote. Sending a
        message ends it too."""
        self._control(chat_id, "Control/ClearTyping")

    def listen(self, reconnect=True, presence=()):
        """A live feed of Teams events: who's typing, new messages, edits and deletions, plus the
        online status of the user ids in `presence`, starting with their current state. Iterate it,
        ideally in a `with` block, and see `LiveEvent`. It reconnects by itself and yields a
        `connected` event each time; anything sent while it was down is missed, so catch up with
        `chat_messages` or `stream`. Uses Teams' own push service, undocumented, no Graph fallback."""
        if not self.native:
            raise ValueError("this needs the native Teams API: a brokered client")
        from ._live import LiveConnection

        return LiveConnection(self, reconnect, presence)

    # --- Outlook ------------------------------------------------------------
    def mail_startup(self, top=25, timezone_name="UTC", conversations=True, focused=True):
        """What Outlook on the web loads at boot, in one request: `findConversation` (the first
        `top` conversations), `findFolders`, `owaUserConfig` and more, as Outlook's own raw JSON
        (not Graph's shape). Undocumented, needs a brokered client; no Graph fallback because Graph
        has no equivalent."""
        sort = [
            {
                "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": f"conversation:{field}"},
                "Order": "Descending",
            }
            for field in ("LastDeliveryOrRenewTime", "LastDeliveryTime")
        ]
        headers = {
            "action": "StartupData",
            "X-AnchorMailbox": f"SMTP:{self._upn()}",
            "x-req-source": "Mail",
            "x-owa-correlationid": str(uuid.uuid4()),
            "x-owa-sessionid": str(uuid.uuid4()),
            "x-message-count": str(top),
            "folderparams": json.dumps({"TimeZoneStr": timezone_name, "FolderPaneBitFlags": 2}),
            "messageparams": json.dumps(
                {
                    "TimeZoneStr": timezone_name,
                    "InboxReadingPanePosition": 1,
                    "IsFocusedInboxOn": focused,
                    "BootWithConversationView": conversations,
                    "SortResults": sort,
                    "IsSenderScreeningSettingEnabled": False,
                }
            ),
        }
        return self.request("POST", _OWA_STARTUP, headers=headers, scope=OUTLOOK_SCOPE)
