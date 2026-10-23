"""Teams: chats, channels, messages (read + send + lifecycle), shared images
and files, and online meetings."""

import re
from datetime import datetime
from urllib.parse import unquote

from ._base import GRAPH
from ._people import DirectoryMixin
from ._text import text_to_html

_MEMBER = "#microsoft.graph.aadUserConversationMember"


def _iso(value):
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds").replace("+00:00", "Z")
    return value


class TeamsMixin(DirectoryMixin):
    # --- chats --------------------------------------------------------------
    def chats(self, top=25, chat_type=None, expand=True):
        """1:1, group and meeting chats, most recently active first. With
        `expand` each chat carries its `members` and `lastMessagePreview`, so
        you can see who a 1:1 chat is with. `chat_type` filters on
        'oneOnOne' / 'group' / 'meeting'."""
        params = {"$top": min(top, 50)}
        if expand:
            params["$expand"] = "members,lastMessagePreview"
            params["$orderby"] = "lastMessagePreview/createdDateTime desc"
        if chat_type:
            params["$filter"] = f"chatType eq '{chat_type}'"
        return self.list("me/chats", params, limit=top)

    def chat(self, chat_id):
        return self.get(f"chats/{chat_id}", {"$expand": "members,lastMessagePreview"})

    def chat_members(self, chat_id):
        return self.list(f"chats/{chat_id}/members")

    def unread_chats(self, top=50):
        """Chats whose last message is newer than when I last read the chat
        and was not sent by me."""
        me = self._me_id()
        out = []
        for c in self.chats(top=top):
            last = c.get("lastMessagePreview") or {}
            sender = ((last.get("from") or {}).get("user") or {}).get("id")
            seen = (c.get("viewpoint") or {}).get("lastMessageReadDateTime") or ""
            if last.get("createdDateTime") and sender != me and last["createdDateTime"] > seen:
                out.append(c)
        return out

    def find_chat(self, user, top=200):
        """The 1:1 chat with `user` (UPN, email, id or name), or None."""
        u = self.resolve_user(user)
        for c in self.chats(top=top, chat_type="oneOnOne"):
            for m in c.get("members") or []:
                if m.get("userId") == u["id"]:
                    return c
        return None

    def chat_messages(self, chat_id, top=30, since=None):
        """Messages in a chat, newest first; `since` (datetime/ISO) keeps only
        messages created or edited after it."""
        params = {"$top": min(top, 50)}
        if since:
            params["$filter"] = f"lastModifiedDateTime gt {_iso(since)}"
            params["$orderby"] = "lastModifiedDateTime desc"
        return self.list(f"me/chats/{chat_id}/messages", params, limit=top)

    # --- sending (posts as the signed-in user) ------------------------------
    def _mention_body(self, text, html, mentions):
        content = text if html else text_to_html(text)
        if not mentions:
            return {"contentType": "html" if html else "text", "content": text}, None
        tags, out = [], []
        for i, m in enumerate(mentions):
            if isinstance(m, dict) and "id" in m:
                uid, name = m["id"], m.get("name") or m.get("displayName") or m["id"]
            else:
                u = self.resolve_user(m)
                uid, name = u["id"], u.get("displayName") or m
            tag = f'<at id="{i}">{text_to_html(name)}</at>'
            # `@Name` in the text becomes the mention in place; otherwise prepend
            if f"@{name}" in content:
                content = content.replace(f"@{name}", tag, 1)
            else:
                tags.append(tag)
            out.append(
                {
                    "id": i,
                    "mentionText": name,
                    "mentioned": {
                        "user": {"id": uid, "displayName": name, "userIdentityType": "aadUser"}
                    },
                }
            )
        if tags:
            content = " ".join(tags) + " " + content
        return {"contentType": "html", "content": content}, out

    def _message_payload(self, text, html=False, mentions=None, importance=None, subject=None):
        body, ments = self._mention_body(text, html, mentions)
        payload = {"body": body}
        if ments:
            payload["mentions"] = ments
        if importance:
            payload["importance"] = importance  # 'normal' / 'high' / 'urgent'
        if subject:
            payload["subject"] = subject
        return payload

    def send_chat_message(self, chat_id, text, html=False, mentions=None, importance=None):
        """Post a message in a chat as the signed-in user; returns the created
        chatMessage. `mentions` are users (UPN/id/name, or `{"id","name"}`
        dicts); an `@Name` in `text` is replaced by the mention in place,
        otherwise the mentions are prepended."""
        payload = self._message_payload(text, html, mentions, importance)
        return self.post(f"chats/{chat_id}/messages", payload)

    def create_chat(self, users, topic=None):
        """Start a chat with `users` (one = 1:1, Graph hands back the existing
        1:1 chat if there is one; more = group). The signed-in user is added
        automatically."""
        if isinstance(users, str):
            users = [users]
        ids = [self._me_id()] + [self.resolve_user(u)["id"] for u in users]
        members = [
            {
                "@odata.type": _MEMBER,
                "roles": ["owner"],
                "user@odata.bind": f"{GRAPH}/users('{uid}')",
            }
            for uid in ids
        ]
        body = {"chatType": "oneOnOne" if len(ids) == 2 else "group", "members": members}
        if topic and len(ids) > 2:
            body["topic"] = topic
        return self.post("chats", body)

    def send_to(self, user, text, **kw):
        """Send `text` to `user` in your 1:1 chat with them, creating the chat
        if needed. Returns the created chatMessage."""
        chat = self.create_chat([user])  # idempotent for 1:1 chats
        return self.send_chat_message(chat["id"], text, **kw)

    # --- message lifecycle --------------------------------------------------
    def edit_chat_message(self, chat_id, message_id, text, html=False):
        """Edit one of your own chat messages (Graph answers 204)."""
        body = {"contentType": "html" if html else "text", "content": text}
        return self.patch(f"chats/{chat_id}/messages/{message_id}", {"body": body})

    def delete_chat_message(self, chat_id, message_id):
        return self.post(f"users/{self._me_id()}/chats/{chat_id}/messages/{message_id}/softDelete")

    def undelete_chat_message(self, chat_id, message_id):
        return self.post(
            f"users/{self._me_id()}/chats/{chat_id}/messages/{message_id}/undoSoftDelete"
        )

    def react(self, chat_id, message_id, reaction="👍"):
        return self.post(
            f"chats/{chat_id}/messages/{message_id}/setReaction", {"reactionType": reaction}
        )

    def unreact(self, chat_id, message_id, reaction="👍"):
        return self.post(
            f"chats/{chat_id}/messages/{message_id}/unsetReaction", {"reactionType": reaction}
        )

    def _me_tenant(self):
        return {"user": {"id": self._me_id(), "tenantId": self._tenant_id()}}

    def mark_chat_read(self, chat_id):
        return self.post(f"chats/{chat_id}/markChatReadForUser", self._me_tenant())

    def mark_chat_unread(self, chat_id):
        return self.post(f"chats/{chat_id}/markChatUnreadForUser", self._me_tenant())

    # --- chat management ----------------------------------------------------
    def rename_chat(self, chat_id, topic):
        """Set a group chat's topic (1:1 chats have none)."""
        return self.patch(f"chats/{chat_id}", {"topic": topic})

    def add_chat_member(self, chat_id, user, share_history=True):
        uid = self.resolve_user(user)["id"]
        body = {
            "@odata.type": _MEMBER,
            "roles": ["owner"],
            "user@odata.bind": f"{GRAPH}/users('{uid}')",
        }
        if share_history:
            body["visibleHistoryStartDateTime"] = "0001-01-01T00:00:00Z"
        return self.post(f"chats/{chat_id}/members", body)

    def remove_chat_member(self, chat_id, membership_id):
        """`membership_id` is the member's `id` from `chat_members()`."""
        return self.delete(f"chats/{chat_id}/members/{membership_id}")

    def pinned_messages(self, chat_id):
        return self.list(f"chats/{chat_id}/pinnedMessages", {"$expand": "message"})

    def pin_message(self, chat_id, message_id):
        bind = f"{GRAPH}/chats/{chat_id}/messages/{message_id}"
        return self.post(f"chats/{chat_id}/pinnedMessages", {"message@odata.bind": bind})

    def unpin_message(self, chat_id, pinned_id):
        """`pinned_id` is the pin's `id` from `pinned_messages()`."""
        return self.delete(f"chats/{chat_id}/pinnedMessages/{pinned_id}")

    def hide_chat(self, chat_id):
        return self.post(f"chats/{chat_id}/hideForUser", self._me_tenant())

    def unhide_chat(self, chat_id):
        return self.post(f"chats/{chat_id}/unhideForUser", self._me_tenant())

    # --- teams & channels ---------------------------------------------------
    def teams(self):
        return self.list("me/joinedTeams")

    def channels(self, team_id):
        return self.list(f"teams/{team_id}/channels")

    def channel_messages(self, team_id, channel_id, top=30):
        """Top-level posts in a channel (replies: `channel_message_replies`)."""
        return self.list(
            f"teams/{team_id}/channels/{channel_id}/messages", {"$top": min(top, 50)}, limit=top
        )

    def channel_message_replies(self, team_id, channel_id, message_id, top=50):
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

    def send_channel_message(
        self, team_id, channel_id, text, subject=None, html=False, mentions=None
    ):
        """Start a new thread in a channel. Needs `ChannelMessage.Send`, which
        is tenant-dependent on the FOCI clients (a 403 means it's missing)."""
        payload = self._message_payload(text, html, mentions, subject=subject)
        return self.post(f"teams/{team_id}/channels/{channel_id}/messages", payload)

    def reply_to_channel_message(
        self, team_id, channel_id, message_id, text, html=False, mentions=None
    ):
        payload = self._message_payload(text, html, mentions)
        return self.post(
            f"teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies", payload
        )

    # --- images & files shared in messages ----------------------------------
    @staticmethod
    def _message_path(message_id, chat_id=None, team_id=None, channel_id=None, reply_to=None):
        if chat_id:
            return f"chats/{chat_id}/messages/{message_id}"
        if team_id and channel_id:
            base = f"teams/{team_id}/channels/{channel_id}/messages"
            return f"{base}/{reply_to}/replies/{message_id}" if reply_to else f"{base}/{message_id}"
        raise ValueError("pass chat_id, or team_id and channel_id")

    def message_images(
        self, message_id, chat_id=None, team_id=None, channel_id=None, reply_to=None
    ):
        """Images pasted into a message (its hostedContents), as a list of
        `{"id", "content_type", "data"}`."""
        path = self._message_path(message_id, chat_id, team_id, channel_id, reply_to)
        out = []
        for h in self.list(f"{path}/hostedContents"):
            data = self.get(f"{path}/hostedContents/{h['id']}/$value", raw="bytes")
            out.append({"id": h["id"], "content_type": h.get("contentType"), "data": data})
        return out

    @staticmethod
    def message_files(msg):
        """Files attached to a message (`reference` attachments): name and
        `contentUrl`, ready for `download_shared(url)`."""
        return [
            {"id": a.get("id"), "name": a.get("name"), "url": a.get("contentUrl")}
            for a in (msg or {}).get("attachments") or []
            if a.get("contentType") == "reference"
        ]

    # --- online meetings ----------------------------------------------------
    @staticmethod
    def _join_url(join_url_or_event):
        url = join_url_or_event
        if isinstance(url, dict):
            url = (url.get("onlineMeeting") or {}).get("joinUrl") or url.get("onlineMeetingUrl")
        if not url:
            raise ValueError("no Teams join URL on that event")
        return url

    @staticmethod
    def meeting_thread_id(join_url_or_event):
        """The meeting chat's thread id (`19:meeting_...@thread.v2`), parsed
        from a Teams join URL or an event dict."""
        url = TeamsMixin._join_url(join_url_or_event)
        m = re.search(r"meetup-join/([^/?]+)", url)
        if not m:
            raise ValueError(f"not a Teams meeting join URL: {url}")
        return unquote(m.group(1))

    def online_meeting(self, join_url_or_event):
        """Resolve a Teams meeting from its join URL or a calendar event dict.
        Needs OnlineMeetings.Read, which the FOCI clients lack on many tenants
        (403); `meeting_chat_messages` works without it."""
        url = self._join_url(join_url_or_event)
        found = self.list("me/onlineMeetings", {"$filter": f"JoinWebUrl eq '{url}'"})
        return found[0] if found else None

    def meeting_chat_messages(self, meeting_or_event, top=50):
        """Messages of a meeting's chat, from an onlineMeeting, an event dict
        or a join URL (no OnlineMeetings scope needed - the thread id is in
        the join URL)."""
        m = meeting_or_event
        if isinstance(m, dict) and m.get("chatInfo"):
            thread = m["chatInfo"]["threadId"]
        else:
            thread = self.meeting_thread_id(m)
        return self.chat_messages(thread, top=top)

    def transcripts(self, meeting_id):
        """Transcripts of a meeting (needs OnlineMeetingTranscript.Read.All,
        tenant-dependent - 403 without it)."""
        return self.list(f"me/onlineMeetings/{meeting_id}/transcripts")

    def transcript_text(self, meeting_id, transcript_id):
        """A transcript as WebVTT text."""
        return self.get(
            f"me/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content",
            {"$format": "text/vtt"},
            raw=True,
        )
