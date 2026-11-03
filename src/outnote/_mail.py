"""Outlook mail: read, attachments, folders, triage, drafts and sending."""

import base64
import os

from ._base import C_TEAMS, GraphError, Transport

# What a list needs. Measured on 25 messages: about 45 ms for these, 215 ms once `toRecipients` is
# added, 505 ms for everything. `email(id)` fetches the rest of one message.
LIST_SELECT = (
    "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,"
    "importance,flag,conversationId,parentFolderId"
)


def _recipients(addrs):
    if not addrs:
        return []
    if isinstance(addrs, str):
        addrs = [addrs]
    return [{"emailAddress": {"address": a}} for a in addrs]


def _file_attachments(attachments):
    """`attachments` items are file paths or `(name, bytes)` tuples."""
    out = []
    for a in attachments or []:
        if isinstance(a, (str, os.PathLike)):
            with open(a, "rb") as f:
                name, data = os.path.basename(a), f.read()
        else:
            name, data = a
        out.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentBytes": base64.b64encode(data).decode(),
            }
        )
    return out


def _message(to, subject, body, html=False, cc=None, bcc=None, attachments=None):
    msg = {
        "subject": subject,
        "body": {"contentType": "HTML" if html else "Text", "content": body},
        "toRecipients": _recipients(to),
    }
    if cc:
        msg["ccRecipients"] = _recipients(cc)
    if bcc:
        msg["bccRecipients"] = _recipients(bcc)
    if attachments:
        msg["attachments"] = _file_attachments(attachments)
    return msg


class MailMixin(Transport):
    # --- read ---------------------------------------------------------------
    def emails(self, folder="inbox", top=25, search=None, select=None, unread=False):
        """Messages in `folder` (a well-known name like 'inbox'/'sentitems' or
        a folder id), newest first. `unread=True` keeps only unread ones. Returns
        list fields with `bodyPreview`, no `body` or `toRecipients`, which is about
        10x faster: use `email(id)` for one full message, `select="*"` for all fields."""
        params = {"$top": top, "$select": None if select == "*" else select or LIST_SELECT}
        if search:
            params["$search"] = f'"{search}"'
        else:
            params["$orderby"] = "receivedDateTime desc"
            if unread:
                params["$filter"] = "isRead eq false"
        return self.list(f"me/mailFolders/{folder}/messages", params, limit=top)

    def unread_emails(self, folder="inbox", top=25):
        return self.emails(folder, top, unread=True)

    def email(self, message_id):
        return self.get(f"me/messages/{message_id}")

    def mail_folders(self, include_children=False):
        """Mail folders (top-level; `include_children=True` walks subfolders
        too, each tagged with its `path`)."""
        top = self.list("me/mailFolders", {"$top": 100})
        if not include_children:
            return top
        out, level = [], [(f, f["displayName"]) for f in top]
        while level:  # one $batch per depth instead of one request per folder
            out.extend({**f, "path": path} for f, path in level)
            parents = [(f, path) for f, path in level if f.get("childFolderCount")]
            found = self.batch(
                [f"me/mailFolders/{f['id']}/childFolders?$top=100" for f, _ in parents]
            )
            level = []
            for (_, path), res in zip(parents, found, strict=True):
                if isinstance(res, GraphError):
                    raise res
                level.extend((c, f"{path}/{c['displayName']}") for c in res.get("value", []))
        return out

    # Mailbox settings need MailboxSettings.ReadWrite, which the Teams client has and Office doesn't.
    def mailbox_settings(self):
        """Time zone, language, date formats, working hours and auto-replies."""
        return self.get("me/mailboxSettings", client=C_TEAMS)

    def auto_replies(self):
        """Your out-of-office setting: `status`, the two reply messages and any schedule."""
        return self.get("me/mailboxSettings/automaticRepliesSetting", client=C_TEAMS)

    def working_hours(self):
        """Your working days and hours (`daysOfWeek`, `startTime`, `endTime`, `timeZone`)."""
        return self.get("me/mailboxSettings/workingHours", client=C_TEAMS)

    def mail_rules(self):
        """Your inbox rules."""
        return self.list("me/mailFolders/inbox/messageRules", client=C_TEAMS)

    def categories(self):
        """Your Outlook categories (`displayName`, `color`)."""
        return self.list("me/outlook/masterCategories")

    def mail_tips(self, addresses):
        """Before sending: auto-replies and full mailboxes of `addresses`."""
        if isinstance(addresses, str):
            addresses = [addresses]
        body = {
            "EmailAddresses": addresses,
            "MailTipsOptions": "automaticReplies,mailboxFullStatus",
        }
        return self.post("me/getMailTips", body)["value"]

    # --- attachments --------------------------------------------------------
    def attachments(self, message_id):
        """Attachment metadata (no content) for a message; [] if it has none."""
        items = self.list(f"me/messages/{message_id}/attachments")
        keep = ("id", "name", "contentType", "size", "isInline", "@odata.type")
        return [{k: a.get(k) for k in keep} for a in items]

    def download_attachment(self, message_id, attachment_id, dest=None):
        """An attachment's bytes (file attachments are base64-decoded; item
        attachments come back as MIME). Writes to `dest` if given."""
        a = self.get(f"me/messages/{message_id}/attachments/{attachment_id}")
        if a.get("contentBytes") is not None:
            data = base64.b64decode(a["contentBytes"])
        elif a.get("@odata.type", "").endswith("referenceAttachment"):
            raise ValueError(f"reference attachment (a link, not a file): {a.get('sourceUrl')}")
        else:
            data = self.get(
                f"me/messages/{message_id}/attachments/{attachment_id}/$value", raw="bytes"
            )
        if dest:
            with open(dest, "wb") as f:
                f.write(data)
        return data

    # --- send (mutates the mailbox) ----------------------------------------
    def send_mail(
        self, to, subject, body, html=False, cc=None, bcc=None, attachments=None, save=True
    ):
        """Send a mail as the signed-in user. `to`/`cc`/`bcc` are an address or
        a list; `attachments` are file paths or `(name, bytes)`. Returns None
        (Graph answers 202)."""
        msg = _message(to, subject, body, html, cc, bcc, attachments)
        return self.post("me/sendMail", {"message": msg, "saveToSentItems": save})

    def reply(self, message_id, comment, reply_all=False):
        """Reply (or reply-all) to a message with `comment`; sends immediately."""
        action = "replyAll" if reply_all else "reply"
        return self.post(f"me/messages/{message_id}/{action}", {"comment": comment})

    def forward(self, message_id, to, comment=""):
        return self.post(
            f"me/messages/{message_id}/forward",
            {"comment": comment, "toRecipients": _recipients(to)},
        )

    def create_draft(self, to, subject, body, html=False, cc=None, bcc=None, attachments=None):
        """Save a draft (not sent); returns the draft message."""
        return self.post("me/messages", _message(to, subject, body, html, cc, bcc, attachments))

    def send_draft(self, message_id):
        return self.post(f"me/messages/{message_id}/send")

    # --- triage -------------------------------------------------------------
    def mark_read(self, message_id, read=True):
        return self.patch(f"me/messages/{message_id}", {"isRead": read})

    def flag(self, message_id, status="flagged"):
        """`status` is 'flagged', 'complete' or 'notFlagged'."""
        return self.patch(f"me/messages/{message_id}", {"flag": {"flagStatus": status}})

    def move_email(self, message_id, folder):
        """Move to a folder id or well-known name ('archive', 'deleteditems',
        ...); returns the moved message (it gets a new id)."""
        return self.post(f"me/messages/{message_id}/move", {"destinationId": folder})

    def delete_email(self, message_id):
        return self.delete(f"me/messages/{message_id}")
