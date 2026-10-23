"""Outlook mail: read, attachments, folders, triage, drafts and sending."""

import base64
import os

from ._base import Transport


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
        a folder id), newest first. `unread=True` keeps only unread ones."""
        params = {"$top": top, "$select": select}
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
        out, stack = [], [(f, f["displayName"]) for f in top]
        while stack:
            f, path = stack.pop(0)
            out.append({**f, "path": path})
            if f.get("childFolderCount"):
                for c in self.list(f"me/mailFolders/{f['id']}/childFolders", {"$top": 100}):
                    stack.append((c, f"{path}/{c['displayName']}"))
        return out

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
