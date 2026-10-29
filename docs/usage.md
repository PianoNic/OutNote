# Usage

## Sign in

```python
from outnote import authenticate

note = authenticate(
    "you@school.ch",
    "your-password",
    totp_secret="JBSWY3DPEHPK3PXP",  # Base32 TOTP, if the account has MFA
)
```

The first call saves `.ms_frt` (the FOCI family refresh token, mode 0600) and
`.ms_session.json` (sign-in cookies); every run after that skips the password
and MFA. For an interactive second factor pass `mfa=` (see below), or
`interactive=True` to finish the first sign-in in a browser, or enroll a TOTP
device once (see [`entrance`](https://github.com/PianoNic/Entrance)).

### Interactive MFA

`mfa=<MfaPrompt>` lets your UI pick the second factor only when Microsoft asks for
it: a number-matching push (`on_push` receives the number to tap) or a code from an
authenticator app (`otp(attempt)`, up to 3 tries), with a way to switch mid-way:

```python
from outnote import authenticate, MfaPrompt, Secret


class Prompt(MfaPrompt):
    def choose(self, methods):  # ["push", "otp"], only what the account offers
        return methods[0]

    def on_push(self, event):  # PushEvent: show event.number
        ...
    def otp(self, attempt):  # block until the user types a code
        return input("Code: ")


note = authenticate("you@school.ch", Secret("password"), mfa=Prompt(), frt_store=None, cookies=None)
```

Wrong password raises `InvalidCredentials`; federated/ADFS or no usable method
`Unsupported`; Conditional Access `Blocked`; a denied or timed-out push
`PushError`. Details: [Entrance docs](https://github.com/PianoNic/Entrance/blob/main/docs/mfa.md).

Already have a refresh token? `OutNote.from_refresh_token(frt)`. Already have a
Graph **access token** (MSAL, a browser session)? `OutNote(access_token)` - but a
single static token only carries its own scopes, so OneNote / SharePoint /
Assignments calls may fail.

## Reading

```python
# mail
note.emails(top=5, unread=True)
for a in note.attachments(msg_id):
    data = note.download_attachment(msg_id, a["id"])

# calendar
note.today()
note.free_busy(["you@school.ch", "friend@school.ch"], start, end)

# Teams chats - who it's with, last message, unread
for chat in note.chats():
    print(chat["members"], note.message_text(chat["lastMessagePreview"]))
note.unread_chats()
note.chat_messages(note.find_chat("friend@school.ch")["id"])

# Teams channels + files
for team in note.teams():
    for ch in note.channels(team["id"]):
        note.channel_messages(team["id"], ch["id"])
        note.channel_files(team["id"], ch["id"])

# OneDrive - bytes and text
data = note.download(item_id)
text = note.extract_text("notes.docx", data)

# OneNote - your notebooks or a class/team's
for nb in note.notebooks(group_id=team["id"]):
    for sec in note.sections(nb["id"], group_id=team["id"]):
        for page in note.pages(sec["id"], group_id=team["id"]):
            note.page_text(page["id"], group_id=team["id"])

# Teams Assignments
note.due_assignments(days=7)

# search across the account (each result carries a usable "id")
note.search_emails("invoice")
note.search_messages("standup")

note.get("me/drive/root/children")  # any Graph endpoint, raw
```

## Writing

These act as the signed-in user - mail is really sent, messages really posted:

```python
note.send_mail("friend@school.ch", "Hi", "Body", attachments=["report.pdf"])
note.send_to(
    "friend@school.ch", "hey @Friend", mentions=["friend@school.ch"]
)  # 1:1 chat, created if needed
msg = note.send_chat_message(chat_id, "hello")
note.edit_chat_message(chat_id, msg["id"], "hello!")
note.react(chat_id, msg["id"], "👍")
note.mark_chat_read(chat_id)
note.create_event("Study", start, end)  # attendees=[...] sends invitations
note.upload("notes.txt", "Documents/notes.txt")  # >4 MB goes through an upload session
note.create_page(section_id, title="Summary", body="<p>...</p>")
note.create_task(list_id, "Hand in essay", due=friday)
```

## CLI

Sign in once - `outnote signin` shows a short code to enter at
`https://login.microsoft.com/device` (no redirect, no copy-paste); after that the
session is saved and later commands need nothing:

```bash
outnote signin          # or: outnote login  (device code; --browser for a browser)
outnote --text chats --unread
outnote --text mail --unread
outnote --text events
outnote send friend@school.ch "on my way"
outnote get me/drive/root/children
```

Prefer headless? Set credentials for the first run instead (read from the
environment, never argv); the saved refresh token is reused after that:

```bash
export OUTNOTE_USER=you@school.ch OUTNOTE_PASSWORD=... OUTNOTE_TOTP=BASE32
outnote --text chats --unread
```
