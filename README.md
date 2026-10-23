<p align="center">
  <img src="assets/outnote-icon.svg" width="140" alt="OutNote Logo">
</p>

<h1 align="center">OutNote</h1>

<p align="center">
  <strong>Headless Microsoft 365 client - mail, calendar, Teams, OneDrive, SharePoint, OneNote, To Do, Planner and Assignments over Graph, without an app registration.</strong>
</p>

<p align="center">
  <a href="https://github.com/PianoNic/OutNote"><img src="https://badgetrack.pianonic.ch/badge?tag=outnote&label=visits&color=7719AA&style=flat" alt="visits"/></a>
  <img src="https://img.shields.io/badge/python-3.10%2B-7719AA" alt="Python 3.10+"/>
  <a href="https://github.com/PianoNic/OutNote/releases"><img src="https://img.shields.io/github/v/release/PianoNic/OutNote?include_prereleases&color=7719AA&label=Latest%20Release" alt="Latest release"/></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-7719AA.svg" alt="MIT license"/></a>
  <a href="https://github.com/PianoNic/Entrance"><img src="https://img.shields.io/badge/built%20on-entrance-7719AA.svg" alt="Built on Entrance"/></a>
</p>

> [!WARNING]
> **For educational and research purposes only.** Use it solely against **accounts you own** or with the **explicit authorization** of the account/tenant owner. It drives a real sign-in with real credentials - treat the cookie jar and any tokens like passwords. See the [Disclaimer](#disclaimer) below.

## About The Project

Every "read my mailbox over Graph" tutorial makes you register an Azure app and get **admin consent** - which a locked-down tenant (a school, a company) simply refuses. OutNote sidesteps that entirely.

It signs in with [`entrance`](https://github.com/PianoNic/Entrance) using the pre-consented **OfficeHome** first-party client, which hands back a **FOCI family refresh token**. A tiny broker then trades that token for whichever first-party client's token a given call needs - all without another sign-in. Each of those clients is already consented in the tenant, so there is **nothing to register and nobody to ask**.

The result: one headless login (password + TOTP MFA, no browser), then clean reads *and writes* across your whole Microsoft 365 account.

## Features

- 🔓 **No app registration, no admin consent** - rides pre-consented Microsoft first-party clients
- 🤖 **Fully headless** - password + TOTP MFA over pure HTTP, no browser (via [`entrance`](https://github.com/PianoNic/Entrance))
- 🧩 **One login, every service** - a FOCI family token is brokered into the right client per call
- 💾 **Log in once** - the family refresh token is persisted atomically (0600) and reused, so later runs skip the sign-in entirely
- 📬 **Broad coverage** - mail, calendar, Teams chats & channels, OneDrive, SharePoint, OneNote, To Do, Planner, Teams Assignments
- ✉️ **Writes too** - send mail, Teams messages (with @mentions), events, uploads, OneNote pages, tasks
- 🧾 **Text out** - Teams/OneNote HTML to plain text, `.docx`/`.pptx`/`.xlsx`/`.pdf` text extraction
- ⚡ **Robust** - 429/5xx backoff with `Retry-After`, `$batch`, delta queries, nextLink paging
- 🖥️ **CLI** - `outnote --text chats --unread` for quick checks
- 📦 **Tiny** - stdlib `urllib` Graph client, one dependency (`ms-entrance`)

## Installation

```bash
pip install outnote
# or
uv add outnote
```

Requires Python 3.10+; `ms-entrance` is pulled in automatically. `pip install "outnote[pdf]"` adds PDF text extraction.

## Quickstart

```python
from outnote import authenticate

note = authenticate(
    "you@school.ch",
    "your-password",
    totp_secret="JBSWY3DPEHPK3PXP",  # Base32 TOTP, if the account has MFA
)

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

The first call saves `.ms_frt` (the family refresh token, 0600) and `.ms_session.json` (sign-in cookies); every run after that skips the password and MFA. Already have a refresh token? `OutNote.from_refresh_token(frt)`.

**Already have a Graph token** (MSAL, a browser session, ...)? Skip auth entirely. A single static token only carries its own scopes, so OneNote/SharePoint/Assignments calls may fail:

```python
from outnote import OutNote

OutNote(access_token).emails()
```

### Writing

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

### CLI

```bash
export OUTNOTE_USER=you@school.ch OUTNOTE_PASSWORD=... OUTNOTE_TOTP=BASE32   # first run only
outnote --text chats --unread
outnote --text mail --unread
outnote --text events
outnote send friend@school.ch "on my way"
outnote get me/drive/root/children
```

Credentials are read from the environment, never from argv; after the first sign-in the saved refresh token is enough.

## What it does

Verified end-to-end against a live locked-down EDU tenant:

| Area | Methods | Endpoint |
|---|---|---|
| Mail | `emails(folder, unread=)`, `unread_emails()`, `email(id)`, `mail_folders()` | `/me/mailFolders/{f}/messages` |
| Mail attachments | `attachments(id)`, `download_attachment(id, att)` | `/me/messages/{id}/attachments` |
| Mail writes | `send_mail`, `reply`, `forward`, `create_draft`, `send_draft`, `mark_read`, `flag`, `move_email`, `delete_email` | `/me/sendMail`, `/me/messages/...` |
| Calendar | `events(start, end)`, `event(id)`, `today()`, `calendars()` | `/me/events`, `/me/calendarView` |
| Scheduling | `free_busy(users, start, end)`, `find_meeting_times(attendees)` | `getSchedule`, `findMeetingTimes` |
| Calendar writes | `create_event`, `update_event`, `delete_event`, `respond_event` | `/me/events` |
| Teams chats | `chats(chat_type=)`, `chat(id)`, `chat_members(id)`, `unread_chats()`, `find_chat(user)`, `chat_messages(id, since=)` | `/me/chats` (members + last message expanded) |
| Teams sending | `send_chat_message(id, text, mentions=)`, `create_chat(users)`, `send_to(user, text)` | `/chats`, `/chats/{id}/messages` |
| Message lifecycle | `edit_chat_message`, `delete_chat_message`, `undelete_chat_message`, `react`, `unreact`, `mark_chat_read`, `mark_chat_unread` | `PATCH`, `softDelete`, `setReaction`, `markChatReadForUser` |
| Chat management | `rename_chat`, `add_chat_member`, `remove_chat_member`, `pinned_messages`, `pin_message`, `unpin_message`, `hide_chat`, `unhide_chat` | `/chats/{id}/...` |
| Teams channels | `teams()`, `channels(t)`, `channel_messages(t, c)`, `channel_message_replies(t, c, m)`, `channel_files(t, c)` | `/me/joinedTeams`, `/teams/{t}/channels/...` |
| Channel posting | `send_channel_message(t, c, text)`, `reply_to_channel_message(t, c, m, text)` | `/teams/{t}/channels/{c}/messages` |
| Shared in messages | `message_images(m, chat_id=)`, `message_files(msg)`, `download_shared(url)` | `hostedContents`, `/shares/u!{url}` |
| Meetings | `meeting_chat_messages(event)`, `meeting_thread_id(event)`, `online_meeting(url)`\*, `transcripts(id)`\* | thread id parsed from the join URL |
| OneDrive | `files(path)`, `file(id)`, `download(id)`, `download_url(id)`, `upload(src, dest)`, `delete_file(id)` | `/me/drive/...` |
| SharePoint | `sites(search)`, `followed_sites()`, `site_drive_items(id)` | `/sites`, `/me/followedSites` |
| OneNote | `notebooks(group_id=)`, `sections`, `pages`, `page_content`, `page_text`, `page_images`, `create_page`, `append_page`, `delete_page` | `/me/onenote/...`, `/groups/{id}/onenote/...` |
| People | `find_users(q)`, `resolve_user(name_or_upn)`, `user(id)`, `contacts()`, `photo()`, `presence()`, `presences(ids)`, `member_of()` | `/users`, `/me/presence` |
| To Do | `task_lists()`, `tasks(list)`, `create_task`, `complete_task`, `delete_task` | `/me/todo/lists` |
| Planner | `planner_tasks()`, `plans(group_id=)`, `plan_tasks(plan)`, `buckets(plan)`, `planner_task_details(task)` | `/me/planner`, `/planner/plans` |
| Assignments | `classes()`, `assignments(class_id=)`, `assignment_resources`, `submissions`, `due_assignments(days)` | Teams Assignments API (see below) |
| Search | `search_emails`, `search_files`, `search_messages`, `search_people`, `search(q, entity)` | `/search/query` |
| Robustness | `batch(requests)`, `delta(path, delta_link)` | `/$batch` (20 per call, throttled sub-requests retried), `/delta` |
| Text | `message_text(msg)`, `html_to_text(html)`, `extract_text(name, data)` | - |
| Anything | `get`, `post`, `patch`, `put`, `delete`, `list` | any endpoint, follows `@odata.nextLink` |

\* needs `OnlineMeetings.Read` / transcript scopes, which no FOCI client carries on many tenants (403). The meeting *chat* works without them.

## How It Works

`authenticate` runs a PKCE authorization-code sign-in for the **OfficeHome** client through `entrance` (which drives Microsoft's sign-in - including TOTP MFA - over pure HTTP and intercepts the `code` at the native-client redirect), then redeems it for a **FOCI family refresh token** and saves it. The `OutNote` client holds a broker that, per call, refreshes that token into the first-party client whose pre-consented scopes cover the request, caching each token until it expires:

| Service | FOCI client | Key scopes it carries |
|---|---|---|
| Mail, Calendar, Chats, Channels, Files, Teams, Contacts, To Do, Planner | Office `d3590ed6-...` | `Mail.ReadWrite`, `Mail.Send`, `Chat.ReadWrite`, `ChannelMessage.Send`, `Files.ReadWrite.All`, `Calendars.ReadWrite`, `Tasks.ReadWrite` |
| OneNote | Teams-classic `1fec8e78-...` | `Notes.ReadWrite.All` |
| SharePoint sites, Presence | OfficeHome `4765445b-...` | `Sites.Read.All`, `Presence.Read.All` |
| Teams Assignments | Office `d3590ed6-...` for `api://8f348934-...` | `EduAssignments.ReadWrite` (on the Assignments API, not Graph) |

Graph's own `/education` endpoints need `EduAssignments.*`, which no FOCI client carries. The Teams Assignments app instead talks to `assignments.edu.cloud.microsoft` (same object shapes, `edu/...` paths) - and the Office client can mint a token for that API, so OutNote does the same. See [docs/clients.md](docs/clients.md) for how this map was found and how to re-probe it on your tenant.

## Caveats

- Works against **managed** Entra tenants (`login.microsoftonline.com`) with these first-party clients pre-consented - the common case for Microsoft 365 / EDU.
- **Push / SMS MFA can't be done headless** - only TOTP (authenticator app). Enroll one and pass its Base32 secret.
- Which scopes each client carries is **tenant-dependent**. Online-meeting lookup and transcripts 403 on the tested tenant.
- SharePoint-index search (`search_files`, `sites(search)`, `search(entity="site")`) depends on the tenant's search service; when it errors, `followed_sites()`, `files()` and `get("me/drive/recent")` still work.
- Microsoft tweaks its login pages and client registrations; if a flow breaks, a user-agent, field name or client ID is usually the culprit.
- Treat `.ms_frt` and `.ms_session.json` like passwords (both are gitignored).

## Disclaimer

OutNote is published for **educational and research purposes only** - to document how Microsoft 365's first-party OAuth2 clients and the Graph API behave for a browserless client.

Use it only against accounts you own, or with the explicit authorization of the account or tenant owner. You alone are responsible for your use of it and for complying with Microsoft's terms, your identity provider's terms, your institution's acceptable-use policy, and applicable law. The software is provided "as is", without warranty of any kind, and the author accepts no liability for misuse. Not affiliated with or endorsed by Microsoft.

## License

[MIT](LICENSE) (c) [PianoNic](https://github.com/PianoNic)
