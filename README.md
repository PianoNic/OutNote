<p align="center">
  <img src="assets/outnote-icon.svg" width="140" alt="OutNote Logo">
</p>

<h1 align="center">OutNote</h1>

<p align="center">
  <strong>Headless Microsoft 365 reader - mail, calendar, Teams, OneDrive, SharePoint and OneNote over Graph, without an app registration.</strong>
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

It signs in with [`entrance`](https://github.com/PianoNic/Entrance) using the pre-consented **OfficeHome** first-party client, which hands back a **FOCI family refresh token**. A tiny broker then trades that token for whichever first-party client's Graph token a given call needs - all without another sign-in. Each of those clients is already consented in the tenant, so there is **nothing to register and nobody to ask**.

The result: one headless login (password + TOTP MFA, no browser), then clean Graph reads across your whole Microsoft 365 account. Built to be the base layer for an MCP server.

## Features

- 🔓 **No app registration, no admin consent** - rides pre-consented Microsoft first-party clients
- 🤖 **Fully headless** - password + TOTP MFA over pure HTTP, no browser (via [`entrance`](https://github.com/PianoNic/Entrance))
- 🧩 **One login, every service** - a FOCI family token is brokered into the right client per call
- 📬 **Broad coverage** - mail, calendar, Teams chats & channels, channel files, OneDrive, SharePoint, OneNote
- 🍪 **Silent reuse** - a persisted session jar means password-free reruns
- 📦 **Tiny** - stdlib `urllib` Graph client, one dependency (`ms-entrance`)

## Installation

```bash
pip install outnote
# or
uv add outnote
```

Requires Python 3.10+; `ms-entrance` is pulled in automatically.

## Quickstart

```python
from outnote import authenticate

note = authenticate(
    "you@school.ch",
    "your-password",
    totp_secret="JBSWY3DPEHPK3PXP",  # Base32 TOTP, if the account has MFA
)

# mail & calendar
note.emails(top=5)
note.events(top=5)

# Teams chats
for chat in note.chats():
    note.chat_messages(chat["id"])

# Teams channels + files
for team in note.teams():
    for ch in note.channels(team["id"]):
        note.channel_messages(team["id"], ch["id"])
        note.channel_files(team["id"], ch["id"])  # files posted in the channel

# OneDrive & SharePoint
note.files()  # OneDrive items
note.download_url(item_id)  # direct, pre-authed download link
note.sites()  # SharePoint sites

# OneNote - list and read page content as HTML
for nb in note.notebooks():
    for sec in note.sections(nb["id"]):
        for page in note.pages(sec["id"]):
            html = note.page_content(page["id"])

note.get("me/drive/root/children")  # any Graph endpoint, raw
```

The first call saves `.ms_session.json`; every call after that skips the password and MFA.

**Already have a Graph token** (MSAL, a browser session, ...)? Skip auth entirely. A single static token only carries its own scopes, so OneNote/SharePoint calls may `403` if it lacks them:

```python
from outnote import OutNote

OutNote(access_token).emails()
```

## What it does

Verified end-to-end against a live locked-down EDU tenant:

| Area | Methods | Graph endpoint |
|---|---|---|
| Mail | `emails()`, `email(id)` | `/me/mailFolders/{f}/messages` |
| Calendar | `events(start, end)` | `/me/events`, `/me/calendarView` |
| Teams chats | `chats()`, `chat_messages(id)` | `/me/chats`, `/me/chats/{id}/messages` |
| Teams channels | `teams()`, `channels(t)`, `channel_messages(t, c)` | `/me/joinedTeams`, `/teams/{t}/channels[/{c}/messages]` |
| Channel files | `channel_files(t, c)` | `filesFolder` → `/drives/{d}/items/{i}/children` |
| OneDrive | `files(path)`, `download_url(id[, drive_id])` | `/me/drive/...` |
| SharePoint | `sites(search)`, `site_drive_items(id)` | `/sites`, `/sites/{id}/drive` |
| OneNote | `notebooks()`, `sections(nb)`, `pages(sec)`, `page_content(id)` | `/me/onenote/...` |
| Contacts / Tasks | `contacts()`, `task_lists()`, `tasks(list)` | `/me/contacts`, `/me/todo/lists` |
| Anything | `get(path)`, `list(path)` | any endpoint, follows `@odata.nextLink` |

## How It Works

`authenticate` runs a PKCE authorization-code sign-in for the **OfficeHome** client through `entrance` (which drives Microsoft's sign-in - including TOTP MFA - over pure HTTP and intercepts the `code` at the native-client redirect), then redeems it for a **FOCI family refresh token**. The `OutNote` client holds a broker that, per call, refreshes that token into the first-party client whose pre-consented scopes cover the request, caching each token until it expires:

| Service | FOCI client | Key scope it carries |
|---|---|---|
| Mail, Calendar, Chats, Channels, Files, Teams, Contacts, Tasks | Office `d3590ed6-...` | `Mail.ReadWrite`, `Chat.ReadWrite`, `ChannelMessage.Read.All`, `Files.ReadWrite.All`, `Calendars.ReadWrite` |
| OneNote (read page content) | Teams-classic `1fec8e78-...` | `Notes.ReadWrite.All` |
| SharePoint sites | OfficeHome `4765445b-...` | `Sites.Read.All` |

Which first-party clients are FOCI-family and pre-consented is tenant-dependent; the three above work on a standard Microsoft 365 / EDU tenant.

## Caveats

- Works against **managed** Entra tenants (`login.microsoftonline.com`) with these first-party clients pre-consented - the common case for Microsoft 365 / EDU.
- **Push / SMS MFA can't be done headless** - only TOTP (authenticator app). Enroll one and pass its Base32 secret.
- Microsoft tweaks its login pages and client registrations; if a flow breaks, a user-agent, field name or client ID is usually the culprit.

## Disclaimer

OutNote is published for **educational and research purposes only** - to document how Microsoft 365's first-party OAuth2 clients and the Graph API behave for a browserless client.

Use it only against accounts you own, or with the explicit authorization of the account or tenant owner. You alone are responsible for your use of it and for complying with Microsoft's terms, your identity provider's terms, your institution's acceptable-use policy, and applicable law. The software is provided "as is", without warranty of any kind, and the author accepts no liability for misuse. Not affiliated with or endorsed by Microsoft.

## License

[MIT](LICENSE) (c) [PianoNic](https://github.com/PianoNic)
