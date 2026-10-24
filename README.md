<p align="center">
  <img src="assets/outnote-icon.svg" width="140" alt="OutNote Logo">
</p>

<h1 align="center">OutNote</h1>

<p align="center">
  <strong>Headless Microsoft 365 over Graph - mail, calendar, Teams, OneDrive, SharePoint, OneNote, To Do, Planner and Assignments, without an app registration.</strong>
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

---

## What is OutNote?

Every "read my mailbox over Graph" tutorial makes you register an Azure app and get **admin consent** - which a locked-down tenant (a school, a company) simply refuses. OutNote sidesteps that entirely.

It signs in (via [`entrance`](https://github.com/PianoNic/Entrance)) with the pre-consented **OfficeHome** first-party client, gets a **FOCI family refresh token**, and a broker trades it for whichever first-party client's token a given call needs - all without another sign-in, and each already consented in the tenant. One headless login (password + TOTP, no browser), then clean reads *and writes* across the whole account.

## Features

- 🔓 **No app registration, no admin consent** - rides pre-consented Microsoft first-party clients
- 🤖 **Fully headless** - password + TOTP MFA over pure HTTP (or a browser fallback via [`entrance`](https://github.com/PianoNic/Entrance))
- 💾 **Log in once** - the family refresh token is persisted (0600) and reused, so later runs skip the sign-in
- 📬 **Broad coverage** - mail, calendar, Teams chats & channels, OneDrive, SharePoint, OneNote, To Do, Planner, Assignments
- ✉️ **Writes too** - send mail, Teams messages (with @mentions), events, uploads, OneNote pages, tasks
- 🧾 **Text out** - Teams/OneNote HTML to text, `.docx`/`.pptx`/`.xlsx`/`.pdf` extraction
- ⚡ **Robust** - 429/5xx backoff, `$batch`, delta queries, nextLink paging
- 🖥️ **CLI** - `outnote --text chats --unread`

## Install

```bash
pip install outnote        # or: uv add outnote
```

Python 3.10+; [`ms-entrance`](https://github.com/PianoNic/Entrance) is pulled in automatically. `pip install "outnote[pdf]"` adds PDF text extraction.

## Quickstart

```python
from outnote import authenticate

note = authenticate("you@school.ch", "password", totp_secret="JBSWY3DPEHPK3PXP")

note.emails(top=5, unread=True)
note.chats()
note.send_to("friend@school.ch", "on my way")
note.download(item_id)
note.search_emails("invoice")
```

## Docs

- [Usage](docs/usage.md) - sign-in, reading, writing, the CLI
- [Method reference](docs/reference.md) - every method, its endpoint, and the caveats
- [FOCI clients](docs/clients.md) - the client/scope map and how to re-probe your tenant

## Disclaimer

OutNote is published for **educational and research purposes only** - to document how Microsoft 365's first-party OAuth2 clients and the Graph API behave for a browserless client.

Use it only against accounts you own, or with the explicit authorization of the account or tenant owner. You alone are responsible for your use of it and for complying with Microsoft's terms, your identity provider's terms, your institution's acceptable-use policy, and applicable law. The software is provided "as is", without warranty of any kind, and the author accepts no liability for misuse. Not affiliated with or endorsed by Microsoft.

## License

[MIT](LICENSE) (c) [PianoNic](https://github.com/PianoNic)
