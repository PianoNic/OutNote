# FOCI clients and the scope map

OutNote never registers an app. It signs in once with a Microsoft first-party
client and then rides the **FOCI** ("Family of Client IDs") refresh token that
sign-in returns: any client in the family can redeem it for its *own*
pre-consented scopes. This page records which client carries which scope on the
tested tenant, and how to re-check it on yours.

## The flow

1. `authenticate()` runs a PKCE authorization-code sign-in for **OfficeHome**
   (`4765445b-32c6-49b0-83e6-1d93765276ca`) through `entrance`, with the
   native-client redirect `https://login.microsoftonline.com/common/oauth2/nativeclient`.
   That redirect redeems server-side, so no SPA `Origin` header is needed.
2. The token response carries `foci: 1` and a **family refresh token** (FRT).
   It is saved atomically (temp file + `os.replace`, mode 0600) to `.ms_frt`.
3. Per call, the `Broker` does a `refresh_token` grant with the FRT and the
   `client_id` of the client whose scopes the call needs, and caches the access
   token per `(client, scope)`. The FRT rotates on use and is written back.

## Map on the tested EDU tenant

| Need | Client | Scope (from the token's `scp`) |
|---|---|---|
| Mail, calendar, chats, channels, files, teams, contacts, To Do, Planner | Office `d3590ed6-52b3-4102-aeff-aad2292ab01c` | `Mail.ReadWrite`, `Mail.Send`, `Chat.ReadWrite`, `ChannelMessage.Send`, `Files.ReadWrite.All`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `People.Read.All` |
| OneNote | Teams-classic `1fec8e78-bce4-4aaf-ab1b-5451cc387264` | `Notes.ReadWrite.All`, `Sites.ReadWrite.All` |
| SharePoint sites, presence | OfficeHome `4765445b-32c6-49b0-83e6-1d93765276ca` | `Sites.Read.All`, `Presence.Read.All` |
| Teams Assignments | Office, resource `api://8f348934-64be-4bb2-bc16-c54c96789f43` | `EduAssignments.ReadWrite`, `EduCurricula.ReadWrite` |

Also FOCI and redeemable, but not needed by OutNote today: Outlook mobile
(`27922004-...`, has `Presence.Read.All`, `Sites.ReadWrite.All`), OneDrive
(`ab9b8c07-...`), Edge (`ecd6b820-...`), Authenticator (`4813382a-...`),
SharePoint (`d326c1ce-...`). On this tenant the Azure CLI / Azure PowerShell /
Visual Studio clients return `invalid_grant` (blocked by policy).

### Not reachable

- **`OnlineMeetings.Read`, `OnlineMeetingTranscript.Read.All`, Graph
  `EduAssignments.*`**: no FOCI client carries them. Teams web itself doesn't
  use Graph for these either.
- **Teams web** (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`, "T2.1") is **not** in the
  family: it is an SPA client whose tokens are bound to the browser origin and
  it rejects the FRT (`AADSTS70000`). The same goes for the Assignments app
  client `ccb65bcd-...`, which Teams web brokers on its own.

### How the Assignments workaround was found

Watching Teams web's network traffic in the Assignments tab showed the app
calling `https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work`
with a token whose audience is `8f348934-64be-4bb2-bc16-c54c96789f43`. Asking
each FOCI client for `api://8f348934-64be-4bb2-bc16-c54c96789f43/.default`
showed that Office, Teams-classic and OfficeHome can all mint that token. The API
mirrors Graph education with `edu/...` paths:

| Graph | Assignments API |
|---|---|
| `/education/me/classes` | `/edu/me/classes` |
| `/education/me/assignments` | `/edu/me/work` |
| `/education/classes/{c}/assignments[/{a}/resources\|/submissions]` | `/edu/classes/{c}/assignments[/{a}/resources\|/submissions]` |

### Meetings without OnlineMeetings scopes

A Teams join URL embeds the meeting chat's thread id
(`.../meetup-join/19%3ameeting_...%40thread.v2/0?...`). `meeting_chat_messages(event)`
parses it and reads `/me/chats/{threadId}/messages`, which only needs
`Chat.ReadWrite`.

## Re-probing your tenant

Scopes are tenant-dependent. With a saved FRT you can list what each client gets:

```python
import base64, json
from outnote._broker import Broker, TokenError, read_secret

broker = Broker(read_secret(".ms_frt"), store=".ms_frt")
for name, cid in {
    "office": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
    "teams-classic": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
    "officehome": "4765445b-32c6-49b0-83e6-1d93765276ca",
    "outlook-mobile": "27922004-5251-4030-b22d-91ecd9a37ea4",
}.items():
    try:
        payload = broker(cid).split(".")[1]
        payload += "=" * (-len(payload) % 4)
        print(name, json.loads(base64.urlsafe_b64decode(payload))["scp"])
    except TokenError as e:
        print(name, "no token:", e.error)
```

Pass a different scope (`broker(cid, "api://<resource>/.default")`) to check a
non-Graph resource. If a call 403s, find the client whose `scp` has the scope
and pass it as `client=` (or change the routing constant in the library).
