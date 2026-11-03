# Speed and caching

OutNote is **stateless by default**: it writes nothing to disk, and every call goes to Microsoft.
The only things it holds are in memory, for the life of the client: access tokens and your user id.

When you want the instant feel of Outlook and Teams on the web, add caching on top in one of two ways:

| You want | Use | You keep |
|---|---|---|
| A cache that just works | `SQLiteCache` | one SQLite file |
| Your own storage (Redis, Postgres, a queue, memory) | `note.stream()` | one cursor string per kind |
| A web app or API, using normal calls only | [frontend-and-api.md](frontend-and-api.md) | short-lived results |

`SQLiteCache` is built on `stream()`, so both behave the same.

## How the real clients do it

Measured from the traffic of Outlook and Teams on the web, against one real account:

| Technique | Outlook | Teams | In OutNote |
|---|---|---|---|
| Load from a local cache, then update | App code only; mail is fetched at boot | 100+ IndexedDB databases, holding chats and messages | `SQLiteCache.items()`, then `.sync()` |
| Fetch only what changed | Yes | Yes | `stream()` cursors |
| One startup request instead of many | `startupdata.ashx` returns config, folders and the first 25 conversations | No | `mail_startup()` |
| Ask only for the fields a list shows | Named server-side views | Own compact API | `emails()` default fields |
| Push instead of polling | Notification channel | WebSocket | Not included: poll `stream()` |

## With SQLite

No dependencies beyond Python itself.

```python
from outnote import SQLiteCache, authenticate

note = authenticate("you@org.com", "password", totp_secret="BASE32")
cache = SQLiteCache("outnote.db", note)

show(cache.items("mail", limit=25, folder="inbox"))  # instant, possibly slightly old
cache.sync("mail", folder="inbox")  # fetches only what changed
show(cache.items("mail", limit=25, folder="inbox"))
```

This is the pattern the web clients use: show what you have straight away, then refresh it. The
first `sync` reads everything; later ones took about 40-120 ms in testing when nothing had changed.

- `sync(kind, **scope)` returns `{"upserted": n, "deleted": n, "reset": bool}`.
- `items(kind, limit=None, **scope)` and `item(kind, id, **scope)` never touch the network.
- Each scope is cached separately, so pass the same `folder=` or `chat_id=` to both.
- If Microsoft drops a cursor, `sync` rebuilds that scope and reports `"reset": True`.
- A failed `sync` keeps every finished page and resumes from there next time.

The file holds your real mail and chats. It's created readable only by you where the OS supports
that, but **it isn't encrypted**, so treat it like the mailbox itself and keep it out of git.

## With your own storage

`stream()` keeps nothing between calls. The cursor is the only state, and you own it:

```python
cursor = my_store.get("mail-cursor")  # None the first time: reads everything

for change in note.stream("mail", cursor, folder="inbox"):
    if change.op == "upsert":
        my_store.put(change.id, change.item)
    elif change.op == "delete":
        my_store.remove(change.id)
    my_store.set("mail-cursor", change.cursor)
```

Each `Change` has `kind`, `op`, `id`, `item` and `cursor`. `op` is `upsert`, `delete`, or `sync`,
which is always last, carries no item, and means you're up to date.

Two rules keep this correct:

1. **Handle changes idempotently.** Saving `change.cursor` after any change is safe, but after a crash
   you may see a few changes again. You'll never miss one.
2. **Handle `CursorExpired`.** It means Microsoft no longer accepts your cursor, so clear that kind
   and start again from `None`.

Cursors are called with the user's token, so `stream()` raises `ValueError` for one that doesn't point to
the expected Microsoft service. If your API accepts cursors from outside, also encrypt or sign them, and
check they belong to that user.

```python
from outnote import CursorExpired

try:
    changes = list(note.stream("mail", cursor))
except CursorExpired:
    my_store.clear_mail()
    changes = list(note.stream("mail"))
```

### What you can stream

| Kind | Options | Source |
|---|---|---|
| `mail` | `folder="inbox"`, `select=` | Graph delta |
| `mail_folders` | | Graph delta |
| `events` | `start=`, `end=` (default: a month back to a year ahead) | Graph delta |
| `contacts` | | Graph delta |
| `files` | | Graph delta |
| `chats` | | Teams' own API |
| `messages` | `chat_id=`, which also accepts `48:notifications`, `48:mentions` and `48:calllogs` | Teams' own API |

`mail` streams one folder, so sync each folder you care about. `chats` and `messages` need a client
from `authenticate` or `from_refresh_token`, because a single static token can't reach Teams' API.

A large first sync can take a while: about 40 seconds for 14,000 OneDrive files in testing, while
1,150 messages took about 1 second.

## Keeping it fresh

There's no push yet, so poll. A sync that finds nothing is cheap, so a short loop is fine:

```python
while True:
    cache.sync("mail", folder="inbox")
    cache.sync("chats")
    time.sleep(30)
```

Keep the interval reasonable and don't run many loops per account, since Graph throttles by app and
user. A chat's messages only change when the chat does, so sync `chats` first and only sync
`messages` for chats whose `lastMessagePreview` moved.

## Faster calls without any cache

- **Mail lists:** `emails()` returns list fields with `bodyPreview`, about 10x faster than every
  field. Use `email(id)` for one full message, or `select="*"` for all fields. Adding `toRecipients`
  costs about 170 ms per 25 messages.
- **Many requests:** `batch()` sends up to 20 per call, and up to 4 calls at once.
- **Long-lived clients:** reuse one client, since connections and tokens are kept in memory.
- **Your ids:** `_me_id()` and `_tenant_id()` come from the token without a request.

## Native APIs

OutNote also uses the APIs the web clients call themselves, where measuring showed a real gain:

| Call | Native route | Why |
|---|---|---|
| `unread_chats()`, `find_chat()` | Teams chat service | Sees every chat: about 0.5 s against 1.1–3.2 s |
| `chat_messages()` | Teams chat service | About 100 ms against 170–230 ms |
| `stream("chats")`, `stream("messages")` | Teams chat service | Graph has no equivalent |
| `mail_startup()` | Outlook `startupdata.ashx` | One request for folders, conversations and config |
| `activity()`, `mentions()`, `call_history()` | Teams chat service | Graph has no equivalent |
| `teams_presence()` | Teams presence service | Works where Graph's presence permission is missing |

Everything else stays on Graph, including `chats()` and `emails()`, because Graph was as fast or
faster there.

**These APIs are undocumented and can change without notice.** Reads with a Graph equivalent fall
back to Graph automatically if a native call fails, and results are converted to Graph's shape. There
are small differences:

- `unread_chats()` and `find_chat()` list 1:1 members by id, with a name only for whoever sent the
  last message. `chat(id)` or `chats()` have every name.
- In `chat_messages()`, a quoted reply's text can be part of the message body instead of an attachment.
- Converted items keep the original under `_native`.

`stream("chats")`, `stream("messages")`, `mail_startup()`, `activity()`, `mentions()`, `call_history()` and
`teams_presence()` have no fallback. To stay on Graph for
everything else, use `authenticate(..., native=False)`.

## Throttling

Calls wait out `Retry-After`, up to about 60 seconds in total per call. After that they raise
`GraphError`, with `status` 429 and `retry_after` telling you how long Graph asked to wait.

Directory reads (`me()`, `user()`, `resolve_user()`, `find_users()`, `member_of()`) can hit a limit
tied to the client id, so when one client is throttled or refused, they immediately try the next
family client (Office, then Teams, then OfficeHome). This goes around that one client's limit rather than waiting it out, so don't use it
to justify high request rates.
