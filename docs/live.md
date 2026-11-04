# Live Teams events

`note.listen()` gives you a live feed of Teams events: messages, edits, deletions, who's typing and who's
online. It uses Trouter, the push connection Teams itself uses, so events arrive without polling. In testing,
a message reached the client in about 300 ms.

**This is undocumented and can change without notice.** There's no Graph fallback, so keep a normal way to
read data too. It needs a client from `authenticate` or `from_refresh_token`.

```python
with note.listen(presence=[user_id]) as live:
    for event in live:
        if event.kind == "connected":
            catch_up()  # anything sent while disconnected was missed
        elif event.kind == "message" and not event.from_me:
            print(event.sender_name, outnote.message_text(event.message))
        elif event.kind == "typing":
            show_typing(event.chat_id, event.sender_id)
        elif event.kind == "presence":
            show_status(event.sender_id, event.availability)
```

`listen()` blocks while you iterate it, so run it in a background thread if your app does other work. Call
`live.close()` to stop it, from any thread. It reconnects by itself, waiting longer after each failure.

## Events

| `event.kind` | Means | Useful fields |
|---|---|---|
| `connected` | The feed is up, on first connect and after every reconnect | |
| `message` | A new message | `chat_id`, `sender_id`, `sender_name`, `from_me`, `message` |
| `message_updated` | A message changed | same as `message` |
| `message_deleted` | A message was deleted | `chat_id`, `message` |
| `typing` | Someone is typing | `chat_id`, `sender_id`, `from_me` |
| `stop_typing` | Someone stopped | `chat_id`, `sender_id` |
| `presence` | Status of someone in `presence=` | `sender_id`, `availability`, `activity`, `from_me` |
| `other` | Anything else, such as read receipts and member changes | `chat_id`, `raw` |

`message` is in Graph's shape, and the original event is always in `raw`.

Things to know, from testing:

- **Typing events usually have no name.** `sender_name` was empty on them, so look the name up from the id,
  for example with `chat_members()`. Messages do include it.
- **Hide the typing indicator yourself.** While someone types, events arrived about every 4 seconds, and
  `stop_typing` never arrived in testing. So hide it after about 6 seconds without a new event, or when a
  message arrives from them.
- **You don't see your own typing.** It wasn't echoed back in testing.
- **`presence` starts with everyone's current state.** Listening doesn't publish your own status, so it
  doesn't make you look online. Status changes were only seen live for the signed-in user, and someone
  who was active in a chat stayed reported as `Offline`, so treat status as a hint.
- **Tokens renew by themselves.** The feed reconnects shortly before its session token expires, which can
  leave a short gap. Catch up on `connected`.

## Showing that you're typing

```python
note.send_typing(chat_id)  # the other people see you typing
note.clear_typing(chat_id)  # stop straight away, if the user clears what they wrote
```

The signal fades by itself, so call `send_typing` about every 6 seconds while the user types, as Teams does.
Sending a message ends it too. **Other people really see this**, so only send it when the user is typing.

## Behind an API

`listen()` holds one long-lived connection per user, so it belongs in a background service, not a request
handler. Forward events to your frontend with something like WebSockets or server-sent events, and check
that each one only goes to the user it belongs to. See [frontend-and-api.md](frontend-and-api.md).
