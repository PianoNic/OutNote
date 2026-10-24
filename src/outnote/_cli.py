"""`outnote` command line: quick reads and sends over the library.

Credentials come from the environment (`OUTNOTE_USER`, `OUTNOTE_PASSWORD`,
`OUTNOTE_TOTP`) or, after the first sign-in, from the saved refresh token
(`OUTNOTE_FRT`, default `.ms_frt`) - never from argv.
"""

import argparse
import json
import os
import sys

from entrance import LoginFailed, TokenError

from ._auth import authenticate
from ._base import GraphError
from ._text import message_text

_SIGNIN_HELP = """Not signed in. Do one of:
  - sign in in your browser (easiest):   outnote signin
  - or set credentials (first run only):
      OUTNOTE_USER / OUTNOTE_PASSWORD / OUTNOTE_TOTP   (set on Windows, export on POSIX)
Either way a refresh token is saved (OUTNOTE_FRT, default .ms_frt), so later runs
need neither. You can also add --browser to any command."""


def _client(browser=False):
    return authenticate(
        os.environ.get("OUTNOTE_USER"),
        os.environ.get("OUTNOTE_PASSWORD"),
        totp_secret=os.environ.get("OUTNOTE_TOTP"),
        cookies=os.environ.get("OUTNOTE_COOKIES", ".ms_session.json"),
        frt_store=os.environ.get("OUTNOTE_FRT", ".ms_frt"),
        interactive="always" if browser else False,
    )


def _who(chat, me_id):
    if chat.get("topic"):
        return chat["topic"]
    others = [
        m.get("displayName") or "?" for m in chat.get("members") or [] if m.get("userId") != me_id
    ]
    return ", ".join(others) or chat.get("chatType", "chat")


def _text(cmd, data, note):
    """A readable one-line-per-item rendering of `data` for `--text`."""
    if cmd == "me":
        return f"{data.get('displayName')} <{data.get('mail') or data.get('userPrincipalName')}>"
    if cmd == "mail":
        return "\n".join(
            f"{'*' if not m.get('isRead') else ' '} {m.get('receivedDateTime', '')[:16]}  "
            f"{(m.get('from') or {}).get('emailAddress', {}).get('name', '?')}: {m.get('subject')}"
            for m in data
        )
    if cmd == "events":
        return "\n".join(
            f"{e['start']['dateTime'][:16]} - {e['end']['dateTime'][11:16]}  {e.get('subject')}"
            for e in data
        )
    if cmd == "chats":
        me_id = note._me_id()
        lines = []
        for c in data:
            last = c.get("lastMessagePreview") or {}
            preview = message_text(last)[:60].replace("\n", " ")
            lines.append(f"{c['id']}  {_who(c, me_id)}: {preview}")
        return "\n".join(lines)
    if cmd == "messages":
        return "\n".join(
            f"{m.get('createdDateTime', '')[:16]}  "
            f"{((m.get('from') or {}).get('user') or {}).get('displayName', '?')}: {message_text(m)}"
            for m in reversed(data)
            if m.get("messageType") == "message"
        )
    if cmd in ("files", "search"):
        return "\n".join(
            f"{i.get('id')}  {i.get('name') or i.get('subject') or i.get('displayName')}"
            for i in data
        )
    return json.dumps(data, indent=2, ensure_ascii=False)


def main(argv=None):
    ap = argparse.ArgumentParser(prog="outnote", description="Microsoft 365 over Graph.")
    ap.add_argument("--text", action="store_true", help="readable output instead of JSON")
    ap.add_argument(
        "--browser",
        action="store_true",
        help="sign in in a browser (no OUTNOTE_USER/PASSWORD needed)",
    )
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser(
        "signin",
        aliases=["login"],
        help="sign in in a browser and save the session for later runs",
    )
    sub.add_parser("me")
    p = sub.add_parser("mail")
    p.add_argument("--folder", default="inbox")
    p.add_argument("--unread", action="store_true")
    p.add_argument("--top", type=int, default=15)
    p = sub.add_parser("events", help="today's events, or the next N with --top")
    p.add_argument("--top", type=int)
    p = sub.add_parser("chats")
    p.add_argument("--unread", action="store_true")
    p.add_argument("--top", type=int, default=20)
    p = sub.add_parser("messages")
    p.add_argument("chat_id")
    p.add_argument("--top", type=int, default=20)
    p = sub.add_parser("send", help="send a Teams 1:1 message")
    p.add_argument("user")
    p.add_argument("text")
    p = sub.add_parser("search")
    p.add_argument("query")
    p.add_argument("--entity", default="message")
    p.add_argument("--top", type=int, default=10)
    p = sub.add_parser("files")
    p.add_argument("path", nargs="?", default="root")
    p = sub.add_parser("get", help="raw GET of any Graph path")
    p.add_argument("path")
    a = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    signin = a.cmd in ("signin", "login")

    try:
        note = _client(browser=a.browser or signin)
    except LoginFailed:
        print(_SIGNIN_HELP, file=sys.stderr)
        return 2
    except TokenError as e:
        print(f"Sign-in failed: {e}", file=sys.stderr)
        if e.error in ("invalid_grant", "interaction_required"):
            print("\n" + _SIGNIN_HELP, file=sys.stderr)
        return 2

    if signin:
        me = note.me()
        store = os.environ.get("OUTNOTE_FRT", ".ms_frt")
        print(f"Signed in as {me.get('displayName')} <{me.get('userPrincipalName')}>.")
        print(f"Session saved to {store} - later commands reuse it, no sign-in needed.")
        return 0

    try:
        data = _run(a, note)
    except GraphError as e:
        print(f"Graph error: {e}", file=sys.stderr)
        return 1

    out = _text(a.cmd, data, note) if a.text else json.dumps(data, indent=2, ensure_ascii=False)
    print(out)
    return 0


def _run(a, note):
    if a.cmd == "me":
        data = note.me()
    elif a.cmd == "mail":
        data = note.emails(a.folder, a.top, unread=a.unread)
    elif a.cmd == "events":
        data = note.events(top=a.top) if a.top else note.today()
    elif a.cmd == "chats":
        data = note.unread_chats(a.top) if a.unread else note.chats(a.top)
    elif a.cmd == "messages":
        data = note.chat_messages(a.chat_id, a.top)
    elif a.cmd == "send":
        data = note.send_to(a.user, a.text)
    elif a.cmd == "search":
        data = note.search(a.query, a.entity, a.top)
    elif a.cmd == "files":
        data = note.files(a.path)
    else:
        data = note.get(a.path)
    return data


if __name__ == "__main__":
    raise SystemExit(main())
