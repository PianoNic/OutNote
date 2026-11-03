"""An optional local cache: `stream()` written into SQLite. The client itself stays stateless.

    cache = SQLiteCache("outnote.db", note)
    cache.items("mail")          # instant, from disk, possibly slightly old
    cache.sync("mail")           # fetches only what changed
    cache.items("mail")

Uses only the standard library. The file holds real mail and chats: it is created readable by you
alone where the OS supports that, and is not encrypted. See docs/caching.md.
"""

import json
import os
import sqlite3
import threading

from ._base import CursorExpired

_SORT_FIELDS = ("receivedDateTime", "createdDateTime", "lastModifiedDateTime")
_SCHEMA = """
CREATE TABLE IF NOT EXISTS items (
    kind TEXT NOT NULL, scope TEXT NOT NULL, id TEXT NOT NULL,
    sort TEXT NOT NULL DEFAULT '', body TEXT NOT NULL,
    PRIMARY KEY (kind, scope, id)
);
CREATE INDEX IF NOT EXISTS items_by_time ON items (kind, scope, sort DESC);
CREATE TABLE IF NOT EXISTS cursors (
    kind TEXT NOT NULL, scope TEXT NOT NULL, cursor TEXT NOT NULL,
    PRIMARY KEY (kind, scope)
);
"""


def _sort_key(item):
    preview = (item.get("lastMessagePreview") or {}).get("createdDateTime")
    start = (
        (item.get("start") or {}).get("dateTime") if isinstance(item.get("start"), dict) else None
    )
    return preview or start or next((item[f] for f in _SORT_FIELDS if item.get(f)), "")


def _scope_key(scope):
    return json.dumps(scope, sort_keys=True) if scope else ""


class SQLiteCache:
    """A cache of `stream()` kinds in one SQLite file. Safe to share between threads."""

    def __init__(self, path, client):
        self.client = client
        if path != ":memory:" and not os.path.exists(path):
            os.close(os.open(path, os.O_CREAT | os.O_WRONLY, 0o600))
        self._db = sqlite3.connect(path, check_same_thread=False)
        self._lock = threading.Lock()
        with self._lock, self._db:
            if path != ":memory:":
                self._db.execute("PRAGMA journal_mode=WAL")
            self._db.executescript(_SCHEMA)

    def sync(self, kind, **scope):
        """Bring one kind up to date and return `{"upserted": n, "deleted": n, "reset": bool}`.
        `scope` is passed to `stream` (`folder=`, `chat_id=`, ...) and each scope is cached
        separately. If the server has dropped the cursor, that scope is rebuilt from scratch."""
        key = _scope_key(scope)
        stats = {"upserted": 0, "deleted": 0, "reset": False}
        cursor = self.cursor(kind, **scope)
        try:
            self._apply(kind, key, self.client.stream(kind, cursor, **scope), stats)
        except CursorExpired:
            stats["reset"] = True
            with self._lock, self._db:
                self._db.execute("DELETE FROM items WHERE kind=? AND scope=?", (kind, key))
                self._db.execute("DELETE FROM cursors WHERE kind=? AND scope=?", (kind, key))
            self._apply(kind, key, self.client.stream(kind, None, **scope), stats)
        return stats

    def _apply(self, kind, key, changes, stats):
        """Commit each time the cursor moves (once a page): items and their cursor land together,
        and a commit per item made a 1150-message sync take 40 s instead of 1 s."""
        pending = None
        try:
            for change in changes:
                with self._lock:
                    if change.op == "upsert":
                        self._db.execute(
                            "INSERT OR REPLACE INTO items (kind, scope, id, sort, body)"
                            " VALUES (?, ?, ?, ?, ?)",
                            (kind, key, change.id, _sort_key(change.item), json.dumps(change.item)),
                        )
                        stats["upserted"] += 1
                    elif change.op == "delete":
                        self._db.execute(
                            "DELETE FROM items WHERE kind=? AND scope=? AND id=?",
                            (kind, key, change.id),
                        )
                        stats["deleted"] += 1
                    if change.cursor != pending:
                        self._db.execute(
                            "INSERT OR REPLACE INTO cursors (kind, scope, cursor) VALUES (?, ?, ?)",
                            (kind, key, change.cursor),
                        )
                        self._db.commit()
                        pending = change.cursor
        except BaseException:
            with self._lock:
                self._db.rollback()  # the unfinished page replays from the saved cursor
            raise

    def items(self, kind, limit=None, **scope):
        """Cached items of a kind, newest first. Never touches the network."""
        sql = "SELECT body FROM items WHERE kind=? AND scope=? ORDER BY sort DESC, id"
        args = [kind, _scope_key(scope)]
        if limit is not None:
            sql += " LIMIT ?"
            args.append(limit)
        with self._lock:
            return [json.loads(row[0]) for row in self._db.execute(sql, args)]

    def item(self, kind, item_id, **scope):
        """One cached item, or None. Never touches the network."""
        with self._lock:
            row = self._db.execute(
                "SELECT body FROM items WHERE kind=? AND scope=? AND id=?",
                (kind, _scope_key(scope), item_id),
            ).fetchone()
        return json.loads(row[0]) if row else None

    def cursor(self, kind, **scope):
        """The saved cursor, or None if this kind and scope were never synced."""
        with self._lock:
            row = self._db.execute(
                "SELECT cursor FROM cursors WHERE kind=? AND scope=?",
                (kind, _scope_key(scope)),
            ).fetchone()
        return row[0] if row else None

    def close(self):
        with self._lock:
            self._db.close()
