"""Native routing, Graph-shape conversion, stream() and SQLiteCache. No network."""

import base64
import json

import pytest

import outnote
from outnote import _base
from outnote._native import chat_from_native, message_from_native

ME = "11111111-1111-1111-1111-111111111111"
BOB = "22222222-2222-2222-2222-222222222222"
ONE_ON_ONE = f"19:{ME}_{BOB}@unq.gbl.spaces"
AUTHZ = {
    "tokens": {"skypeToken": "sk", "expiresIn": 3600},
    "regionGtms": {"chatService": "https://chat.test"},
}


def jwt(**claims):
    body = base64.urlsafe_b64encode(json.dumps(claims).encode()).rstrip(b"=").decode()
    return f"h.{body}.s"


def conv(cid=ONE_ON_ONE, product="OneToOneChat", when="2026-11-03T08:00:00.000Z", sender=BOB):
    return {
        "id": cid,
        "threadProperties": {"productThreadType": product, "topic": None},
        "properties": {"consumptionhorizon": "1793692800000;1;1"},
        "lastMessage": {
            "id": "9",
            "originalarrivaltime": when,
            "from": f"https://x/contacts/8:orgid:{sender}",
            "imdisplayname": "Bob",
            "messagetype": "RichText/Html",
            "content": "<p>hi</p>",
        },
    }


class Fake:
    """Stands in for `Transport.request`, answering by URL substring; first match wins."""

    def __init__(self, routes):
        self.routes, self.calls = routes, []

    def __call__(self, method, path, body=None, params=None, raw=False, client=None, **kw):
        self.calls.append({"method": method, "path": path, "params": params, **kw})
        for key, answer in self.routes:
            if key in path:
                if isinstance(answer, Exception):
                    raise answer
                return answer(params) if callable(answer) else answer
        raise AssertionError(f"unexpected call {method} {path}")


def client(monkeypatch, routes, native=True):
    c = outnote.OutNote(
        lambda client_id, scope=None: jwt(oid=ME, tid="t", upn="me@x.ch"), native=native
    )
    fake = Fake(routes)
    monkeypatch.setattr(c, "request", fake)
    c.fake = fake
    return c


# --- identity without a request --------------------------------------------------
def test_ids_come_from_the_token(monkeypatch):
    c = client(monkeypatch, [])
    assert (c._me_id(), c._tenant_id(), c._upn()) == (ME, "t", "me@x.ch")
    assert c.fake.calls == []


def test_static_token_never_goes_native():
    assert outnote.OutNote("tok").native is False


# --- conversion to Graph's shape --------------------------------------------------
def test_chat_conversion():
    chat = chat_from_native(conv())
    assert chat["chatType"] == "oneOnOne"
    assert [m["userId"] for m in chat["members"]] == [ME, BOB]
    assert [m["displayName"] for m in chat["members"]] == [None, "Bob"]
    assert chat["lastMessagePreview"]["from"]["user"] == {"id": BOB, "displayName": "Bob"}
    assert chat["lastMessagePreview"]["body"] == {"contentType": "html", "content": "<p>hi</p>"}
    assert chat["viewpoint"]["lastMessageReadDateTime"].endswith("Z")
    assert outnote.message_text(chat["lastMessagePreview"]) == "hi"


@pytest.mark.parametrize("product", ["TeamsStandardChannel", "TeamsTeam", "StreamOfNotes", None])
def test_things_graph_does_not_call_chats_are_dropped(product):
    assert chat_from_native(conv(product=product)) is None


def test_message_conversion():
    files = json.dumps([{"itemid": "f", "fileName": "a.pdf", "objectUrl": "https://f/a.pdf"}])
    m = message_from_native(
        {
            "id": "1",
            "messagetype": "Text",
            "content": "hello",
            "originalarrivaltime": "2026-01-01T00:00:00.000Z",
            "from": f"https://x/contacts/8:orgid:{BOB}",
            "imdisplayname": "Bob",
            "properties": {"files": files, "edittime": "1793692800000"},
        },
        "c",
    )
    assert m["messageType"] == "message" and m["chatId"] == "c"
    assert m["body"] == {"contentType": "text", "content": "hello"}
    assert m["lastModifiedDateTime"] != m["createdDateTime"]
    assert outnote.OutNote.message_files(m) == [
        {"id": "f", "name": "a.pdf", "url": "https://f/a.pdf"}
    ]
    bot = message_from_native(
        {"id": "2", "messagetype": "ThreadActivity/AddMember", "from": "https://x/contacts/28:bot"}
    )
    assert bot["messageType"] == "systemEventMessage" and bot["from"]["application"]["id"] == "bot"


# --- routing and fallback ----------------------------------------------------------
def test_native_chats_are_sorted_across_pages(monkeypatch):
    first = {
        "conversations": [conv("19:a@thread.v2", "Chat", "2026-01-01T00:00:00Z")],
        "_metadata": {"backwardLink": "https://chat.test/older"},
    }
    older = {
        "conversations": [conv("19:b@thread.v2", "Chat", "2026-02-01T00:00:00Z")],
        "_metadata": {},
    }
    c = client(
        monkeypatch,
        [("authz", AUTHZ), ("older", older), ("v1/users/ME/conversations", first)],
    )
    assert [x["id"] for x in c._native_chats(25)] == ["19:b@thread.v2", "19:a@thread.v2"]
    authz, *others = c.fake.calls
    assert authz["method"] == "POST" and authz["scope"] == "https://api.spaces.skype.com/.default"
    assert others[0]["headers"] == {"Authentication": "skypetoken=sk"}
    assert not any("graph.microsoft.com" in x["path"] for x in c.fake.calls)


def test_skype_session_is_reused(monkeypatch):
    page = {"conversations": [], "_metadata": {}}
    c = client(monkeypatch, [("authz", AUTHZ), ("conversations", page)])
    c._native_chats(5)
    c._native_chats(5)
    assert sum("authz" in x["path"] for x in c.fake.calls) == 1


def test_chat_list_stays_on_graph(monkeypatch):
    c = client(monkeypatch, [("me/chats", {"value": [{"id": "x"}]})])
    assert c.chats(top=1) == [{"id": "x"}]
    assert c.chats(top=1, expand=False) == [{"id": "x"}]
    assert all(x["path"].startswith(_base.GRAPH) for x in c.fake.calls)


@pytest.mark.parametrize("failure", [outnote.GraphError(503, "down"), KeyError("regionGtms")])
def test_native_failure_falls_back_to_graph(monkeypatch, failure):
    c = client(monkeypatch, [("authz", failure), ("me/chats", {"value": []})])
    assert c.unread_chats() == []
    assert any("me/chats" in x["path"] for x in c.fake.calls)


def test_no_unread_natively_does_not_ask_graph_too(monkeypatch):
    c = client(monkeypatch, [("authz", AUTHZ), ("conversations", {"conversations": []})])
    assert c.unread_chats() == []
    assert not any(x["path"].startswith(_base.GRAPH) for x in c.fake.calls)


def test_native_can_be_turned_off(monkeypatch):
    c = client(monkeypatch, [("me/chats", {"value": []})], native=False)
    c.unread_chats()
    assert not any("authz" in x["path"] for x in c.fake.calls)


def test_find_chat_uses_ids_without_listing_members(monkeypatch):
    page = {"conversations": [conv()], "_metadata": {}}
    c = client(
        monkeypatch, [("authz", AUTHZ), ("conversations", page), (f"users/{BOB}", {"id": BOB})]
    )
    assert c.find_chat(BOB)["id"] == ONE_ON_ONE


def test_unread_uses_native_read_state(monkeypatch):
    page = {
        "conversations": [
            conv("19:a@thread.v2", "Chat", "2030-01-01T00:00:00.000Z"),
            conv("19:old@thread.v2", "Chat", "2000-01-01T00:00:00.000Z"),
            conv("19:mine@thread.v2", "Chat", "2030-01-01T00:00:00.000Z", sender=ME),
        ],
        "_metadata": {},
    }
    c = client(monkeypatch, [("authz", AUTHZ), ("conversations", page)])
    assert [x["id"] for x in c.unread_chats()] == ["19:a@thread.v2"]


def test_chat_messages_filter_by_since(monkeypatch):
    msgs = {
        "messages": [
            {"id": "new", "messagetype": "Text", "originalarrivaltime": "2026-02-01T00:00:00Z"},
            {"id": "old", "messagetype": "Text", "originalarrivaltime": "2026-01-01T00:00:00Z"},
        ],
        "_metadata": {},
    }
    c = client(monkeypatch, [("authz", AUTHZ), ("messages", msgs)])
    got = c.chat_messages("19:a@thread.v2", since="2026-01-15T00:00:00Z")
    assert [m["id"] for m in got] == ["new"]


def test_mail_startup_sends_what_outlook_sends(monkeypatch):
    c = client(monkeypatch, [("startupdata", {"findConversation": {}})])
    assert "findConversation" in c.mail_startup(top=10)
    call = c.fake.calls[0]
    assert call["scope"] == "https://outlook.office.com/.default"
    assert call["headers"]["X-AnchorMailbox"] == "SMTP:me@x.ch"
    assert call["headers"]["x-message-count"] == "10"


# --- features found in the web clients ----------------------------------------------
@pytest.mark.parametrize(
    ("method", "stream"),
    [
        ("activity", "48:notifications"),
        ("mentions", "48:mentions"),
        ("call_history", "48:calllogs"),
    ],
)
def test_teams_streams(monkeypatch, method, stream):
    page = {"messages": [{"id": "1", "messagetype": "Text", "content": "x"}], "_metadata": {}}
    c = client(monkeypatch, [("authz", AUTHZ), ("messages", page)])
    assert [m["id"] for m in getattr(c, method)(5)] == ["1"]
    assert stream in c.fake.calls[-1]["path"]


def test_teams_streams_need_a_brokered_client():
    with pytest.raises(ValueError):
        outnote.OutNote("tok").activity()


def test_teams_presence(monkeypatch):
    authz = {**AUTHZ, "regionGtms": {**AUTHZ["regionGtms"], "unifiedPresence": "https://ups.test"}}
    c = client(monkeypatch, [("authz", authz), ("getpresence", [{"mri": "m"}])])
    assert c.teams_presence([BOB]) == [{"mri": "m"}]
    assert c.fake.calls[-1]["path"].startswith("https://ups.test/")


def test_mailbox_settings_use_the_teams_client(monkeypatch):
    seen = []

    def request(method, path, body=None, params=None, raw=False, client=None, **kw):
        seen.append((path, client))
        return {"value": []}

    c = outnote.OutNote("tok")
    monkeypatch.setattr(c, "request", request)
    c.mailbox_settings()
    c.auto_replies()
    c.working_hours()
    c.mail_rules()
    assert {client for _, client in seen} == {outnote.C_TEAMS}
    assert [path.split("?")[0].removeprefix(_base.GRAPH + "/") for path, _ in seen] == [
        "me/mailboxSettings",
        "me/mailboxSettings/automaticRepliesSetting",
        "me/mailboxSettings/workingHours",
        "me/mailFolders/inbox/messageRules",
    ]


def test_graph_extras_use_documented_paths(note):
    for path in (
        "me/people",
        "me/outlook/masterCategories",
        "me/drive/recent",
        "me/drive/sharedWithMe",
    ):
        note.rec.answers[("GET", _base.GRAPH + "/" + path)] = {"value": [{"id": path}]}
    assert note.people(1)[0]["id"] == "me/people"
    assert note.categories()[0]["id"] == "me/outlook/masterCategories"
    assert note.recent_files(1)[0]["id"] == "me/drive/recent"
    assert note.shared_with_me(1)[0]["id"] == "me/drive/sharedWithMe"
    note.rec.answers[("POST", "me/getMailTips")] = {"value": [{"mailboxFull": False}]}
    assert note.mail_tips("a@x.ch") == [{"mailboxFull": False}]
    assert note.rec.calls[-1]["body"]["EmailAddresses"] == ["a@x.ch"]


@pytest.fixture
def note(monkeypatch):
    c = outnote.OutNote("tok")
    rec = Recorder()
    monkeypatch.setattr(c, "request", rec)
    c.rec = rec
    return c


class Recorder:
    def __init__(self):
        self.calls, self.answers = [], {}

    def __call__(self, method, path, body=None, params=None, raw=False, client=None, **kw):
        self.calls.append({"method": method, "path": path, "body": body})
        for (m, prefix), resp in self.answers.items():
            if m == method and path.startswith(prefix):
                return resp
        raise AssertionError(f"unexpected {method} {path}")


# --- stream ---------------------------------------------------------------------
def test_stream_graph_pages_deletes_and_final_cursor(monkeypatch):
    pages = {
        "/delta": {
            "value": [{"id": "1"}, {"id": "2", "@removed": {}}],
            "@odata.nextLink": "https://g/next",
        },
        "next": {"value": [{"id": "3"}], "@odata.deltaLink": "https://g/done"},
    }
    c = client(monkeypatch, [("/delta", pages["/delta"]), ("https://g/next", pages["next"])])
    got = list(c.stream("mail"))
    assert [(x.op, x.id) for x in got] == [
        ("upsert", "1"),
        ("delete", "2"),
        ("upsert", "3"),
        ("sync", None),
    ]
    assert got[0].item == {"id": "1"} and got[1].item is None
    assert got[1].cursor == "https://g/next"  # last of the page: safe to resume after it
    assert got[0].cursor not in ("https://g/next", "https://g/done")  # replays, never skips
    assert got[-1].cursor == "https://g/done"


def test_stream_resumes_and_reports_no_changes(monkeypatch):
    c = client(
        monkeypatch, [("https://g/done", {"value": [], "@odata.deltaLink": "https://g/done2"})]
    )
    got = list(c.stream("mail", "https://g/done"))
    assert [(x.op, x.cursor) for x in got] == [("sync", "https://g/done2")]


def test_expired_cursor_is_typed(monkeypatch):
    c = client(monkeypatch, [("https://g/old", outnote.GraphError(410, "gone"))])
    with pytest.raises(outnote.CursorExpired):
        list(c.stream("mail", "https://g/old"))


def test_a_failure_without_a_cursor_is_not_called_expired(monkeypatch):
    c = client(monkeypatch, [("/delta", outnote.GraphError(404, "no such folder"))])
    with pytest.raises(outnote.GraphError) as e:
        list(c.stream("mail", folder="nope"))
    assert not isinstance(e.value, outnote.CursorExpired)


def test_unknown_kind_is_rejected(monkeypatch):
    with pytest.raises(ValueError):
        list(client(monkeypatch, []).stream("nope"))


def test_stream_chats_walks_sync_state_until_empty(monkeypatch):
    first = {"conversations": [conv()], "_metadata": {"syncState": "https://chat.test/s1"}}
    second = {
        "conversations": [conv(product="TeamsTeam")],
        "_metadata": {"syncState": "https://chat.test/s2"},
    }
    third = {"conversations": [], "_metadata": {"syncState": "https://chat.test/s2"}}
    c = client(
        monkeypatch,
        [("authz", AUTHZ), ("/s1", second), ("/s2", third), ("conversations", first)],
    )
    got = list(c.stream("chats"))
    assert [(x.op, x.id) for x in got] == [("upsert", ONE_ON_ONE), ("sync", None)]
    assert got[-1].cursor == "https://chat.test/s2"


def test_stream_messages_marks_deletions(monkeypatch):
    page = {
        "messages": [
            {"id": "1", "messagetype": "Text", "properties": {"deletetime": "1793692800000"}}
        ],
        "_metadata": {"syncState": "https://chat.test/s"},
    }
    empty = {"messages": [], "_metadata": {"syncState": "https://chat.test/s"}}
    c = client(monkeypatch, [("authz", AUTHZ), ("chat.test/s", empty), ("messages", page)])
    got = list(c.stream("messages", chat_id="19:a@thread.v2"))
    assert [(x.op, x.id) for x in got] == [("delete", "1"), ("sync", None)]


def test_chat_streams_need_a_brokered_client():
    with pytest.raises(ValueError):
        list(outnote.OutNote("tok").stream("chats"))


# --- SQLite cache -----------------------------------------------------------------
class Scripted:
    """A client whose stream() replays scripted runs and records the cursors it was given."""

    def __init__(self, *runs):
        self.runs, self.cursors = list(runs), []

    def stream(self, kind, cursor=None, **scope):
        self.cursors.append(cursor)
        run = self.runs.pop(0)
        if isinstance(run, Exception):
            raise run
        return iter(run)


def change(op, item_id, cursor, **fields):
    item = {"id": item_id, **fields} if op == "upsert" else None
    return outnote.Change("mail", op, item_id, item, cursor)


def test_cache_syncs_serves_and_resumes(tmp_path):
    first = [
        change("upsert", "old", "c1", receivedDateTime="2026-01-01"),
        change("upsert", "new", "c1", receivedDateTime="2026-02-01"),
        change("sync", None, "c2"),
    ]
    second = [change("delete", "old", "c3"), change("sync", None, "c3")]
    api = Scripted(first, second)
    cache = outnote.SQLiteCache(str(tmp_path / "o.db"), api)
    assert cache.items("mail") == []
    assert cache.sync("mail") == {"upserted": 2, "deleted": 0, "reset": False}
    assert [i["id"] for i in cache.items("mail")] == ["new", "old"]
    assert [i["id"] for i in cache.items("mail", limit=1)] == ["new"]
    assert cache.item("mail", "old")["id"] == "old" and cache.item("mail", "nope") is None
    assert cache.sync("mail")["deleted"] == 1
    assert api.cursors == [None, "c2"]
    assert [i["id"] for i in cache.items("mail")] == ["new"]
    assert cache.cursor("mail") == "c3"
    cache.close()


def test_cache_survives_a_restart(tmp_path):
    path = str(tmp_path / "o.db")
    a = outnote.SQLiteCache(path, Scripted([change("upsert", "1", "c"), change("sync", None, "c")]))
    a.sync("mail")
    a.close()
    b = outnote.SQLiteCache(path, None)
    assert [i["id"] for i in b.items("mail")] == ["1"] and b.cursor("mail") == "c"
    b.close()


def test_cache_keeps_scopes_apart():
    api = Scripted(
        [change("upsert", "in", "a"), change("sync", None, "a")],
        [change("upsert", "sent", "b"), change("sync", None, "b")],
    )
    cache = outnote.SQLiteCache(":memory:", api)
    cache.sync("mail", folder="inbox")
    cache.sync("mail", folder="sentitems")
    assert [i["id"] for i in cache.items("mail", folder="inbox")] == ["in"]
    assert cache.cursor("mail", folder="sentitems") == "b"
    assert cache.items("mail") == []


def test_cache_rebuilds_when_the_cursor_expires():
    api = Scripted(
        [change("upsert", "stale", "a"), change("sync", None, "a")],
        outnote.CursorExpired(410, "gone"),
        [change("upsert", "fresh", "b"), change("sync", None, "b")],
    )
    cache = outnote.SQLiteCache(":memory:", api)
    cache.sync("mail")
    assert cache.sync("mail")["reset"] is True
    assert [i["id"] for i in cache.items("mail")] == ["fresh"]
    assert api.cursors == [None, "a", None]


def test_cache_keeps_progress_when_a_run_fails_midway():
    def broken():
        yield change("upsert", "1", "c1")
        raise outnote.GraphError(503, "down")

    cache = outnote.SQLiteCache(":memory:", Scripted(broken()))
    with pytest.raises(outnote.GraphError):
        cache.sync("mail")
    assert cache.cursor("mail") == "c1" and cache.item("mail", "1")
