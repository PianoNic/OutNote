"""Typing and the live feed, with a fake socket. No network."""

import json
import threading
import time

import pytest

import outnote
from outnote import _live
from outnote._live import LiveConnection, LiveEvent, event_from, parse_frame

ME = "11111111-1111-1111-1111-111111111111"
BOB = "22222222-2222-2222-2222-222222222222"
CHAT = "19:abc@thread.v2"
AUTHZ = {
    "tokens": {"skypeToken": "sk", "expiresIn": 3600},
    "regionGtms": {"chatService": "https://chat.test"},
}


def resource(kind="Text", sender=BOB, **more):
    return {
        "id": "1",
        "messagetype": kind,
        "content": "hi",
        "from": f"https://x/contacts/8:orgid:{sender}",
        "imdisplayname": "Bob",
        "conversationLink": f"https://x/v1/users/ME/conversations/{CHAT}",
        "originalarrivaltime": "2026-01-01T00:00:00.000Z",
        **more,
    }


def raw(kind="Text", event="NewMessage", **more):
    return {"resourceType": event, "resource": resource(kind, **more)}


# --- event parsing ----------------------------------------------------------------
def test_typing_events():
    e = event_from(raw("Control/Typing"), ME)
    assert (e.kind, e.chat_id, e.sender_id, e.sender_name) == ("typing", CHAT, BOB, "Bob")
    assert e.from_me is False and e.message is None
    assert event_from(raw("Control/ClearTyping"), ME).kind == "stop_typing"


def test_message_events_use_graph_shape():
    e = event_from(raw("RichText/Html"), ME)
    assert e.kind == "message" and e.message["chatId"] == CHAT
    assert e.message["body"]["contentType"] == "html"
    assert event_from(raw(event="MessageUpdate"), ME).kind == "message_updated"
    gone = raw(event="MessageUpdate", properties={"deletetime": "1793692800000"})
    assert event_from(gone, ME).kind == "message_deleted"


def test_own_activity_is_marked():
    assert event_from(raw("Control/Typing", sender=ME), ME).from_me is True


@pytest.mark.parametrize(
    "item",
    [
        raw("ThreadActivity/AddMember"),
        raw("Control/Something"),
        {"resourceType": "ThreadUpdate", "resource": {"id": "x"}},
        {"resourceType": "NewMessage"},
    ],
)
def test_other_events_stay_other_and_keep_raw(item):
    e = event_from(item, ME)
    assert e.kind == "other" and e.raw is item and e.message is None


def test_encoded_chat_ids_are_decoded():
    item = raw()
    item["resource"]["conversationLink"] = "https://x/conversations/19%3Aabc%40thread.v2"
    assert event_from(item).chat_id == CHAT


def test_parse_frame():
    assert parse_frame("1::") == ("1", "", "")
    assert parse_frame('3:::{"a":1}') == ("3", "", '{"a":1}')
    assert parse_frame('5:7+::{"name":"x:y"}') == ("5", "7+", '{"name":"x:y"}')


# --- the connection ---------------------------------------------------------------
class FakeSocket:
    def __init__(self, frames):
        self.frames, self.sent, self.closed = list(frames), [], False

    def recv(self, timeout=None):
        if self.frames:
            item = self.frames.pop(0)
            if isinstance(item, Exception):
                raise item
            return type("M", (), {"text": item})()
        time.sleep(0.01)
        return None

    def send(self, message):
        self.sent.append(message)

    def close(self):
        self.closed = True


def delivery(*items, id=9):
    body = json.dumps({"eventMessages": list(items)})
    return "3:::" + json.dumps({"id": id, "method": "POST", "url": "/x", "body": body})


@pytest.fixture
def connected(monkeypatch):
    """A LiveConnection whose _connect installs the next fake socket."""
    client = outnote.OutNote("tok")
    client._ids["me"] = ME
    client._ids["chatsvc"] = {"exp": time.time() + 3600, "headers": {}, "url": "https://chat.test"}
    sockets = []

    def make(frames_per_connection, **kw):
        queue = [FakeSocket(f) for f in frames_per_connection]
        live = LiveConnection(client, **kw)

        def connect():
            if not queue:
                raise OSError("no more connections")
            live._ws = queue.pop(0)
            sockets.append(live._ws)

        monkeypatch.setattr(live, "_connect", connect)
        monkeypatch.setattr(live, "_send", lambda text: live._ws.send(text))
        monkeypatch.setattr(live, "_sleep", lambda s: None)
        return live, sockets

    return make


def take(live, n):
    out = []
    for event in live:
        out.append(event)
        if len(out) >= n:
            live.close()
    return out


def test_events_arrive_and_deliveries_are_acknowledged(connected):
    live, sockets = connected([["1::", delivery(raw("Control/Typing"))]])
    events = take(live, 2)
    assert [e.kind for e in events] == ["connected", "typing"]
    ack = json.loads(sockets[0].sent[0].split(":::", 1)[1])
    assert ack == {"id": 9, "status": 200, "headers": {}, "body": ""}


def test_heartbeats_and_acks_are_answered(connected):
    live, sockets = connected([["2::", '5:4+::{"name":"trouter.connected"}', delivery(raw())]])
    take(live, 2)
    assert "2:::" in sockets[0].sent and "6:::4+[]" in sockets[0].sent


def test_bad_deliveries_are_survived(connected):
    bad = "3:::" + json.dumps({"id": 1, "body": "not json"})
    live, _ = connected([["3:::not json", bad, delivery(raw())]])
    assert [e.kind for e in take(live, 2)] == ["connected", "message"]


def test_it_reconnects_and_says_so(connected):
    live, sockets = connected([[OSError("dropped")], ["1::", delivery(raw())]])
    events = take(live, 3)
    assert [e.kind for e in events] == ["connected", "connected", "message"]
    assert sockets[0].closed


def test_server_disconnect_frame_reconnects(connected):
    live, _ = connected([["0::"], [delivery(raw())]])
    assert [e.kind for e in take(live, 3)] == ["connected", "connected", "message"]


def test_failing_connections_back_off_and_stop_on_close(connected, monkeypatch):
    live, _ = connected([])
    waits = []

    def sleep(seconds):
        waits.append(seconds)
        if len(waits) == 8:
            live.close()

    monkeypatch.setattr(live, "_sleep", sleep)
    assert list(live) == []
    assert waits[:3] == [2, 4, 8] and max(waits) == _live._MAX_BACKOFF


def test_reconnect_can_be_turned_off(connected):
    live, _ = connected([[OSError("dropped")]], reconnect=False)
    with pytest.raises(OSError):
        list(live)


def test_timeouts_are_not_failures(connected):
    class ReadTimeout(Exception):
        pass

    live, sockets = connected([[ReadTimeout("x"), delivery(raw())]])
    assert [e.kind for e in take(live, 2)] == ["connected", "message"]
    assert len(sockets) == 1


def test_close_from_another_thread_stops_it(connected):
    live, sockets = connected([[]])
    got = []
    t = threading.Thread(target=lambda: got.extend(live))
    t.start()
    time.sleep(0.1)
    live.close()
    t.join(3)
    assert not t.is_alive() and [e.kind for e in got] == ["connected"] and sockets[0].closed


def test_context_manager_closes(connected):
    live, _ = connected([[]])
    with live:
        pass
    assert live._closed


def test_an_expiring_session_reconnects_with_a_new_one(connected, monkeypatch):
    live, sockets = connected([[], [delivery(raw())]])
    client = live.client
    client._ids["chatsvc"]["exp"] = time.time() + 90  # valid, but inside the 120 s margin
    monkeypatch.setattr(client, "request", lambda *a, **k: AUTHZ)
    events = take(live, 3)
    assert [e.kind for e in events] == ["connected", "connected", "message"]
    assert client._ids["chatsvc"]["exp"] > time.time() + 3000  # signed in again
    assert len(sockets) == 2


# --- presence ---------------------------------------------------------------------
def presence_delivery():
    body = {
        "isSnapshot": True,
        "presence": [
            {
                "mri": f"8:orgid:{BOB}",
                "presence": {"availability": "Offline", "activity": "OffWork"},
            },
            {"mri": f"8:orgid:{ME}", "presence": {"availability": "Available"}},
        ],
    }
    return "3:::" + json.dumps({"id": 2, "url": "/unifiedPresence", "body": json.dumps(body)})


def test_presence_events(connected):
    live, sockets = connected([[presence_delivery()]])
    events = take(live, 3)
    assert [e.kind for e in events] == ["connected", "presence", "presence"]
    assert (events[1].sender_id, events[1].availability, events[1].activity) == (
        BOB,
        "Offline",
        "OffWork",
    )
    assert events[1].from_me is False and events[2].from_me is True
    assert '"id": 2' in sockets[0].sent[0]  # acknowledged like any delivery


def test_watching_presence_subscribes_without_publishing(monkeypatch):
    c = outnote.OutNote(lambda client_id, scope=None: "t")
    c._ids["chatsvc"] = {
        "presence": "https://ups.test",
        "url": "u",
        "headers": {},
        "exp": time.time() + 999,
    }
    calls = []
    monkeypatch.setattr(c, "request", lambda *a, **k: calls.append((a, k)) or {})
    LiveConnection(c, presence=[BOB])._watch_presence("ep", "https://trouter/f/1/")
    (method, url, body), kw = calls[0]
    assert method == "POST" and url == "https://ups.test/v1/pubsub/subscriptions/ep"
    assert body["trouterUri"] == "https://trouter/f/1/"
    assert body["subscriptionsToAdd"] == [{"mri": f"8:orgid:{BOB}", "source": "ups"}]
    assert kw["scope"] == "https://presence.teams.microsoft.com/.default"
    assert len(calls) == 1  # never touches /me/endpoints, which would publish your own status


def test_a_failed_presence_subscription_does_not_stop_the_feed(monkeypatch):
    c = outnote.OutNote(lambda client_id, scope=None: "t")
    c._ids["chatsvc"] = {
        "presence": "https://ups.test",
        "url": "u",
        "headers": {},
        "exp": time.time() + 999,
    }

    def refused(*a, **k):
        raise outnote.GraphError(400, "no")

    monkeypatch.setattr(c, "request", refused)
    LiveConnection(c, presence=[BOB])._watch_presence("ep", "https://t/")  # logs, doesn't raise


def test_no_presence_requested_means_no_subscription(monkeypatch):
    c = outnote.OutNote(lambda client_id, scope=None: "t")
    assert LiveConnection(c).presence == []
    assert c.listen.__defaults__ == (True, ())


# --- sending typing ---------------------------------------------------------------
def brokered(monkeypatch):
    c = outnote.OutNote(lambda client_id, scope=None: "tok")
    calls = []

    def request(method, path, body=None, **kw):
        calls.append({"method": method, "path": path, "body": body, **kw})
        return AUTHZ if "authz" in path else {"OriginalArrivalTime": 1}

    monkeypatch.setattr(c, "request", request)
    return c, calls


@pytest.mark.parametrize(
    ("method", "kind"),
    [("send_typing", "Control/Typing"), ("clear_typing", "Control/ClearTyping")],
)
def test_typing_is_sent_to_that_chat(monkeypatch, method, kind):
    c, calls = brokered(monkeypatch)
    getattr(c, method)(CHAT)
    call = calls[-1]
    assert call["method"] == "POST" and call["path"].endswith(f"/conversations/{CHAT}/messages")
    assert call["body"] == {
        "content": "",
        "messagetype": kind,
        "contenttype": "Application/Message",
    }
    assert call["headers"] == {"Authentication": "skypetoken=sk"}


@pytest.mark.parametrize("call", ["send_typing", "clear_typing"])
def test_typing_needs_a_brokered_client(call):
    with pytest.raises(ValueError):
        getattr(outnote.OutNote("tok"), call)(CHAT)


def test_listen_needs_a_brokered_client():
    with pytest.raises(ValueError):
        outnote.OutNote("tok").listen()


def test_listen_returns_a_connection_without_connecting(monkeypatch):
    c, calls = brokered(monkeypatch)
    assert isinstance(c.listen(), LiveConnection) and calls == []


def test_exports():
    assert outnote.LiveEvent is LiveEvent and outnote.LiveConnection is LiveConnection
