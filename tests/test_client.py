"""Unit tests for the full client: transport, writes, Teams, files, text.
No network - requests are captured instead of sent."""

import base64
import io
import json
import zipfile
from datetime import datetime, timezone

import httpx
import pytest

import outnote
from outnote import _base, _cli
from outnote._text import extract_text, html_to_text, message_text

ME = "11111111-1111-1111-1111-111111111111"
DIANA = "22222222-2222-2222-2222-222222222222"


class Recorder:
    """Stands in for `Transport.request`: records calls, answers from a
    `(method, path-prefix) -> response` table (first match wins)."""

    def __init__(self):
        self.calls = []
        self.answers = {}

    def __call__(self, method, path, body=None, params=None, raw=False, client=None, **kw):
        self.calls.append(
            {
                "method": method,
                "path": path,
                "body": body,
                "params": params,
                "raw": raw,
                "client": client,
                **kw,
            }
        )
        short = path.removeprefix(_base.GRAPH + "/")  # list() passes full URLs
        for (m, prefix), resp in self.answers.items():
            if m == method and (path.startswith(prefix) or short.startswith(prefix)):
                return resp(path, body, params) if callable(resp) else resp
        return None


@pytest.fixture
def note(monkeypatch):
    c = outnote.OutNote("tok")
    rec = Recorder()
    monkeypatch.setattr(c, "request", rec)
    c._ids.update(me=ME, tenant="tenant-1")
    c.rec = rec
    return c


class FakeResp:
    def __init__(self, body):
        self.body = body

    def read(self):
        return self.body

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


# --- transport ----------------------------------------------------------------
def test_send_handles_empty_body(monkeypatch):
    monkeypatch.setattr(_base, "urlopen", lambda req, timeout: FakeResp(b""))
    assert outnote.OutNote("tok").post("me/sendMail", {"message": {}}) is None


def test_request_methods_and_bodies(monkeypatch):
    seen = []

    def fake_urlopen(req, timeout):
        seen.append(req)
        return FakeResp(b'{"ok": 1}')

    monkeypatch.setattr(_base, "urlopen", fake_urlopen)
    c = outnote.OutNote("tok")
    assert c.patch("me/messages/1", {"isRead": True}) == {"ok": 1}
    c.put("me/drive/root:/a.txt:/content", b"hi", content_type="text/plain")
    c.delete("me/events/1")
    assert [r.get_method() for r in seen] == ["PATCH", "PUT", "DELETE"]
    assert json.loads(seen[0].data) == {"isRead": True}
    assert seen[0].headers["Content-type"] == "application/json"
    assert seen[1].data == b"hi" and seen[1].headers["Content-type"] == "text/plain"
    assert seen[2].data is None


def test_raw_bytes(monkeypatch):
    monkeypatch.setattr(_base, "urlopen", lambda req, timeout: FakeResp(b"\x89PNG"))
    assert outnote.OutNote("tok").get("me/photo/$value", raw="bytes") == b"\x89PNG"


def test_inline_query_is_percent_encoded():
    c = outnote.OutNote("tok")
    url = c._url('me/messages?$search="Mündlichprüfung x y"&$top=5')
    assert " " not in url and '"' not in url
    assert url == (
        _base.GRAPH + "/me/messages?$search=%22M%C3%BCndlichpr%C3%BCfung%20x%20y%22&$top=5"
    )
    # an already-encoded nextLink passes through unchanged
    nxt = "https://graph.microsoft.com/v1.0/me/messages?%24skip=10&%24skiptoken=a%2Bb"
    assert c._url(nxt) == nxt


def test_batch_maps_results_and_retries_throttled(monkeypatch):
    c = outnote.OutNote("tok")
    rounds = []

    def fake_post(path, body, client=None, **kw):
        assert path == "$batch"
        rounds.append([r["id"] for r in body["requests"]])
        out = []
        for r in body["requests"]:
            if r["id"] == "1" and len(rounds) == 1:
                out.append({"id": "1", "status": 429, "headers": {"Retry-After": "0"}})
            elif r["id"] == "2":
                out.append({"id": "2", "status": 404, "body": {"error": {"message": "nope"}}})
            else:
                out.append({"id": r["id"], "status": 200, "body": {"url": r["url"]}})
        return {"responses": out}

    monkeypatch.setattr(c, "post", fake_post)
    monkeypatch.setattr(_base.time, "sleep", lambda s: None)
    res = c.batch(["me", {"url": "me/events"}, "/me/nothing"])
    assert res[0] == {"url": "/me"} and res[1] == {"url": "/me/events"}
    assert isinstance(res[2], outnote.GraphError) and res[2].status == 404
    assert rounds == [["0", "1", "2"], ["1"]]


def test_batch_chunks_by_20(monkeypatch):
    c = outnote.OutNote("tok")
    sizes = []

    def fake_post(path, body, client=None, **kw):
        sizes.append(len(body["requests"]))
        return {"responses": [{"id": r["id"], "status": 200, "body": 1} for r in body["requests"]]}

    monkeypatch.setattr(c, "post", fake_post)
    assert c.batch(["me"] * 45) == [1] * 45
    assert sizes == [20, 20, 5]


def test_delta_follows_to_deltalink(monkeypatch):
    c = outnote.OutNote("tok")
    pages = iter(
        [
            {"value": [1], "@odata.nextLink": "https://x/n"},
            {"value": [2], "@odata.deltaLink": "https://x/d"},
        ]
    )
    urls = []
    monkeypatch.setattr(c, "get", lambda url, **k: urls.append(url) or next(pages))
    assert c.delta("me/mailFolders/inbox/messages") == ([1, 2], "https://x/d")
    assert urls[0].endswith("/me/mailFolders/inbox/messages/delta")


# --- from_refresh_token (OutNote API; the Broker lives in entrance now) ------
def test_from_refresh_token():
    note = outnote.OutNote.from_refresh_token("frt")
    assert note.token is None and callable(note._tp)


# --- mail -----------------------------------------------------------------------
def test_send_mail_payload(note, tmp_path):
    f = tmp_path / "a.txt"
    f.write_bytes(b"hello")
    note.send_mail("x@y.ch", "Hi", "Body", cc=["c@y.ch"], attachments=[str(f), ("b.bin", b"\0")])
    call = note.rec.calls[-1]
    assert call["method"] == "POST" and call["path"] == "me/sendMail"
    msg = call["body"]["message"]
    assert msg["toRecipients"] == [{"emailAddress": {"address": "x@y.ch"}}]
    assert msg["ccRecipients"][0]["emailAddress"]["address"] == "c@y.ch"
    assert msg["body"] == {"contentType": "Text", "content": "Body"}
    assert [a["name"] for a in msg["attachments"]] == ["a.txt", "b.bin"]
    assert base64.b64decode(msg["attachments"][0]["contentBytes"]) == b"hello"


def test_mail_triage(note):
    note.mark_read("m1")
    note.flag("m1", "complete")
    note.move_email("m1", "archive")
    note.delete_email("m1")
    note.reply("m1", "ok", reply_all=True)
    note.forward("m1", ["a@b.ch"], "fyi")
    c = note.rec.calls
    assert (c[0]["method"], c[0]["body"]) == ("PATCH", {"isRead": True})
    assert c[1]["body"] == {"flag": {"flagStatus": "complete"}}
    assert c[2]["path"] == "me/messages/m1/move" and c[2]["body"] == {"destinationId": "archive"}
    assert c[3]["method"] == "DELETE"
    assert c[4]["path"] == "me/messages/m1/replyAll"
    assert c[5]["body"]["toRecipients"][0]["emailAddress"]["address"] == "a@b.ch"


def test_unread_filter(note, monkeypatch):
    seen = {}
    monkeypatch.setattr(note, "list", lambda path, params=None, **k: seen.update(params) or [])
    note.unread_emails()
    assert seen["$filter"] == "isRead eq false"


def test_download_attachment_decodes(note):
    note.rec.answers[("GET", "me/messages/m/attachments/a")] = {
        "contentBytes": base64.b64encode(b"PDF").decode()
    }
    assert note.download_attachment("m", "a") == b"PDF"


# --- calendar -------------------------------------------------------------------
def test_create_event_payload(note):
    note.create_event(
        "Test",
        datetime(2026, 1, 1, 10, tzinfo=timezone.utc),
        "2026-01-01T11:00:00",
        attendees="a@b.ch",
        online=True,
    )
    body = note.rec.calls[-1]["body"]
    assert body["start"] == {"dateTime": "2026-01-01T10:00:00", "timeZone": "UTC"}
    assert body["end"] == {"dateTime": "2026-01-01T11:00:00", "timeZone": "UTC"}
    assert body["attendees"][0]["emailAddress"]["address"] == "a@b.ch"
    assert body["isOnlineMeeting"] is True


def test_free_busy(note):
    note.rec.answers[("POST", "me/calendar/getSchedule")] = {"value": [{"scheduleId": "a"}]}
    res = note.free_busy("a@b.ch", "2026-01-01T08:00:00Z", "2026-01-01T17:00:00Z")
    assert res == [{"scheduleId": "a"}]
    body = note.rec.calls[-1]["body"]
    assert body["schedules"] == ["a@b.ch"] and body["availabilityViewInterval"] == 30
    assert body["startTime"] == {"dateTime": "2026-01-01T08:00:00", "timeZone": "UTC"}


# --- teams ----------------------------------------------------------------------
def test_mentions_in_place_and_prepended(note):
    mentions = [{"id": DIANA, "name": "Diana"}, {"id": ME, "name": "Me"}]
    note.send_chat_message("c1", "hi @Diana <3", mentions=mentions)
    body = note.rec.calls[-1]["body"]
    assert body["body"]["contentType"] == "html"
    assert body["body"]["content"] == '<at id="1">Me</at> hi <at id="0">Diana</at> &lt;3'
    assert body["mentions"][0]["mentioned"]["user"]["id"] == DIANA


def test_plain_chat_message(note):
    note.send_chat_message("c1", "hello")
    assert note.rec.calls[-1]["body"] == {"body": {"contentType": "text", "content": "hello"}}


def test_create_chat_one_on_one(note, monkeypatch):
    monkeypatch.setattr(note, "resolve_user", lambda u: {"id": DIANA})
    note.create_chat("diana@x.ch")
    body = note.rec.calls[-1]["body"]
    assert body["chatType"] == "oneOnOne"
    binds = [m["user@odata.bind"] for m in body["members"]]
    assert binds == [f"{_base.GRAPH}/users('{ME}')", f"{_base.GRAPH}/users('{DIANA}')"]


def test_send_to_creates_then_sends(note, monkeypatch):
    monkeypatch.setattr(note, "resolve_user", lambda u: {"id": DIANA})
    note.rec.answers[("POST", "chats")] = lambda p, b, q: (
        {"id": "chat-1"} if p == "chats" else {"id": "msg"}
    )
    assert note.send_to("diana@x.ch", "yo") == {"id": "msg"}
    assert note.rec.calls[-1]["path"] == "chats/chat-1/messages"


def _chat(cid, created, sender, read):
    return {
        "id": cid,
        "lastMessagePreview": {"createdDateTime": created, "from": {"user": {"id": sender}}},
        "viewpoint": {"lastMessageReadDateTime": read},
    }


def test_unread_chats(note, monkeypatch):
    chats = [
        _chat("a", "2026-02", DIANA, "2026-01"),
        _chat("b", "2026-02", ME, "2026-01"),
        _chat("c", "2026-01", DIANA, "2026-02"),
    ]
    monkeypatch.setattr(note, "chats", lambda top=50: chats)
    assert [c["id"] for c in note.unread_chats()] == ["a"]


def test_find_chat(note, monkeypatch):
    monkeypatch.setattr(note, "resolve_user", lambda u: {"id": DIANA})
    one = [{"id": "x", "members": [{"userId": ME}, {"userId": DIANA}]}]
    monkeypatch.setattr(note, "chats", lambda top=200, chat_type=None: one)
    assert note.find_chat("diana")["id"] == "x"


def test_chat_lifecycle_paths(note):
    note.react("c", "m", "❤️")
    note.delete_chat_message("c", "m")
    note.mark_chat_read("c")
    note.pin_message("c", "m")
    note.hide_chat("c")
    c = note.rec.calls
    assert c[0]["path"] == "chats/c/messages/m/setReaction"
    assert c[0]["body"] == {"reactionType": "❤️"}
    assert c[1]["path"] == f"users/{ME}/chats/c/messages/m/softDelete"
    assert c[2]["body"] == {"user": {"id": ME, "tenantId": "tenant-1"}}
    assert c[3]["body"]["message@odata.bind"].endswith("/chats/c/messages/m")
    assert c[4]["path"] == "chats/c/hideForUser"


def test_message_files():
    msg = {
        "attachments": [
            {"contentType": "reference", "name": "a.docx", "contentUrl": "u"},
            {"contentType": "messageReference"},
        ]
    }
    assert outnote.OutNote.message_files(msg) == [{"id": None, "name": "a.docx", "url": "u"}]


# --- directory ------------------------------------------------------------------
def test_resolve_user_by_upn(note):
    note.rec.answers[("GET", "users/diana")] = {"id": DIANA, "displayName": "Diana"}
    assert note.resolve_user("diana@x.ch")["id"] == DIANA


def test_resolve_user_ambiguous_name(note):
    note.rec.answers[("GET", "users")] = {
        "value": [{"id": "1", "displayName": "Ann A"}, {"id": "2", "displayName": "Ann B"}]
    }
    with pytest.raises(LookupError):
        note.resolve_user("Ann")


def test_find_users_sends_consistency_header(note):
    note.rec.answers[("GET", "users")] = {"value": []}
    note.find_users("Diana")
    call = note.rec.calls[-1]
    assert call["headers"] == {"ConsistencyLevel": "eventual"}
    assert "displayName%3ADiana" in call["path"]


# --- search ---------------------------------------------------------------------
def test_search_keeps_hit_id(monkeypatch):
    c = outnote.OutNote("tok")
    hit = {"hitId": "AAMk", "summary": "s", "resource": {"subject": "x"}}
    monkeypatch.setattr(
        c, "post", lambda *a, **k: {"value": [{"hitsContainers": [{"hits": [hit]}]}]}
    )
    assert c.search_emails("x") == [{"subject": "x", "id": "AAMk", "_summary": "s"}]


# --- files ----------------------------------------------------------------------
def test_simple_upload(note):
    note.upload(b"data", "Folder/a.txt")
    call = note.rec.calls[-1]
    assert call["method"] == "PUT"
    assert call["path"].startswith("me/drive/root:/Folder/a.txt:/content")
    assert call["body"] == b"data"


def test_large_upload_uses_session(note, monkeypatch):
    note.rec.answers[("POST", "me/drive/root:/big.bin:/createUploadSession")] = {
        "uploadUrl": "https://up"
    }
    chunks = []
    monkeypatch.setattr(
        note, "_upload_chunks", lambda url, data: chunks.append((url, len(data))) or {"id": "i"}
    )
    assert note.upload(b"x" * (5 * 1024 * 1024), "big.bin") == {"id": "i"}
    assert chunks == [("https://up", 5 * 1024 * 1024)]


def test_onenote_group_paths_and_images(note):
    img = "https://graph.microsoft.com/v1.0/me/onenote/resources/r1/$value"
    html = (
        f'<html><body><p>Lern<b>ziele</b></p><img src="{img}" data-src-type="image/png"/>'
        "</body></html>"
    )
    note.rec.answers[("GET", "groups/g/onenote/pages/p/content")] = html
    note.rec.answers[("GET", img)] = b"IMG"
    assert note.page_text("p", group_id="g") == "Lernziele"
    imgs = note.page_images("p", group_id="g")
    assert imgs[0]["data"] == b"IMG" and imgs[0]["content_type"] == "image/png"
    assert all(c["client"] == outnote.C_ONENOTE for c in note.rec.calls)


def test_create_page_xhtml(note):
    note.create_page("s", title="T", body="<p>x</p>")
    call = note.rec.calls[-1]
    assert call["content_type"] == "application/xhtml+xml"
    assert "<title>T</title>" in call["body"]


# --- text -----------------------------------------------------------------------
def test_html_to_text():
    html = (
        '<div><p>Hello&nbsp;<at id="0">Diana</at></p><ul><li>one</li><li>two</li></ul>'
        '<attachment id="x"></attachment></div>'
    )
    assert html_to_text(html) == "Hello Diana\n\n- one\n\n- two"


def test_message_text_plain():
    assert message_text({"body": {"contentType": "text", "content": "a &amp; b"}}) == "a & b"


def _zip(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for n, d in files.items():
            z.writestr(n, d)
    return buf.getvalue()


def test_extract_text_docx_pptx_unknown():
    docx = _zip(
        {
            "word/document.xml": "<w:document><w:p><w:r><w:t>Hallo</w:t></w:r></w:p>"
            "<w:p><w:t>Welt &amp; Co</w:t></w:p></w:document>"
        }
    )
    assert extract_text("a.docx", docx) == "Hallo\nWelt & Co"
    pptx = _zip(
        {
            "ppt/slides/slide2.xml": "<a:p><a:t>B</a:t></a:p>",
            "ppt/slides/slide1.xml": "<a:p><a:t>A</a:t></a:p>",
        }
    )
    assert extract_text("deck.pptx", pptx) == "A\n\nB"
    assert extract_text("x.txt", b"plain") == "plain"
    assert extract_text("pic.png", b"\x89PNG") == ""
    assert extract_text("broken.docx", b"not a zip") == ""


# --- cli ------------------------------------------------------------------------
def test_cli_text_rendering(monkeypatch, capsys):
    fake = outnote.OutNote("tok")
    monkeypatch.setattr(fake, "me", lambda: {"displayName": "N", "mail": "n@x.ch"})
    monkeypatch.setattr(_cli, "_client", lambda device=False, browser=False: fake)
    assert _cli.main(["--text", "me"]) == 0
    assert capsys.readouterr().out.strip() == "N <n@x.ch>"


def test_cli_not_signed_in_prints_help_not_traceback(monkeypatch, capsys):
    from entrance import LoginFailed

    def no_creds(device=False, browser=False):
        raise LoginFailed("nope")

    monkeypatch.setattr(_cli, "_client", no_creds)
    assert _cli.main(["chats", "--unread"]) == 2
    err = capsys.readouterr().err
    assert "Not signed in" in err and "outnote signin" in err
    assert "Traceback" not in err


def test_cli_signin_defaults_to_device_code(monkeypatch, capsys):
    fake = outnote.OutNote("tok")
    monkeypatch.setattr(fake, "me", lambda: {"displayName": "N", "userPrincipalName": "n@x.ch"})
    seen = {}
    monkeypatch.setattr(
        _cli,
        "_client",
        lambda device=False, browser=False: seen.update(device=device, browser=browser) or fake,
    )
    assert _cli.main(["login"]) == 0  # alias for signin
    assert seen == {"device": True, "browser": False}  # device code by default
    assert "Signed in as N <n@x.ch>" in capsys.readouterr().out


def test_cli_signin_browser_flag_uses_browser(monkeypatch, capsys):
    fake = outnote.OutNote("tok")
    monkeypatch.setattr(fake, "me", lambda: {"displayName": "N", "userPrincipalName": "n@x.ch"})
    seen = {}
    monkeypatch.setattr(
        _cli,
        "_client",
        lambda device=False, browser=False: seen.update(device=device, browser=browser) or fake,
    )
    assert _cli.main(["signin", "--browser"]) == 0
    assert seen == {"device": False, "browser": True}


# --- fixes found by watching the Teams web client --------------------------------
def test_education_uses_assignments_api_with_scoped_token(note):
    note.rec.answers[("GET", "https://assignments.edu.cloud.microsoft")] = {"value": [{"id": "a"}]}
    assert note.assignments() == [{"id": "a"}]
    call = note.rec.calls[-1]
    assert call["path"].startswith("https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work")
    assert call["scope"] == "api://8f348934-64be-4bb2-bc16-c54c96789f43/.default"


def test_scoped_token_needs_a_broker():
    seen = []
    note = outnote.OutNote(lambda client, scope=None: seen.append(scope) or "t")
    assert note._token(outnote.C_OFFICE, "api://x/.default") == "t" and seen == ["api://x/.default"]
    with pytest.raises(ValueError):
        outnote.OutNote("static")._token(outnote.C_OFFICE, "api://x/.default")


def test_presence_routes_to_officehome(note):
    note.presence()
    assert note.rec.calls[-1]["client"] == outnote.C_SITES


def test_meeting_thread_id_from_join_url(note):
    url = (
        "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0"
        "?context=%7b%22Tid%22%3a%22t%22%7d"
    )
    event = {"onlineMeeting": {"joinUrl": url}}
    assert outnote.OutNote.meeting_thread_id(event) == "19:meeting_abc@thread.v2"
    note.rec.answers[("GET", "https://graph.microsoft.com/v1.0/me/chats/")] = {"value": []}
    note.meeting_chat_messages(event)
    assert "me/chats/19:meeting_abc@thread.v2/messages" in note.rec.calls[-1]["path"]


# --- shared keep-alive client ---------------------------------------------------
def _mock_shared_client(monkeypatch, handler):
    import entrance._http as shared

    c = httpx.Client(transport=httpx.MockTransport(handler))
    monkeypatch.setattr(shared, "_client", c)
    return c


def test_requests_share_one_client_and_follow_redirects(monkeypatch):
    seen = []

    def handler(req):
        seen.append((req.method, str(req.url), req.headers.get("authorization")))
        if req.url.host == "graph.microsoft.com" and req.url.path.endswith("/content"):
            return httpx.Response(302, headers={"location": "https://files.example/x"})
        return httpx.Response(200, content=b"bytes")

    _mock_shared_client(monkeypatch, handler)
    c = outnote.OutNote("tok")
    assert c.get("me/drive/items/1/content", raw="bytes") == b"bytes"
    assert seen[0][2] == "Bearer tok"
    assert seen[1][1] == "https://files.example/x"
    assert seen[1][2] is None  # the Bearer token is not forwarded to another host


def test_http_errors_keep_graph_error_and_retry_semantics(monkeypatch):
    calls = []

    def handler(req):
        calls.append(1)
        if len(calls) == 1:
            return httpx.Response(503, headers={"Retry-After": "0"}, json={})
        return httpx.Response(404, json={"error": {"message": "gone"}})

    _mock_shared_client(monkeypatch, handler)
    with pytest.raises(outnote.GraphError, match="404.*gone"):
        outnote.OutNote("tok").get("me/messages/1")
    assert len(calls) == 2


def test_connection_failure_is_a_url_error(monkeypatch):
    def handler(req):
        raise httpx.ConnectTimeout("no route")

    _mock_shared_client(monkeypatch, handler)
    with pytest.raises(OSError):
        outnote.OutNote("tok").get("me")
