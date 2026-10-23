"""Smoke tests: public surface, URL building, pagination, PKCE, broker routing.
No network."""

import base64
import hashlib
import io
import json
import zipfile
from types import SimpleNamespace

import pytest

import outnote
from outnote import _auth, _graph


class _Resp:
    def __init__(self, body):
        self.body = body

    def read(self):
        return self.body

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


@pytest.fixture
def graph(monkeypatch):
    """Fake Graph: `graph.add(method, url_fragment, reply)`; the longest
    matching fragment wins. Sent requests land in `graph.sent`."""
    routes = []
    sent = []

    def fake(req, timeout=None):
        sent.append(req)
        hits = [r for r in routes if r[0] == req.get_method() and r[1] in req.full_url]
        if not hits:
            raise AssertionError(f"unexpected {req.get_method()} {req.full_url}")
        reply = max(hits, key=lambda r: len(r[1]))[2]
        if isinstance(reply, Exception):
            raise reply
        if reply is None:
            return _Resp(b"")
        return _Resp(reply if isinstance(reply, bytes) else json.dumps(reply).encode())

    monkeypatch.setattr(_graph, "urlopen", fake)
    note = outnote.OutNote("tok")
    return SimpleNamespace(
        note=note,
        sent=sent,
        add=lambda m, frag, reply: routes.append((m, frag, reply)),
        body=lambda i=-1: json.loads(sent[i].data) if sent[i].data else None,
    )


def test_public_surface():
    assert callable(outnote.authenticate)
    assert issubclass(outnote.GraphError, Exception)
    assert outnote.OutNote("tok").token == "tok"


def test_token_required():
    try:
        outnote.OutNote("")
    except ValueError:
        pass
    else:
        raise AssertionError("empty token should raise")


def test_callable_token_provider_is_used_per_client():
    seen = []
    note = outnote.OutNote(lambda client: seen.append(client) or "tok-" + client[:4])
    # static view: token is None when a provider is supplied
    assert note.token is None
    # the provider is invoked with the routed client id
    assert note._tp(_graph.C_ONENOTE) == "tok-1fec"
    assert _graph.C_ONENOTE in seen


def test_url_building():
    c = outnote.OutNote("tok")
    assert c._url("me/messages") == _graph.GRAPH + "/me/messages"
    assert c._url("me", {"$top": 5, "$select": None}) == _graph.GRAPH + "/me?%24top=5"
    nxt = "https://graph.microsoft.com/v1.0/me/messages?$skip=10"
    assert c._url(nxt, {"$top": 5}) == nxt + "&%24top=5"


def test_list_follows_nextlink(monkeypatch):
    pages = iter(
        [
            {"value": [1, 2], "@odata.nextLink": "https://x/next"},
            {"value": [3]},
        ]
    )
    c = outnote.OutNote("tok")
    monkeypatch.setattr(c, "get", lambda *a, **k: next(pages))
    assert c.list("me/messages") == [1, 2, 3]


def test_list_respects_limit(monkeypatch):
    c = outnote.OutNote("tok")
    monkeypatch.setattr(c, "get", lambda *a, **k: {"value": [1, 2, 3, 4, 5]})
    assert c.list("me/messages", limit=2) == [1, 2]


def test_onenote_routes_to_onenote_client(monkeypatch):
    calls = {}
    c = outnote.OutNote("tok")
    monkeypatch.setattr(
        c, "list", lambda path, *a, **k: calls.setdefault("client", k.get("client"))
    )
    c.notebooks()
    assert calls["client"] == _graph.C_ONENOTE


def test_search_builds_query_and_unwraps_hits(monkeypatch):
    captured = {}
    c = outnote.OutNote("tok")

    def fake_post(path, body, client=None):
        captured["path"] = path
        captured["body"] = body
        captured["client"] = client
        return {"value": [{"hitsContainers": [{"hits": [{"resource": {"id": "a"}}]}]}]}

    monkeypatch.setattr(c, "post", fake_post)
    out = c.search_files("budget")
    assert out == [{"id": "a"}]
    assert captured["path"] == "search/query"
    assert captured["body"]["requests"][0]["entityTypes"] == ["driveItem"]
    assert captured["body"]["requests"][0]["query"]["queryString"] == "budget"
    assert captured["client"] == _graph.C_OFFICE


def test_pkce_is_valid_s256():
    v, ch = _auth._pkce()
    expect = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=")
    assert ch == expect.decode()
    assert "=" not in v and "=" not in ch


def test_inline_query_is_encoded():
    c = outnote.OutNote("tok")
    url = c._url('me/messages?$search="Mündlichprüfung x y ö"&$top=5')
    assert " " not in url and '"' not in url
    assert url == (
        _graph.GRAPH
        + "/me/messages?$search=%22M%C3%BCndlichpr%C3%BCfung%20x%20y%20%C3%B6%22&$top=5"
    )
    # already-encoded links and stray percent signs
    assert c._url("https://x/a?q=%C3%B6") == "https://x/a?q=%C3%B6"
    assert c._url("me/x?q=50% off") == _graph.GRAPH + "/me/x?q=50%25%20off"


def test_search_keeps_hit_id(monkeypatch):
    c = outnote.OutNote("tok")
    hit = {"hitId": "AAMk1", "resource": {"subject": "Hi"}}
    monkeypatch.setattr(
        c, "post", lambda *a, **k: {"value": [{"hitsContainers": [{"hits": [hit]}]}]}
    )
    assert c.search_emails("hi")[0] == {"subject": "Hi", "id": "AAMk1"}


def _fake_token_endpoint(monkeypatch, replies):
    calls = []

    def fake(url, data):
        calls.append(data)
        return replies.pop(0)

    monkeypatch.setattr(_auth, "_post_token", fake)
    return calls


def test_broker_persists_rotated_token_atomically(tmp_path, monkeypatch):
    import os
    import stat

    store = tmp_path / ".ms_frt"
    _fake_token_endpoint(monkeypatch, [{"access_token": "a1", "refresh_token": "frt2"}])
    b = _auth._Broker("frt1", frt_path=str(store))
    assert store.read_text() == "frt1"
    assert stat.S_IMODE(os.stat(store).st_mode) == 0o600
    assert b(_graph.C_OFFICE) == "a1"
    assert store.read_text() == "frt2"
    assert [p.name for p in tmp_path.iterdir()] == [".ms_frt"]  # no temp leftovers


def test_broker_loads_store_and_clears_it_on_invalid_grant(tmp_path, monkeypatch):
    store = tmp_path / ".ms_frt"
    store.write_text("saved\n")
    calls = _fake_token_endpoint(monkeypatch, [{"error": "invalid_grant"}])
    b = _auth._Broker(frt_path=str(store))
    try:
        b(_graph.C_OFFICE)
    except outnote.AuthError as e:
        assert e.error == "invalid_grant"
    else:
        raise AssertionError("invalid_grant should raise")
    assert calls[0]["refresh_token"] == "saved"
    assert not store.exists()


def test_authenticate_reuses_store_without_sign_in(tmp_path, monkeypatch):
    store = tmp_path / ".ms_frt"
    store.write_text("saved")
    _fake_token_endpoint(monkeypatch, [{"access_token": "a", "refresh_token": "next"}])
    note = outnote.authenticate(frt_store=str(store))  # entrance is never imported
    assert note._tp(_graph.C_OFFICE) == "a"
    assert store.read_text() == "next"


def test_from_refresh_token(monkeypatch):
    _fake_token_endpoint(monkeypatch, [{"access_token": "a"}])
    note = outnote.OutNote.from_refresh_token("frt", read_only=True)
    assert note.read_only and note._tp(_graph.C_OFFICE) == "a"


def test_empty_reply_is_none_and_empty_post_has_length(graph):
    graph.add("POST", "/me/sendMail", None)
    assert graph.note.post("me/sendMail") is None
    assert graph.sent[-1].data == b""


def test_download_attachment_variants(graph, tmp_path):
    base = "/me/messages/m1/attachments/"
    graph.add(
        "GET",
        base + "a1",
        {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "k.txt",
            "contentBytes": base64.b64encode(b"hello").decode(),
        },
    )
    graph.add("GET", base + "a2", {"@odata.type": "#microsoft.graph.itemAttachment", "name": "fw"})
    graph.add("GET", base + "a2/$value", b"MIME")
    graph.add("GET", base + "a3", {"@odata.type": "#microsoft.graph.referenceAttachment"})
    n = graph.note
    assert n.download_attachment("m1", "a1") == b"hello"
    assert n.download_attachment("m1", "a1", dest=str(tmp_path)) == str(tmp_path / "k.txt")
    assert (tmp_path / "k.txt").read_bytes() == b"hello"
    assert n.download_attachment("m1", "a2") == b"MIME"
    with pytest.raises(ValueError):
        n.download_attachment("m1", "a3")


def test_download_fetches_without_bearer(graph):
    graph.add(
        "GET",
        "/me/drive/items/i1",
        {"name": "x.bin", "@microsoft.graph.downloadUrl": "https://dl.example/x?sig=1"},
    )
    graph.add("GET", "https://dl.example/x", b"\x00\x01")
    assert graph.note.download("i1") == b"\x00\x01"
    assert "Authorization" not in graph.sent[-1].headers
    assert graph.sent[0].headers["Authorization"] == "Bearer tok"


def _zip(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, text in files.items():
            z.writestr(name, text)
    return buf.getvalue()


def test_extract_text():
    docx = _zip(
        {
            "word/document.xml": "<w:document><w:body>"
            "<w:p><w:r><w:t>Lernziele</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>A &amp; B</w:t><w:tab/><w:t>C</w:t></w:r></w:p>"
            "</w:body></w:document>"
        }
    )
    assert outnote.extract_text("x.docx", docx) == "Lernziele\nA & B\tC"
    pptx = _zip(
        {
            "ppt/slides/slide10.xml": "<a:p><a:t>ten</a:t></a:p>",
            "ppt/slides/slide2.xml": "<a:p><a:t>two</a:t></a:p>",
        }
    )
    assert outnote.extract_text("deck.PPTX", pptx) == "two\n\nten"
    assert outnote.extract_text("n.md", "\ufeff# hi".encode()) == "# hi"
    assert outnote.extract_text("img.png", b"\x89PNG") == ""
    assert outnote.extract_text("broken.docx", b"not a zip") == ""


def test_html_to_text():
    teams = (
        '<p>Hi <at id="0">Anna</at>, see <attachment id="1"></attachment>'
        '<emoji alt="👍"><img src="x"></emoji></p><ul><li>one</li><li>two</li></ul>'
    )
    assert outnote.html_to_text(teams) == "Hi Anna, see 👍\n\n- one\n- two"
    page = "<html><head><title>Mathe</title></head><body><h2>Ziele</h2><p><b>Bruch</b>rechnen</p>"
    assert outnote.html_to_text(page, markdown=True) == "# Mathe\n\n## Ziele\n\n**Bruch**rechnen"


def test_onenote_group_scope_and_images(graph):
    page = (
        '<html><body><p>x</p><img alt="Ziele" src="https://graph.microsoft.com/v1.0/r/1/$value"'
        ' data-fullres-src="https://graph.microsoft.com/v1.0/r/1full/$value"'
        ' data-fullres-src-type="image/png"></body></html>'
    )
    graph.add("GET", "/groups/g1/onenote/pages/p1/content", page.encode())
    graph.add("GET", "/r/1full/$value", b"PNG")
    imgs = graph.note.page_images("p1", group_id="g1")
    assert imgs == [
        {
            "alt": "Ziele",
            "content_type": "image/png",
            "url": "https://graph.microsoft.com/v1.0/r/1full/$value",
            "data": b"PNG",
        }
    ]
    assert graph.note.page_text("p1", group_id="g1") == "x\n\n[Ziele]"


def test_pages_sorted_newest_first(graph):
    graph.add(
        "GET",
        "/me/onenote/sections/s1/pages",
        {
            "value": [
                {"id": "a", "lastModifiedDateTime": "2026-01-01"},
                {"id": "b", "lastModifiedDateTime": "2026-02-01"},
            ]
        },
    )
    assert [p["id"] for p in graph.note.pages("s1")] == ["b", "a"]


def test_photo_none_on_404(graph):
    from urllib.error import HTTPError

    graph.add("GET", "/me/photo/$value", HTTPError("u", 404, "nf", {}, io.BytesIO(b"{}")))
    assert graph.note.photo() is None


def test_try_clients_falls_back_on_403_and_remembers(monkeypatch):
    note = outnote.OutNote(lambda c: "t-" + c)
    calls = []

    def call(client):
        calls.append(client)
        if client == _graph.C_OFFICE:
            raise outnote.GraphError(403, "nope")
        return "ok"

    assert note._try_clients("presence", call) == "ok"
    assert note._try_clients("presence", call) == "ok"
    assert calls == [_graph.C_OFFICE, _graph.C_TEAMS, _graph.C_TEAMS]


def test_batch_chunks_maps_errors_and_retries_throttled(graph, monkeypatch):
    monkeypatch.setattr(_graph.time, "sleep", lambda s: None)
    n = graph.note
    replies = iter(
        [
            {
                "responses": [{"id": str(i), "status": 200, "body": {"n": i}} for i in range(1, 20)]
                + [{"id": "0", "status": 429, "headers": {"Retry-After": "1"}}]
            },
            {
                "responses": [
                    {"id": "20", "status": 404, "body": {"error": {"message": "gone"}}},
                    {"id": "21", "status": 200, "body": {"n": 21}},
                ]
            },
            {"responses": [{"id": "0", "status": 200, "body": {"n": 0}}]},
        ]
    )
    sent = []
    monkeypatch.setattr(
        n, "post", lambda path, body, client=None: sent.append(body) or next(replies)
    )
    reqs = [f"me/messages/{i}" for i in range(21)] + [
        {"url": _graph.GRAPH + "/me/sendMail", "method": "post", "body": {"x": 1}}
    ]
    out = n.batch(reqs)
    assert [len(b["requests"]) for b in sent] == [20, 2, 1]
    assert sent[1]["requests"][1] == {
        "id": "21",
        "method": "POST",
        "url": "/me/sendMail",
        "body": {"x": 1},
        "headers": {"Content-Type": "application/json"},
    }
    assert out[0] == {"n": 0} and out[19] == {"n": 19} and out[21] == {"n": 21}
    assert isinstance(out[20], outnote.GraphError) and out[20].status == 404


def test_delta_follows_pages_and_reuses_token(graph):
    graph.add(
        "GET",
        "/me/mailFolders/inbox/messages/delta",
        {"value": [1], "@odata.nextLink": "https://g/p2"},
    )
    graph.add("GET", "https://g/p2", {"value": [2], "@odata.deltaLink": "https://g/d1"})
    graph.add(
        "GET",
        "https://g/d1",
        {"value": [{"id": "x", "@removed": {}}], "@odata.deltaLink": "https://g/d2"},
    )
    items, token = graph.note.delta("me/mailFolders/inbox/messages")
    assert (items, token) == ([1, 2], "https://g/d1")
    assert graph.note.delta("ignored", token=token) == (
        [{"id": "x", "@removed": {}}],
        "https://g/d2",
    )
