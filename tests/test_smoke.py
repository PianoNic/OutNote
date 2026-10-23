"""Smoke tests: public surface, URL building, pagination, PKCE, broker routing.
No network."""

import base64
import hashlib

import outnote
from outnote import _auth, _graph


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
