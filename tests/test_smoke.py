"""Smoke tests: public surface, URL building, pagination, PKCE, broker routing.
No network."""

import outnote
from outnote import _graph


def test_public_surface():
    assert callable(outnote.authenticate)
    assert issubclass(outnote.GraphError, Exception)
    assert outnote.OutNote("tok").token == "tok"


def test_push_surface_is_re_exported():
    assert callable(outnote.push_login)
    assert issubclass(outnote.PushError, Exception)
    for state in ("NUMBER_MATCHING", "PENDING", "APPROVED", "DENIED", "TIMED_OUT"):
        assert isinstance(getattr(outnote, state), str)


def test_authenticate_forwards_push_to_foci_login(monkeypatch):
    from outnote import _auth

    captured = {}

    class FakeBroker:
        refresh_token = "frt"

        def __call__(self, client_id, scope=None):
            return "token"

    monkeypatch.setattr(
        _auth, "foci_login", lambda *a, **k: captured.update(a=a, k=k) or FakeBroker()
    )
    cb = lambda ev: None  # noqa: E731
    note = _auth.authenticate(
        "you@school.ch", "pw", push=True, on_push=cb, push_timeout=30, push_interval=1
    )

    assert isinstance(note, outnote.OutNote)
    assert captured["k"]["push"] is True
    assert captured["k"]["on_push"] is cb
    assert captured["k"]["push_timeout"] == 30
    assert captured["k"]["push_interval"] == 1


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
