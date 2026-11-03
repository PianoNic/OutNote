"""Identity from the token, directory reads around per-client throttling, compact mail lists, the
retry budget and parallel batches. No network."""

import base64
import io
import json
import threading
import time
from urllib.error import HTTPError

import pytest

import outnote
from outnote import _base

ME = "11111111-1111-1111-1111-111111111111"


def jwt(**claims):
    body = base64.urlsafe_b64encode(json.dumps(claims).encode()).rstrip(b"=").decode()
    return f"h.{body}.s"


def test_ids_come_from_the_token_without_a_request(monkeypatch):
    c = outnote.OutNote(lambda client, scope=None: jwt(oid=ME, tid="t"))
    monkeypatch.setattr(c, "request", lambda *a, **k: pytest.fail("no request expected"))
    assert (c._me_id(), c._tenant_id()) == (ME, "t")


def test_an_opaque_token_still_finds_the_id_with_a_request(monkeypatch):
    c = outnote.OutNote("opaque")
    monkeypatch.setattr(c, "request", lambda *a, **k: {"id": ME})
    assert c._me_id() == ME


def directory_client(monkeypatch, answer):
    c = outnote.OutNote(lambda client_id, scope=None: "t")
    seen = []

    def request(method, path, body=None, params=None, raw=False, client=None, retry=True, **kw):
        seen.append((client, retry))
        return answer(client)

    monkeypatch.setattr(c, "request", request)
    return c, seen


def test_directory_moves_to_the_next_client_when_throttled(monkeypatch):
    def answer(client):
        if client == _base.C_OFFICE:
            raise outnote.GraphError(429, "busy", retry_after=20)
        return {"id": "me"}

    c, seen = directory_client(monkeypatch, answer)
    assert c.me() == {"id": "me"}
    assert seen == [(_base.C_OFFICE, False), (_base.C_TEAMS, False)]  # no waiting on the way


def test_directory_uses_only_office_when_it_answers(monkeypatch):
    c, seen = directory_client(monkeypatch, lambda client: {"id": "me"})
    c.me()
    assert seen == [(_base.C_OFFICE, False)]


def test_directory_last_client_waits_and_real_errors_stop(monkeypatch):
    def busy(client):
        raise outnote.GraphError(429, "busy")

    c, seen = directory_client(monkeypatch, busy)
    with pytest.raises(outnote.GraphError):
        c.me()
    assert seen[-1] == (_base.C_OFFICEHOME, True) and len(seen) == 3

    def missing(client):
        raise outnote.GraphError(404, "no such user")

    c, seen = directory_client(monkeypatch, missing)
    with pytest.raises(outnote.GraphError):
        c.user("nobody")
    assert len(seen) == 1


def test_static_token_has_one_directory_identity(monkeypatch):
    c = outnote.OutNote("tok")
    seen = []
    monkeypatch.setattr(
        c, "request", lambda *a, client=None, retry=True, **k: seen.append(retry) or {}
    )
    c.me()
    assert seen == [True]


def test_older_client_names_still_work():
    assert outnote.C_ONENOTE == outnote.C_TEAMS and outnote.C_SITES == outnote.C_OFFICEHOME


def test_emails_are_compact_unless_asked(monkeypatch):
    seen = []
    c = outnote.OutNote("tok")
    monkeypatch.setattr(c, "list", lambda path, params=None, **k: seen.append(params) or [])
    c.emails()
    fields = seen[-1]["$select"].split(",")
    assert "bodyPreview" in fields and "body" not in fields and "toRecipients" not in fields
    c.emails(select="id")
    assert seen[-1]["$select"] == "id"
    c.emails(select="*")
    assert seen[-1]["$select"] is None


def test_mail_folder_tree_uses_one_batch_per_depth(monkeypatch):
    c = outnote.OutNote("tok")
    monkeypatch.setattr(
        c,
        "list",
        lambda path, params=None, **k: [
            {"id": "a", "displayName": "A", "childFolderCount": 1},
            {"id": "b", "displayName": "B", "childFolderCount": 1},
        ],
    )
    calls = []

    def batch(paths):
        calls.append(paths)
        if len(calls) == 1:
            return [{"value": [{"id": "a1", "displayName": "A1"}]}, {"value": []}]
        return []

    monkeypatch.setattr(c, "batch", batch)
    tree = c.mail_folders(include_children=True)
    assert [f["path"] for f in tree] == ["A", "B", "A/A1"]
    assert len(calls[0]) == 2


def test_throttled_call_gives_up_at_the_budget(monkeypatch):
    slept = []

    def always_429(req, timeout):
        raise HTTPError(req.full_url, 429, "x", {"Retry-After": "40"}, io.BytesIO(b"{}"))

    monkeypatch.setattr(_base, "urlopen", always_429)
    monkeypatch.setattr(_base.time, "sleep", slept.append)
    with pytest.raises(outnote.GraphError) as e:
        outnote.OutNote("tok").get("messages")
    assert slept == [40.0]  # a second 40 s wait would pass the 60 s budget
    assert e.value.status == 429 and e.value.retry_after == 40.0


def test_fetch_failures_are_graph_errors(monkeypatch):
    def not_acceptable(req, timeout):
        raise HTTPError(req.full_url, 406, "Not Acceptable", {}, io.BytesIO(b""))

    monkeypatch.setattr(_base, "urlopen", not_acceptable)
    with pytest.raises(outnote.GraphError) as e:
        _base.fetch("https://x/y")
    assert e.value.status == 406


def test_batches_run_in_parallel_and_keep_order(monkeypatch):
    c = outnote.OutNote("tok")
    lock = threading.Lock()
    active = {"now": 0, "peak": 0}

    def post(path, body, client=None):
        with lock:
            active["now"] += 1
            active["peak"] = max(active["peak"], active["now"])
        time.sleep(0.05)
        with lock:
            active["now"] -= 1
        return {
            "responses": [{"id": r["id"], "status": 200, "body": r["id"]} for r in body["requests"]]
        }

    monkeypatch.setattr(c, "post", post)
    got = c.batch([f"me/x{i}" for i in range(45)])
    assert got == [str(i) for i in range(45)]
    assert active["peak"] > 1
