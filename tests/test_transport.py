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


# --- download links -----------------------------------------------------------
LINK = "@microsoft.graph.downloadUrl"


class FakeHttp:
    """Answers `/content` with a redirect, like Graph."""

    def __init__(self, status=302):
        self.status, self.urls = status, []

    def request(self, method, url, **kw):
        from entrance._http import Headers, Response

        self.urls.append(url)
        return Response(self.status, Headers(location="https://dl.test/redirected"), b"")


def files_client(monkeypatch, item, status=302):
    c = outnote.OutNote("tok")
    seen = []
    monkeypatch.setattr(c, "get", lambda path, *a, **k: seen.append(path) or item)
    http = FakeHttp(status)
    monkeypatch.setattr(_base, "http_client", lambda: http)
    monkeypatch.setattr("outnote._files.fetch", lambda url, **k: url.encode())
    c.seen, c.http = seen, http
    return c


def test_the_link_on_the_item_is_used_when_present(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", "file": {}, LINK: "https://dl.test/direct"})
    assert c.download_url("x") == "https://dl.test/direct"
    assert c.http.urls == []


def test_a_file_without_a_link_falls_back_to_content(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", "name": "a.zip", "file": {}})
    assert c.download("x") == b"https://dl.test/redirected"
    assert c.http.urls[0].endswith("/me/drive/items/x/content")


def test_download_shared_falls_back_to_content(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", "file": {}})
    assert c.download_shared("https://share.test/f") == b"https://dl.test/redirected"
    assert "/shares/u!" in c.http.urls[0] and c.http.urls[0].endswith("/driveItem/content")


def test_folders_get_a_clear_error_and_no_content_request(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", "folder": {"childCount": 1}})
    assert c.download_url("x") is None
    with pytest.raises(ValueError, match="folder"):
        c.download("x")
    with pytest.raises(ValueError, match="folder"):
        c.download_shared("https://share.test/f")
    assert c.http.urls == []


def test_a_file_with_no_link_at_all_is_a_graph_error(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", "file": {}}, status=404)
    with pytest.raises(outnote.GraphError):
        c.download("x")
    with pytest.raises(outnote.GraphError):
        c.download_shared("https://share.test/f")


def test_a_shared_item_is_found_in_its_own_drive(monkeypatch):
    shared = {
        "id": "local",
        "remoteItem": {"id": "real", "parentReference": {"driveId": "theirs"}},
        LINK: "https://dl.test/direct",
    }
    c = files_client(monkeypatch, shared)
    c.download_url(shared)
    assert c.seen == ["drives/theirs/items/real"]


def test_items_and_ids_both_work(monkeypatch):
    c = files_client(monkeypatch, {"id": "x", LINK: "https://dl.test/direct"})
    c.download_url("x")
    c.download_url("x", drive_id="d")
    c.download_url({"id": "x", "parentReference": {"driveId": "d2"}})
    c.download_url({"id": "x"})
    assert c.seen == [
        "me/drive/items/x",
        "drives/d/items/x",
        "drives/d2/items/x",
        "me/drive/items/x",
    ]


# --- suggest_recipients --------------------------------------------------------
def person(name, address):
    return {"id": name, "displayName": name, "scoredEmailAddresses": [{"address": address}]}


def people_client(monkeypatch, people=(), directory=()):
    c = outnote.OutNote("tok")
    calls = {"people": [], "users": []}
    monkeypatch.setattr(
        c,
        "get",
        lambda path, params=None, **k: calls["people"].append(params) or {"value": list(people)},
    )
    monkeypatch.setattr(
        c, "find_users", lambda text, top: calls["users"].append(text) or list(directory)
    )
    c.calls = calls
    return c


def test_suggestions_put_your_contacts_first_and_skip_the_directory_when_full(monkeypatch):
    c = people_client(monkeypatch, [person("Ann", "ann@x.ch"), person("Abe", "abe@x.ch")])
    got = c.suggest_recipients("a", top=2)
    assert got == [
        {"name": "Ann", "address": "ann@x.ch", "id": "Ann"},
        {"name": "Abe", "address": "abe@x.ch", "id": "Abe"},
    ]
    assert c.calls["people"][0]["$search"] == '"a"' and c.calls["users"] == []


def test_suggestions_fill_up_from_the_directory_without_duplicates(monkeypatch):
    directory = [
        {"id": "1", "displayName": "Ann again", "mail": "ANN@x.ch"},
        {"id": "2", "displayName": "Ada", "mail": None, "userPrincipalName": "ada@x.ch"},
        {"id": "3", "displayName": "No address"},
    ]
    c = people_client(monkeypatch, [person("Ann", "ann@x.ch")], directory)
    got = c.suggest_recipients("a", top=5)
    assert [g["address"] for g in got] == ["ann@x.ch", "ada@x.ch"]


def test_suggestions_with_no_text_are_your_top_contacts(monkeypatch):
    c = people_client(monkeypatch, [person("Ann", "ann@x.ch")])
    assert c.suggest_recipients()[0]["name"] == "Ann"
    assert "$search" not in c.calls["people"][0] and c.calls["users"] == []


def test_suggestions_skip_people_without_an_address_and_clean_the_query(monkeypatch):
    nobody = {"id": "x", "displayName": "No mail", "scoredEmailAddresses": []}
    c = people_client(monkeypatch, [nobody])
    assert c.suggest_recipients(' a"b ') == []
    assert c.calls["people"][0]["$search"] == '"ab"'


# --- search_chat ---------------------------------------------------------------
def search_client(monkeypatch, *pages):
    c = outnote.OutNote("tok")
    bodies = []
    answers = list(pages)

    def post(path, body, client=None):
        bodies.append(body["requests"][0])
        hits, more = answers.pop(0)
        return {
            "value": [
                {
                    "hitsContainers": [
                        {
                            "hits": [
                                {"hitId": h[0], "summary": "s", "resource": {"chatId": h[1]}}
                                for h in hits
                            ],
                            "moreResultsAvailable": more,
                        }
                    ]
                }
            ]
        }

    monkeypatch.setattr(c, "post", post)
    c.bodies = bodies
    return c


def test_search_chat_keeps_that_chat_across_pages(monkeypatch):
    c = search_client(
        monkeypatch,
        ([("1", "other"), ("2", "mine")], True),
        ([("3", "mine"), ("4", "other")], False),
    )
    got = c.search_chat("mine", "word")
    assert [g["id"] for g in got] == ["2", "3"] and got[0]["_summary"] == "s"
    assert "from" not in c.bodies[0] and c.bodies[1]["from"] == 2
    assert c.bodies[0]["entityTypes"] == ["chatMessage"]


def test_search_chat_stops_at_top_and_at_scan(monkeypatch):
    many = [(str(i), "mine") for i in range(25)]
    c = search_client(monkeypatch, (many, True))
    assert len(c.search_chat("mine", "w", top=5)) == 5 and len(c.bodies) == 1
    others = [(str(i), "other") for i in range(25)]
    c = search_client(monkeypatch, (others, True), (others, True), (others, True))
    assert c.search_chat("mine", "w", scan=50) == [] and len(c.bodies) == 2


def test_search_is_unchanged(monkeypatch):
    c = search_client(monkeypatch, ([("1", "x")], True))
    assert [r["id"] for r in c.search("w", "chatMessage", 3)] == ["1"]
    assert c.bodies[0]["size"] == 3
