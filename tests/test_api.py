"""outnote.api: the cache, the client pool and the response helpers. No network."""

import json
import threading
import time

import pytest

import outnote
from outnote import api
from outnote.api import ApiCache, ClientPool, error_response, respond


def test_fresh_answers_from_memory():
    calls = []
    c = ApiCache(fresh=60)
    fetch = lambda: calls.append(1) or {"n": len(calls)}  # noqa: E731
    c.get("u", "a", fetch)
    assert c.get("u", "a", fetch).data == {"n": 1} and len(calls) == 1


def test_stale_answers_at_once_and_refreshes_in_the_background():
    c = ApiCache(fresh=0, stale=60)
    c.get("u", "a", lambda: {"v": 1})
    started = time.perf_counter()
    done = threading.Event()

    def slow():
        time.sleep(0.2)
        done.set()
        return {"v": 2}

    assert c.get("u", "a", slow).data == {"v": 1}
    assert time.perf_counter() - started < 0.1
    assert done.wait(2)
    time.sleep(0.05)
    assert c.get("u", "a", lambda: {"v": 3}, fresh=60).data == {"v": 2}


def test_expired_waits_for_new_data():
    c = ApiCache(fresh=0, stale=0)
    c.get("u", "a", lambda: {"v": 1})
    assert c.get("u", "a", lambda: {"v": 2}).data == {"v": 2}


def test_users_never_share_results():
    c = ApiCache()
    c.get("alice", "mail", lambda: {"who": "alice"})
    assert c.get("bob", "mail", lambda: {"who": "bob"}).data == {"who": "bob"}


def test_identical_requests_share_one_fetch():
    c = ApiCache()
    calls = []

    def slow():
        calls.append(1)
        time.sleep(0.2)
        return {}

    threads = [threading.Thread(target=lambda: c.get("u", "x", slow)) for _ in range(10)]
    [t.start() for t in threads]
    [t.join() for t in threads]
    assert len(calls) == 1


def boom():
    raise RuntimeError("down")


def test_failures_serve_old_data_raise_otherwise_and_are_not_remembered():
    c = ApiCache(fresh=0, stale=0)
    c.get("u", "x", lambda: {"v": 1})
    assert c.get("u", "x", boom).data == {"v": 1}
    empty = ApiCache()
    with pytest.raises(RuntimeError):
        empty.get("u", "y", boom)
    assert empty.get("u", "y", lambda: {"v": 2}).data == {"v": 2}


def test_invalidate_beats_a_slow_old_read():
    c = ApiCache()
    release = threading.Event()

    def old():
        release.wait(2)
        return {"v": "old"}

    t = threading.Thread(target=lambda: c.get("u", "m", old))
    t.start()
    time.sleep(0.05)
    c.invalidate("u", "m")
    assert c.get("u", "m", lambda: {"v": "new"}).data == {"v": "new"}
    release.set()
    t.join()
    time.sleep(0.05)
    assert c.get("u", "m", lambda: {"v": "x"}).data == {"v": "new"}


def test_invalidate_uses_prefix_and_only_that_user():
    c = ApiCache()
    c.get("a", "mail", lambda: 1)
    c.get("a", "mail/1", lambda: 1)
    c.get("a", "chats", lambda: 1)
    c.get("b", "mail", lambda: 2)
    c.invalidate("a", "mail")
    assert set(c._entries) == {("a", "chats"), ("b", "mail")}
    c.invalidate("a")
    assert set(c._entries) == {("b", "mail")}


def test_size_is_bounded_and_nothing_is_left_running():
    c = ApiCache(max_entries=3)
    for i in range(6):
        c.get("u", str(i), lambda i=i: i)
    time.sleep(0.05)
    assert len(c._entries) == 3 and not c._running


# --- responses --------------------------------------------------------------------
def entry():
    return ApiCache().get("u", "x", lambda: {"a": 1})


def test_respond_sends_json_privately():
    e = entry()
    status, headers, body = respond(e)
    assert status == 200 and json.loads(body) == {"a": 1}
    assert headers["ETag"] == e.etag and "private" in headers["Cache-Control"]
    assert headers["Content-Type"] == "application/json"


@pytest.mark.parametrize("template", ["{}", "W/{}", '"other", {}', "*"])
def test_respond_304_when_the_copy_matches(template):
    e = entry()
    status, headers, body = respond(e, template.format(e.etag))
    assert status == 304 and body == b"" and headers["ETag"] == e.etag


def test_respond_200_when_the_copy_differs():
    assert respond(entry(), '"old"')[0] == 200


def test_etag_changes_with_the_data():
    c = ApiCache()
    assert c.get("u", "a", lambda: 1).etag != c.get("u", "b", lambda: 2).etag


def test_error_responses():
    status, headers, body = error_response(outnote.GraphError(429, "busy", retry_after=7.5))
    assert status == 429 and headers["Retry-After"] == "7" and b"busy" in body
    assert error_response(outnote.GraphError(404, "gone"))[0] == 404
    assert error_response(outnote.GraphError(403, "no"))[0] == 403
    status, headers, _ = error_response(outnote.GraphError(500, "oops"))
    assert status == 502 and "Retry-After" not in headers


# --- client pool ------------------------------------------------------------------
class FakeBroker:
    """Rotates its refresh token on every token request, like the real one can."""

    created: list = []

    def __init__(self, refresh_token, tenant="common", store=None):
        self.refresh_token, self.tenant, self.calls = refresh_token, tenant, []
        FakeBroker.created.append(self)

    def __call__(self, client_id, *scope):
        self.calls.append((client_id, *scope))
        self.refresh_token += "+"
        return f"access-for-{client_id}"


@pytest.fixture
def broker(monkeypatch):
    FakeBroker.created = []
    monkeypatch.setattr(api, "Broker", FakeBroker)


def test_pool_reuses_one_client_per_user(broker):
    loaded = []
    pool = ClientPool(lambda u: loaded.append(u) or f"rt-{u}")
    a = pool.get("a")
    assert pool.get("a") is a and pool.get("b") is not a
    assert loaded == ["a", "b"]
    assert FakeBroker.created[0].refresh_token == "rt-a"


def test_pool_saves_rotated_tokens(broker):
    saved = []
    pool = ClientPool(lambda u: "rt", lambda u, t: saved.append((u, t)))
    client = pool.get("a")
    assert client._tp("office") == "access-for-office"
    client._tp("sites", "https://x/.default")
    assert saved == [("a", "rt+"), ("a", "rt++")]
    assert FakeBroker.created[0].calls[1] == ("sites", "https://x/.default")


def test_a_failed_save_does_not_break_the_request(broker):
    def bad(user, token):
        raise OSError("db down")

    assert ClientPool(lambda u: "rt", bad).get("a")._tp("x") == "access-for-x"


def test_pool_is_brokered_so_native_stays_available(broker):
    assert ClientPool(lambda u: "rt").get("a").native is True
    assert ClientPool(lambda u: "rt", native=False).get("b").native is False


def test_missing_token_is_clear(broker):
    with pytest.raises(LookupError, match="no refresh token"):
        ClientPool(lambda u: None).get("a")


def test_drop_and_limit(broker):
    pool = ClientPool(lambda u: "rt", max_clients=2)
    first = pool.get("a")
    pool.drop("a")
    assert pool.get("a") is not first
    pool.get("b")
    pool.get("c")
    assert set(pool._clients) == {"b", "c"}


def test_token_requests_for_one_user_run_one_at_a_time(broker):
    client = ClientPool(lambda u: "rt").get("a")
    inside, peak = [0], [0]
    original = FakeBroker.__call__

    def slow(self, client_id, *scope):
        inside[0] += 1
        peak[0] = max(peak[0], inside[0])
        time.sleep(0.01)
        inside[0] -= 1
        return original(self, client_id, *scope)

    FakeBroker.__call__ = slow
    try:
        threads = [threading.Thread(target=client._tp, args=("x",)) for _ in range(10)]
        [t.start() for t in threads]
        [t.join() for t in threads]
    finally:
        FakeBroker.__call__ = original
    assert peak[0] == 1


def test_importing_outnote_alone_does_not_load_the_helpers():
    import subprocess
    import sys

    code = "import sys, outnote; print('outnote.api' in sys.modules)"
    assert subprocess.check_output([sys.executable, "-c", code], text=True).strip() == "False"
