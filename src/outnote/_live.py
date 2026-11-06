"""Live Teams events over Trouter, the push connection Teams itself uses. Undocumented, so it can
change without notice. Nothing here runs unless you call `listen()`, and it keeps no state on disk.

    with note.listen() as live:
        for event in live:
            if event.kind == "typing":
                print(event.sender_id, "is typing in", event.chat_id)
            elif event.kind == "message":
                print(event.sender_name, event.message["body"]["content"])
"""

import json
import logging
import time
import uuid
from dataclasses import dataclass
from datetime import timedelta
from typing import Any
from urllib.parse import quote, unquote, urlencode

from ._base import GraphError
from ._native import PRESENCE_SCOPE, message_from_native

log = logging.getLogger("outnote.live")

_ALLOCATE = "https://go.trouter.teams.microsoft.com/v4/a"
_REGISTRAR = "https://teams.microsoft.com/registrar/prod/V2/registrations"
_APP = "TeamsCDLWebWorker"
_TEMPLATE = "TeamsCDLWebWorker_2.7"
_VERSION = "1415/1.0.0.0"
_HEARTBEAT = 25  # seconds; the server allows about 70
_POLL = 2  # seconds to wait for a frame, so close() takes effect quickly
_MAX_BACKOFF = 60


@dataclass
class LiveEvent:
    """One live event. `kind` is `typing`, `stop_typing`, `message`, `message_updated`,
    `message_deleted`, `presence` (someone you asked to watch: see `sender_id`, `availability` and
    `activity`), `connected` (the feed is up; events may have been missed before it, so catch up with
    `chat_messages` or `stream`), or `other`. `from_me` is True for your own activity. The original
    event is in `raw`."""

    kind: str
    chat_id: str | None = None
    sender_id: str | None = None
    sender_name: str | None = None
    from_me: bool = False
    message: dict | None = None
    raw: Any = None
    availability: str | None = None
    activity: str | None = None


def parse_frame(frame):
    """Split a socket.io frame into `(type, ack_id, payload)`."""
    parts = frame.split(":", 3)
    return parts[0], parts[1] if len(parts) > 1 else "", parts[3] if len(parts) > 3 else ""


def event_from(raw, me=None):
    """A `LiveEvent` from one item of Trouter's `eventMessages`."""
    resource = raw.get("resource") or {}
    link = resource.get("conversationLink") or ""
    chat_id = unquote(link.rsplit("/", 1)[-1]) if link else None
    mri = (resource.get("from") or "").rsplit("/", 1)[-1]
    sender = mri.removeprefix("8:orgid:") if mri else None
    name = resource.get("imdisplayname") or None
    event = LiveEvent("other", chat_id, sender, name, bool(me and sender == me), None, raw)
    kind = resource.get("messagetype")
    if raw.get("resourceType") not in ("NewMessage", "MessageUpdate") or not kind:
        return event
    if kind == "Control/Typing":
        event.kind = "typing"
    elif kind == "Control/ClearTyping":
        event.kind = "stop_typing"
    elif str(kind).startswith(("Control/", "ThreadActivity/", "Event/")):
        return event
    else:
        event.message = message_from_native(resource, chat_id)
        if event.message["deletedDateTime"]:
            event.kind = "message_deleted"
        else:
            new = raw["resourceType"] == "NewMessage"
            event.kind = "message" if new else "message_updated"
    return event


def presence_events(body, me=None):
    """`LiveEvent`s from a Trouter presence delivery, which starts with everyone's current state."""
    events = []
    for item in body.get("presence") or []:
        state = item.get("presence") or {}
        user = (item.get("mri") or "").removeprefix("8:orgid:") or None
        events.append(
            LiveEvent(
                "presence",
                sender_id=user,
                from_me=bool(me and user == me),
                raw=item,
                availability=state.get("availability"),
                activity=state.get("activity"),
            )
        )
    return events


class _Renewal(Exception):
    """The session token is about to expire: a planned reconnect, not a failure."""


class LiveConnection:
    """An iterator of `LiveEvent`s that reconnects by itself, waiting longer after each failure. Stop
    it with `close()`, from another thread if needed, or by leaving the `with` block.

    The Teams session lasts about 75 minutes. Shortly before it expires, the feed signs in again and
    reconnects straight away, even with `reconnect=False`, and yields `connected` again. Events sent
    during that short gap are missed, so catch up on `connected` as usual."""

    def __init__(self, client, reconnect=True, presence=()):
        self.client, self.reconnect = client, reconnect
        self.presence = list(presence)
        self._closed = False
        self._ws = None
        self._http = None

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def close(self):
        self._closed = True

    def __iter__(self):
        failures = 0
        try:
            while not self._closed:
                try:
                    self._connect()
                    yield LiveEvent("connected")
                    failures = 0
                    yield from self._frames()
                except _Renewal:
                    log.info("live session renewed")
                    self._disconnect()
                    continue
                except Exception as e:  # noqa: BLE001 - network, auth or protocol: all mean reconnect
                    if not self.reconnect:
                        raise
                    log.warning("live connection failed, reconnecting: %s", e)
                self._disconnect()
                if not self.reconnect:
                    return
                failures += 1
                self._sleep(min(2**failures, _MAX_BACKOFF))
        finally:
            self._disconnect()

    def _sleep(self, seconds):
        end = time.monotonic() + seconds
        while not self._closed and time.monotonic() < end:
            time.sleep(min(0.25, seconds))

    def _request(self, method, url, headers, body=None):
        """Without automatic decompression, which fails on Trouter's replies."""
        from entrance._http import Response, _headers
        from wreq import Method

        kw = {"headers": {**headers, "Accept-Encoding": "identity"}}
        if body:
            kw["body"] = body
        elif method == "POST":
            kw["headers"]["Content-Length"] = "0"
        with self._http.request(getattr(Method, method), url, **kw) as r:
            response = Response(r.status.as_int(), _headers(r.headers), bytes(r.bytes()))
        if response.status_code >= 300:
            raise OSError(f"{url.split('?')[0]} answered {response.status_code}")
        return response

    def _connect(self):
        """Allocate a session, open the socket, and register it for chat events."""
        import wreq.blocking as wb

        self._http = wb.Client(gzip=False, brotli=False, deflate=False, zstd=False)
        token = self.client._chatsvc()["headers"]["Authentication"].split("=", 1)[1]
        headers = {"X-Skypetoken": token}
        epid = str(uuid.uuid4())
        session = self._request("POST", f"{_ALLOCATE}?epid={epid}", headers).json()
        server = session["socketio"].rstrip("/") + "/"
        tc = json.dumps(
            {"cv": "2024.23.01.2", "ua": "TeamsCDL", "hr": "", "v": _VERSION}, separators=(",", ":")
        )
        params = {
            "v": "v4",
            "tc": tc,
            "con_num": f"{int(time.time() * 1000)}_1",
            "epid": epid,
            "ccid": session.get("ccid", ""),
            "auth": "true",
            "timeout": "40",
            **(session.get("connectparams") or {}),
        }
        query = urlencode(params, quote_via=quote)
        shake = self._request("GET", f"{server}socket.io/1/?{query}", headers).text
        sid = shake.split(":")[0]
        url = f"{server.replace('https://', 'wss://')}socket.io/1/websocket/{sid}?{query}"
        self._ws = self._http.websocket(url, headers=headers)
        registration = {
            "clientDescription": {
                "appId": _APP,
                "aesKey": "",
                "languageId": "en-US",
                "platform": "edge",
                "templateKey": _TEMPLATE,
                "platformUIVersion": _VERSION,
            },
            "registrationId": str(uuid.uuid4()),
            "nodeId": "",
            "transports": {"TROUTER": [{"context": "", "path": session["surl"], "ttl": 86400}]},
        }
        self._request(
            "POST",
            _REGISTRAR,
            {**headers, "Content-Type": "application/json"},
            json.dumps(registration).encode(),
        )
        if self.presence:
            self._watch_presence(epid, session["surl"])

    def _watch_presence(self, epid, trouter_uri):
        """Ask Teams to push these people's status here. This only reads: it doesn't publish your
        own status, so listening doesn't make you look online."""
        body = {
            "trouterUri": trouter_uri,
            "shouldPurgePreviousSubscriptions": True,
            "subscriptionsToAdd": [{"mri": f"8:orgid:{u}", "source": "ups"} for u in self.presence],
            "subscriptionsToRemove": [],
        }
        service = self.client._chatsvc()["presence"]
        url = f"{service}/v1/pubsub/subscriptions/{epid}"
        headers = {  # the service requires these
            "x-ms-client-user-agent": "Teams-V2-Web",
            "x-ms-correlation-id": str(uuid.uuid4()),
            "x-ms-client-version": _VERSION,
            "x-ms-endpoint-id": epid,
        }
        try:
            self.client.request("POST", url, body, headers=headers, scope=PRESENCE_SCOPE)
        except GraphError as e:  # messages and typing still work without it
            log.warning("could not watch presence: %s", e)

    def _disconnect(self):
        ws, self._ws = self._ws, None
        if ws is not None:
            try:
                ws.close()
            except Exception:  # noqa: BLE001 - already gone
                log.debug("closing the live socket failed", exc_info=True)

    def _send(self, text):
        import wreq

        self._ws.send(wreq.Message.from_text(text))

    def _recv(self):
        message = self._ws.recv(timeout=timedelta(seconds=_POLL))
        if message is None:
            return None
        text = message.text
        text = text() if callable(text) else text
        if text is None:
            return None
        return text if isinstance(text, str) else bytes(text).decode("utf-8", "replace")

    def _frames(self):
        me = self.client._me_id()
        expires = self.client._chatsvc()["exp"]
        beat = time.monotonic()
        while not self._closed:
            if time.time() > expires - 120:
                self.client._ids.pop("chatsvc", None)  # so the reconnect signs in again
                raise _Renewal()
            if time.monotonic() - beat > _HEARTBEAT:
                self._send("2:::")
                beat = time.monotonic()
            try:
                frame = self._recv()
            except Exception as e:
                if "imeout" in type(e).__name__ or "timed out" in str(e).lower():
                    continue
                raise
            if not frame:
                continue
            kind, ack, payload = parse_frame(frame)
            if kind == "0":
                raise OSError("the server closed the live connection")
            if kind == "2":
                self._send("2:::")
            elif kind == "5" and ack.endswith("+"):
                self._send(f"6:::{ack.removesuffix('+')}+[]")
            elif kind == "3":
                yield from self._delivery(payload, me)

    def _delivery(self, payload, me):
        """Trouter wraps each delivery as a request, which must be answered or it's sent again."""
        try:
            delivery = json.loads(payload)
        except ValueError:
            log.debug("unreadable delivery ignored")
            return
        self._send(
            "3:::"
            + json.dumps({"id": delivery.get("id"), "status": 200, "headers": {}, "body": ""})
        )
        try:
            body = json.loads(delivery.get("body") or "{}")
        except (TypeError, ValueError):
            return
        if body.get("presence"):
            yield from presence_events(body, me)
            return
        items = body.get("eventMessages") or ([body] if body.get("resourceType") else [])
        for raw in items:
            yield event_from(raw, me)
