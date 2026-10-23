"""Calendar: events, calendars, scheduling (free/busy) and event writes."""

from datetime import datetime, timedelta, timezone

from ._base import Transport


def _dt(value, tz="UTC"):
    if isinstance(value, datetime):
        if value.tzinfo:
            value = value.astimezone(timezone.utc).replace(tzinfo=None)
            tz = "UTC"
        value = value.isoformat(timespec="seconds")
    return {"dateTime": value, "timeZone": tz}


def _iso(value):
    if isinstance(value, datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    return value


def _attendees(users, kind="required"):
    if isinstance(users, str):
        users = [users]
    return [{"emailAddress": {"address": u}, "type": kind} for u in users or []]


class CalendarMixin(Transport):
    def events(self, start=None, end=None, top=25, calendar_id=None):
        """Events; with `start`/`end` (ISO strings or datetimes) the expanded
        calendarView for that window, recurring instances included."""
        base = f"me/calendars/{calendar_id}" if calendar_id else "me"
        if start and end:
            params = {
                "startDateTime": _iso(start),
                "endDateTime": _iso(end),
                "$top": top,
                "$orderby": "start/dateTime",
            }
            return self.list(f"{base}/calendarView", params, limit=top)
        return self.list(f"{base}/events", {"$top": top, "$orderby": "start/dateTime"}, limit=top)

    def event(self, event_id):
        return self.get(f"me/events/{event_id}")

    def today(self, top=50):
        """Today's events (local day of this machine)."""
        start = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
        return self.events(start, start + timedelta(days=1), top)

    def calendars(self):
        return self.list("me/calendars")

    # --- scheduling ---------------------------------------------------------
    def free_busy(self, users, start, end, interval=30):
        """Free/busy for `users` (addresses) over a window. Each result has
        `scheduleId`, `availabilityView` (one digit per `interval` minutes:
        0 free, 1 tentative, 2 busy, 3 oof, 4 elsewhere) and `scheduleItems`."""
        if isinstance(users, str):
            users = [users]
        body = {
            "schedules": list(users),
            "startTime": _dt(_iso(start).rstrip("Z")),
            "endTime": _dt(_iso(end).rstrip("Z")),
            "availabilityViewInterval": interval,
        }
        return (self.post("me/calendar/getSchedule", body) or {}).get("value", [])

    def find_meeting_times(self, attendees, duration_minutes=30, start=None, end=None, max_=5):
        """Suggested slots when all `attendees` (addresses) are free."""
        body = {
            "attendees": _attendees(attendees),
            "meetingDuration": f"PT{int(duration_minutes)}M",
            "maxCandidates": max_,
            "returnSuggestionReasons": True,
        }
        if start and end:
            body["timeConstraint"] = {
                "activityDomain": "work",
                "timeSlots": [
                    {"start": _dt(_iso(start).rstrip("Z")), "end": _dt(_iso(end).rstrip("Z"))}
                ],
            }
        r = self.post("me/findMeetingTimes", body) or {}
        return r.get("meetingTimeSuggestions", [])

    # --- writes (mutate the calendar; attendees get invitations) ------------
    def create_event(
        self,
        subject,
        start,
        end,
        attendees=None,
        body=None,
        location=None,
        online=False,
        tz="UTC",
        calendar_id=None,
    ):
        """Create an event. `start`/`end` are datetimes (aware ones are
        converted to UTC) or ISO strings interpreted in `tz`. Any `attendees`
        are sent invitations. `online=True` adds a Teams meeting."""
        ev = {"subject": subject, "start": _dt(start, tz), "end": _dt(end, tz)}
        if attendees:
            ev["attendees"] = _attendees(attendees)
        if body:
            ev["body"] = {"contentType": "Text", "content": body}
        if location:
            ev["location"] = {"displayName": location}
        if online:
            ev["isOnlineMeeting"] = True
            ev["onlineMeetingProvider"] = "teamsForBusiness"
        path = f"me/calendars/{calendar_id}/events" if calendar_id else "me/events"
        return self.post(path, ev)

    def update_event(self, event_id, **fields):
        """PATCH raw Graph event fields (`subject=`, `start=_dt(...)`, ...).
        `start`/`end` may be given as datetimes or ISO strings (UTC)."""
        for k in ("start", "end"):
            if k in fields and not isinstance(fields[k], dict):
                fields[k] = _dt(fields[k])
        return self.patch(f"me/events/{event_id}", fields)

    def delete_event(self, event_id):
        return self.delete(f"me/events/{event_id}")

    def respond_event(self, event_id, response="accept", comment=None, send=True):
        """`response` is 'accept', 'tentativelyAccept' or 'decline'."""
        body = {"sendResponse": send}
        if comment:
            body["comment"] = comment
        return self.post(f"me/events/{event_id}/{response}", body)
