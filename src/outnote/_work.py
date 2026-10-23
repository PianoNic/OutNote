"""To Do, Planner, Education (classes & assignments) and unified search."""

from datetime import datetime, timedelta, timezone

from ._base import C_OFFICE, C_SITES, Transport

# which client's token each unified-search entity type needs
_SEARCH_CLIENT = {
    "message": C_OFFICE,
    "event": C_OFFICE,
    "driveItem": C_OFFICE,
    "chatMessage": C_OFFICE,
    "person": C_OFFICE,
    "site": C_SITES,
    "listItem": C_SITES,
}
C_EDU = C_OFFICE  # can mint a token for the Teams Assignments API
EDU_API = "https://assignments.edu.cloud.microsoft/api/v1.0"
EDU_SCOPE = "api://8f348934-64be-4bb2-bc16-c54c96789f43/.default"


class TasksMixin(Transport):
    # --- to do --------------------------------------------------------------
    def task_lists(self):
        return self.list("me/todo/lists")

    def tasks(self, list_id, top=50, include_completed=True):
        params = {"$top": top}
        if not include_completed:
            params["$filter"] = "status ne 'completed'"
        return self.list(f"me/todo/lists/{list_id}/tasks", params, limit=top)

    def create_task(self, list_id, title, due=None, body=None, importance=None):
        """Add a To Do task; `due` is a date/datetime or ISO string."""
        task = {"title": title}
        if due:
            if not isinstance(due, str):
                due = due.isoformat()
            task["dueDateTime"] = {"dateTime": due, "timeZone": "UTC"}
        if body:
            task["body"] = {"contentType": "text", "content": body}
        if importance:
            task["importance"] = importance  # 'low' / 'normal' / 'high'
        return self.post(f"me/todo/lists/{list_id}/tasks", task)

    def complete_task(self, list_id, task_id, done=True):
        status = "completed" if done else "notStarted"
        return self.patch(f"me/todo/lists/{list_id}/tasks/{task_id}", {"status": status})

    def delete_task(self, list_id, task_id):
        return self.delete(f"me/todo/lists/{list_id}/tasks/{task_id}")

    # --- planner ------------------------------------------------------------
    def planner_tasks(self):
        """Planner tasks assigned to you, across every team's plans."""
        return self.list("me/planner/tasks")

    def plans(self, group_id=None):
        """Your plans, or a team's (`group_id` = the team id)."""
        path = f"groups/{group_id}/planner/plans" if group_id else "me/planner/plans"
        return self.list(path)

    def plan_tasks(self, plan_id):
        return self.list(f"planner/plans/{plan_id}/tasks")

    def buckets(self, plan_id):
        return self.list(f"planner/plans/{plan_id}/buckets")

    def planner_task_details(self, task_id):
        """Description, checklist and references of a Planner task."""
        return self.get(f"planner/tasks/{task_id}/details")


class EducationMixin(Transport):
    """Teams Assignments. Graph's /education endpoints need EduAssignments.*,
    which no FOCI client carries; the Teams Assignments app instead calls its
    own API (same shapes, `edu/...` paths), and the Office client can mint a
    token for that one."""

    def _edu(self, path, params=None, limit=None):
        return self.list(f"{EDU_API}/{path}", params, limit=limit, client=C_EDU, scope=EDU_SCOPE)

    def classes(self):
        return self._edu("edu/me/classes")

    def assignments(self, class_id=None, top=50):
        """Your assignments across classes (with your submission and its
        outcomes expanded), or one class's."""
        if class_id:
            return self._edu(f"edu/classes/{class_id}/assignments", {"$top": top}, top)
        params = {"$top": top, "$expand": "submissions($expand=outcomes)"}
        return self._edu("edu/me/work", params, top)

    def assignment_resources(self, class_id, assignment_id):
        """Instructions material / files attached to an assignment."""
        return self._edu(f"edu/classes/{class_id}/assignments/{assignment_id}/resources")

    def submissions(self, class_id, assignment_id):
        """Your submission(s) for an assignment (status, grade via outcomes)."""
        return self._edu(
            f"edu/classes/{class_id}/assignments/{assignment_id}/submissions",
            {"$expand": "outcomes"},
        )

    def due_assignments(self, days=7):
        """Assignments due within the next `days` days, soonest first."""
        now = datetime.now(timezone.utc)
        horizon = (now + timedelta(days=days)).isoformat().replace("+00:00", "Z")
        start = now.isoformat().replace("+00:00", "Z")
        due = [
            a
            for a in self.assignments(top=500)
            if a.get("dueDateTime") and start <= a["dueDateTime"] <= horizon
        ]
        return sorted(due, key=lambda a: a["dueDateTime"])


class SearchMixin(Transport):
    def search(self, query, entity="message", top=25):
        """Search one entity type across the account and return the matched
        resources, each with a usable `id` (the hit's id when the resource
        has none) and the hit's `_summary` snippet. `entity` is one of:
        'message' (email), 'driveItem' (files), 'chatMessage' (Teams
        chat/channel messages), 'event' (calendar), 'person', 'site',
        'listItem'. (Graph only allows one of these per request.)"""
        client = _SEARCH_CLIENT.get(entity, C_OFFICE)
        body = {
            "requests": [{"entityTypes": [entity], "query": {"queryString": query}, "size": top}]
        }
        data = self.post("search/query", body, client=client) or {}
        containers = (data.get("value") or [{}])[0].get("hitsContainers") or []
        hits = containers[0].get("hits", []) if containers else []
        out = []
        for h in hits:
            r = dict(h.get("resource") or {})
            r["id"] = r.get("id") or h.get("hitId")
            if h.get("summary"):
                r["_summary"] = h["summary"]
            out.append(r)
        return out

    def search_emails(self, query, top=25):
        return self.search(query, "message", top)

    def search_files(self, query, top=25):
        return self.search(query, "driveItem", top)

    def search_messages(self, query, top=25):
        """Search Teams chat and channel messages."""
        return self.search(query, "chatMessage", top)

    def search_people(self, query, top=25):
        return self.search(query, "person", top)
