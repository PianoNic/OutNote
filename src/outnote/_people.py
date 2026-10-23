"""People & profile: directory lookup, contacts, photo, presence, groups."""

import re

from ._base import C_SITES, GraphError, Transport

_GUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.I)
_EVENTUAL = {"ConsistencyLevel": "eventual"}


class DirectoryMixin(Transport):
    def find_users(self, query, top=10):
        """Directory users whose name or mail matches `query` (prefix match on
        each word, like the Teams people picker)."""
        q = query.replace('"', "")
        params = {
            "$search": f'"displayName:{q}" OR "mail:{q}" OR "userPrincipalName:{q}"',
            "$count": "true",
            "$top": top,
            "$select": "id,displayName,mail,userPrincipalName,jobTitle,department",
        }
        return self.list("users", params, limit=top, headers=_EVENTUAL)

    def user(self, id_or_upn):
        return self.get(f"users/{id_or_upn}")

    def resolve_user(self, user):
        """Exactly one directory user for `user`: a dict with an `id` passes
        through; an id/UPN/email is looked up directly; anything else is a
        name search that must match exactly one person."""
        if isinstance(user, dict) and user.get("id"):
            return user
        cache = self._ids.setdefault("users", {})
        if user in cache:
            return cache[user]
        found = None
        if _GUID.match(user) or "@" in user:
            try:
                found = self.get(
                    f"users/{user}", {"$select": "id,displayName,mail,userPrincipalName"}
                )
            except GraphError as e:
                if e.status != 404:
                    raise
                hits = self.list(
                    "users",
                    {"$filter": f"mail eq '{user}'", "$select": "id,displayName,mail"},
                )
                found = hits[0] if len(hits) == 1 else None
        else:
            hits = self.find_users(user, top=5)
            exact = [h for h in hits if (h.get("displayName") or "").lower() == user.lower()]
            if len(exact) == 1 or len(hits) == 1:
                found = (exact or hits)[0]
            elif hits:
                names = ", ".join(f"{h['displayName']} <{h.get('mail')}>" for h in hits)
                raise LookupError(f"{user!r} is ambiguous: {names}")
        if not found:
            raise LookupError(f"no directory user matches {user!r}")
        cache[user] = found
        return found


class ProfileMixin(Transport):
    def photo(self):
        """Your profile photo as bytes, or None if there is none."""
        try:
            return self.get("me/photo/$value", raw="bytes")
        except GraphError as e:
            if e.status == 404:
                return None
            raise

    def presence(self):
        """Your Teams status (`availability`, `activity`). Presence.Read.All
        rides the OfficeHome client, not Office."""
        return self.get("me/presence", client=C_SITES)

    def presences(self, user_ids):
        r = (
            self.post(
                "communications/getPresencesByUserId", {"ids": list(user_ids)}, client=C_SITES
            )
            or {}
        )
        return r.get("value", [])

    def member_of(self):
        """Groups, teams and directory roles you belong to."""
        return self.list("me/memberOf", {"$top": 100})

    def contacts(self, top=50):
        return self.list("me/contacts", {"$top": top}, limit=top)
