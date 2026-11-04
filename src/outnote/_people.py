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
        return self._directory(self.list, "users", params, limit=top, headers=_EVENTUAL)

    def user(self, id_or_upn):
        return self._directory(self.get, f"users/{id_or_upn}")

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
                found = self._directory(
                    self.get, f"users/{user}", {"$select": "id,displayName,mail,userPrincipalName"}
                )
            except GraphError as e:
                if e.status != 404:
                    raise
                hits = self._directory(
                    self.list,
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
    def suggest_recipients(self, text="", top=5):
        """Email address suggestions as someone types, people you contact most first. Returns
        `{"name", "address", "id"}` dicts. With no `text`, your top contacts. Fills up from the
        directory if your own contacts give too few, which adds about 100 ms."""
        params = {"$top": top, "$select": "id,displayName,scoredEmailAddresses"}
        text = text.strip()
        if text:
            params["$search"] = '"' + text.replace('"', "") + '"'
        out, seen = [], set()

        def add(name, address, person_id):
            key = (address or "").lower()
            if key and key not in seen and len(out) < top:
                seen.add(key)
                out.append({"name": name, "address": address, "id": person_id})

        for p in self.get("me/people", params).get("value", []):
            scored = p.get("scoredEmailAddresses") or []
            add(p.get("displayName"), scored[0].get("address") if scored else None, p.get("id"))
        if text and len(out) < top:
            for u in self.find_users(text, top):
                add(u.get("displayName"), u.get("mail") or u.get("userPrincipalName"), u.get("id"))
        return out

    def people(self, top=25):
        """People you work with most, ranked by Microsoft (`me/people`)."""
        params = {"$top": top, "$select": "id,displayName,jobTitle,department,scoredEmailAddresses"}
        return self.list("me/people", params, limit=top)

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
        return self._directory(self.list, "me/memberOf", {"$top": 100})

    def contacts(self, top=50):
        return self.list("me/contacts", {"$top": top}, limit=top)
