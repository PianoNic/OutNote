# Method reference

Verified end-to-end against a live locked-down EDU tenant.

| Area | Methods | Endpoint |
|---|---|---|
| Mail | `emails(folder, unread=)`, `unread_emails()`, `email(id)`, `mail_folders()` | `/me/mailFolders/{f}/messages` |
| Mail attachments | `attachments(id)`, `download_attachment(id, att)` | `/me/messages/{id}/attachments` |
| Mail writes | `send_mail`, `reply`, `forward`, `create_draft`, `send_draft`, `mark_read`, `flag`, `move_email`, `delete_email` | `/me/sendMail`, `/me/messages/...` |
| Calendar | `events(start, end)`, `event(id)`, `today()`, `calendars()` | `/me/events`, `/me/calendarView` |
| Scheduling | `free_busy(users, start, end)`, `find_meeting_times(attendees)` | `getSchedule`, `findMeetingTimes` |
| Calendar writes | `create_event`, `update_event`, `delete_event`, `respond_event` | `/me/events` |
| Teams chats | `chats(chat_type=)`, `chat(id)`, `chat_members(id)`, `unread_chats()`, `find_chat(user)`, `chat_messages(id, since=)` | `/me/chats` (members + last message expanded) |
| Teams sending | `send_chat_message(id, text, mentions=)`, `create_chat(users)`, `send_to(user, text)` | `/chats`, `/chats/{id}/messages` |
| Message lifecycle | `edit_chat_message`, `delete_chat_message`, `undelete_chat_message`, `react`, `unreact`, `mark_chat_read`, `mark_chat_unread` | `PATCH`, `softDelete`, `setReaction`, `markChatReadForUser` |
| Chat management | `rename_chat`, `add_chat_member`, `remove_chat_member`, `pinned_messages`, `pin_message`, `unpin_message`, `hide_chat`, `unhide_chat` | `/chats/{id}/...` |
| Teams channels | `teams()`, `channels(t)`, `channel_messages(t, c)`, `channel_message_replies(t, c, m)`, `channel_files(t, c)` | `/me/joinedTeams`, `/teams/{t}/channels/...` |
| Channel posting | `send_channel_message(t, c, text)`, `reply_to_channel_message(t, c, m, text)` | `/teams/{t}/channels/{c}/messages` |
| Shared in messages | `message_images(m, chat_id=)`, `message_files(msg)`, `download_shared(url)` | `hostedContents`, `/shares/u!{url}` |
| Meetings | `meeting_chat_messages(event)`, `meeting_thread_id(event)`, `online_meeting(url)`\*, `transcripts(id)`\* | thread id parsed from the join URL |
| OneDrive | `files(path)`, `file(id)`, `download(id)`, `download_url(id)`, `upload(src, dest)`, `delete_file(id)` | `/me/drive/...` |
| SharePoint | `sites(search)`, `followed_sites()`, `site_drive_items(id)` | `/sites`, `/me/followedSites` |
| OneNote | `notebooks(group_id=)`, `sections`, `pages`, `page_content`, `page_text`, `page_images`, `create_page`, `append_page`, `delete_page` | `/me/onenote/...`, `/groups/{id}/onenote/...` |
| People | `find_users(q)`, `resolve_user(name_or_upn)`, `user(id)`, `contacts()`, `photo()`, `presence()`, `presences(ids)`, `member_of()` | `/users`, `/me/presence` |
| To Do | `task_lists()`, `tasks(list)`, `create_task`, `complete_task`, `delete_task` | `/me/todo/lists` |
| Planner | `planner_tasks()`, `plans(group_id=)`, `plan_tasks(plan)`, `buckets(plan)`, `planner_task_details(task)` | `/me/planner`, `/planner/plans` |
| Assignments | `classes()`, `assignments(class_id=)`, `assignment_resources`, `submissions`, `due_assignments(days)` | Teams Assignments API (see [clients.md](clients.md)) |
| Search | `search_emails`, `search_files`, `search_messages`, `search_people`, `search(q, entity)` | `/search/query` |
| Robustness | `batch(requests)`, `delta(path, delta_link)` | `/$batch` (20 per call, throttled sub-requests retried), `/delta` |
| Text | `message_text(msg)`, `html_to_text(html)`, `extract_text(name, data)` | - |
| Anything | `get`, `post`, `patch`, `put`, `delete`, `list` | any endpoint, follows `@odata.nextLink` |

\* needs `OnlineMeetings.Read` / transcript scopes, which no FOCI client carries
on many tenants (403). The meeting *chat* works without them.

## Caveats

- Works against **managed** Entra tenants (`login.microsoftonline.com`) with these
  first-party clients pre-consented - the common case for Microsoft 365 / EDU.
- **Push / SMS MFA can't be done headless** - only TOTP. Enroll an authenticator
  and pass its Base32 secret, or use `interactive=True`.
- Which scopes each client carries is **tenant-dependent**. Online-meeting lookup
  and transcripts 403 on the tested tenant.
- SharePoint-index search (`search_files`, `sites(search)`, `search(entity="site")`)
  depends on the tenant's search service; when it errors, `followed_sites()`,
  `files()` and `get("me/drive/recent")` still work.
- Treat `.ms_frt` and `.ms_session.json` like passwords (both are gitignored).

See [clients.md](clients.md) for the FOCI client/scope map and how it was found.
