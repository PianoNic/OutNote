"""OneDrive, SharePoint and OneNote: browse, download, upload, page text and
images."""

import base64
import os
import re

from ._base import C_ONENOTE, C_SITES, Transport, fetch
from ._text import html_to_text

_SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024  # Graph's limit for a single PUT
_CHUNK = 10 * 320 * 1024  # upload-session chunks must be multiples of 320 KiB
_CONVERT_TIMEOUT = 120  # seconds; Graph converts before it answers


def _item_base(item_id, drive_id=None):
    return f"drives/{drive_id}/items/{item_id}" if drive_id else f"me/drive/items/{item_id}"


def _onenote(group_id=None):
    return f"groups/{group_id}/onenote" if group_id else "me/onenote"


def _write(dest, data):
    if dest:
        with open(dest, "wb") as f:
            f.write(data)
    return data


class FilesMixin(Transport):
    # --- onedrive -----------------------------------------------------------
    def files(self, path="root", top=100):
        """OneDrive items. `path` is 'root' or 'root:/Folder/Sub:'."""
        return self.list(f"me/drive/{path}/children", {"$top": top}, limit=top)

    def file(self, item_id, drive_id=None):
        return self.get(_item_base(item_id, drive_id))

    def download_url(self, item_id, drive_id=None):
        """A direct, pre-authed download URL for a drive item (works for OneDrive
        and channel files — pass the item's drive_id for the latter)."""
        return self.get(_item_base(item_id, drive_id)).get("@microsoft.graph.downloadUrl")

    def download(self, item_id, dest=None, drive_id=None):
        """A drive item's bytes (also written to `dest` if given)."""
        url = self.download_url(item_id, drive_id)
        if not url:
            raise ValueError(f"item {item_id} has no content (a folder?)")
        return _write(dest, fetch(url))

    def converted_url(self, item_id, format="pdf", drive_id=None):
        """A short-lived pre-authenticated URL of a drive item converted to
        `format` (Graph's `content?format=`; "pdf" takes Word, Excel,
        PowerPoint, OpenDocument, RTF and CSV files). Raises `GraphError`
        (406/415) for a file Graph can't convert."""
        return self.redirect(
            f"{_item_base(item_id, drive_id)}/content", {"format": format}, timeout=_CONVERT_TIMEOUT
        )

    def convert(self, item_id, format="pdf", drive_id=None, dest=None):
        """A drive item's bytes converted to `format` (also written to `dest`
        if given). See `converted_url`."""
        url = self.converted_url(item_id, format, drive_id)
        return _write(dest, fetch(url, timeout=_CONVERT_TIMEOUT))

    def download_shared(self, url, dest=None):
        """Bytes behind a SharePoint/OneDrive sharing or content URL — e.g. a
        file attached to a Teams message (`message_files`)."""
        token = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        item = self.get(f"shares/u!{token}/driveItem")
        return _write(dest, fetch(item["@microsoft.graph.downloadUrl"]))

    def upload(self, src, dest_path, drive_id=None, conflict="replace", app_folder=False):
        """Upload a local file path or bytes to `dest_path` ('Folder/name.txt'
        relative to the drive root, or to the app's own folder with
        `app_folder=True`, Graph's `special/approot`). Small files go up in
        one PUT, larger ones through a resumable upload session. Returns the
        driveItem."""
        if isinstance(src, (str, os.PathLike)):
            with open(src, "rb") as f:
                data = f.read()
        else:
            data = bytes(src)
        if app_folder:
            root = f"drives/{drive_id}/special/approot" if drive_id else "me/drive/special/approot"
        else:
            root = f"drives/{drive_id}/root" if drive_id else "me/drive/root"
        target = f"{root}:/{dest_path.lstrip('/')}:"
        if len(data) <= _SIMPLE_UPLOAD_MAX:
            return self.put(f"{target}/content?@microsoft.graph.conflictBehavior={conflict}", data)
        session = self.post(
            f"{target}/createUploadSession",
            {"item": {"@microsoft.graph.conflictBehavior": conflict}},
        )
        return self._upload_chunks(session["uploadUrl"], data)

    def _upload_chunks(self, upload_url, data):
        from urllib.request import Request

        result = None
        for start in range(0, len(data), _CHUNK):
            chunk = data[start : start + _CHUNK]
            end = start + len(chunk) - 1
            req = Request(
                upload_url,
                data=chunk,
                method="PUT",
                headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{len(data)}",
                },
            )  # the session URL is pre-authed; no Bearer header
            result = self._send(req)
        return result

    def delete_file(self, item_id, drive_id=None):
        return self.delete(_item_base(item_id, drive_id))

    # --- sharepoint ---------------------------------------------------------
    def sites(self, search="*"):
        return self.list(f"sites?search={search}", client=C_SITES)

    def followed_sites(self):
        """SharePoint sites you follow (works when site search is down)."""
        return self.list("me/followedSites", client=C_SITES)

    def site_drive_items(self, site_id, path="root", top=100):
        return self.list(
            f"sites/{site_id}/drive/{path}/children", {"$top": top}, limit=top, client=C_SITES
        )

    # --- onenote ------------------------------------------------------------
    def notebooks(self, group_id=None):
        """Your notebooks, or a team/class's with `group_id` (the team id)."""
        return self.list(f"{_onenote(group_id)}/notebooks", client=C_ONENOTE)

    def sections(self, notebook_id=None, group_id=None):
        base = _onenote(group_id)
        path = f"{base}/notebooks/{notebook_id}/sections" if notebook_id else f"{base}/sections"
        return self.list(path, client=C_ONENOTE)

    def pages(self, section_id, top=50, group_id=None):
        """Pages in a section, newest edit first. (List per-section and sort
        here — the all-pages endpoint and $select+$orderby are flaky.)"""
        pages = self.list(
            f"{_onenote(group_id)}/sections/{section_id}/pages",
            {"$top": top},
            limit=top,
            client=C_ONENOTE,
        )
        return sorted(pages, key=lambda p: p.get("lastModifiedDateTime") or "", reverse=True)

    def page_content(self, page_id, group_id=None):
        """A OneNote page's body as HTML (string)."""
        return self.get(f"{_onenote(group_id)}/pages/{page_id}/content", raw=True, client=C_ONENOTE)

    def page_text(self, page_id, group_id=None):
        """A OneNote page as plain text."""
        return html_to_text(self.page_content(page_id, group_id))

    def page_images(self, page_id, group_id=None):
        """Images embedded in a page, as `{"url", "content_type", "data"}`
        (full resolution where OneNote has it). Fetched with the OneNote
        client's token, which those resource URLs require."""
        html = self.page_content(page_id, group_id)
        out = []
        for tag in re.findall(r"<img\b[^>]*>", html, re.I):
            attrs = dict(re.findall(r'([\w-]+)="([^"]*)"', tag))
            url = attrs.get("data-fullres-src") or attrs.get("src")
            if not url or "/resources/" not in url:
                continue
            kind = attrs.get("data-fullres-src-type") or attrs.get("data-src-type")
            data = self.get(url.replace("&amp;", "&"), raw="bytes", client=C_ONENOTE)
            out.append({"url": url, "content_type": kind, "data": data})
        return out

    def create_page(self, section_id, html=None, title=None, body="", group_id=None):
        """Create a page from full XHTML `html`, or from `title` + `body`
        (plain HTML fragment). Returns the page."""
        if html is None:
            from html import escape

            html = (
                f"<!DOCTYPE html><html><head><title>{escape(title or '')}</title></head>"
                f"<body>{body}</body></html>"
            )
        return self.post(
            f"{_onenote(group_id)}/sections/{section_id}/pages",
            html,
            client=C_ONENOTE,
            content_type="application/xhtml+xml",
        )

    def append_page(self, page_id, html, group_id=None):
        """Append an HTML fragment to the end of a page."""
        return self.patch(
            f"{_onenote(group_id)}/pages/{page_id}/content",
            [{"target": "body", "action": "append", "content": html}],
            client=C_ONENOTE,
        )

    def delete_page(self, page_id, group_id=None):
        return self.delete(f"{_onenote(group_id)}/pages/{page_id}", client=C_ONENOTE)
