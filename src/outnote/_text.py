"""Turning Graph payloads into plain text: Teams/OneNote HTML, and the text of
downloaded documents (.docx, .pptx, .xlsx, plain text, optional PDF)."""

import io
import re
import zipfile
from html import escape, unescape
from html.parser import HTMLParser

_BLOCK = {
    "p",
    "div",
    "br",
    "li",
    "tr",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "table",
    "ul",
    "ol",
    "blockquote",
    "pre",
    "hr",
}
_SKIP = {"script", "style", "head", "title", "attachment"}


class _TextParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in _SKIP:
            self._skip += 1
        elif tag in _BLOCK:
            self.parts.append("\n")
            if tag == "li":
                self.parts.append("- ")
        elif tag in ("td", "th"):
            self.parts.append("\t")

    def handle_startendtag(self, tag, attrs):
        if tag in _BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in _SKIP:
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html):
    """Plain text from an HTML fragment: block elements become line breaks,
    list items get a `- ` bullet, entities are unescaped, `<at>` mention names
    are kept and Teams `<attachment>` stubs dropped."""
    if not html:
        return ""
    p = _TextParser()
    p.feed(html)
    p.close()
    text = "".join(p.parts).replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def text_to_html(text):
    """Escape plain text for a Graph HTML body, keeping line breaks."""
    return escape(text).replace("\n", "<br>")


def message_text(msg):
    """A Teams `chatMessage`'s body as plain text (HTML or text bodies)."""
    body = (msg or {}).get("body") or {}
    content = body.get("content") or ""
    if body.get("contentType") == "text":
        return unescape(content).strip()
    return html_to_text(content)


def _xml_text(xml, para_tag):
    # paragraphs -> lines; then drop every tag
    xml = re.sub(rf"</{para_tag}>", "\n", xml)
    xml = re.sub(r"<w:tab/>|<w:br/>", " ", xml)
    return unescape(re.sub(r"<[^>]+>", "", xml))


def _docx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return _xml_text(z.read("word/document.xml").decode("utf-8"), "w:p")


def _pptx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        slides = sorted(
            (n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)),  # type: ignore[union-attr]
        )
        return "\n\n".join(_xml_text(z.read(n).decode("utf-8"), "a:p").strip() for n in slides)


def _xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if "xl/sharedStrings.xml" not in z.namelist():
            return ""
        xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    return "\n".join(
        unescape(re.sub(r"<[^>]+>", "", s)) for s in re.findall(r"<si>(.*?)</si>", xml)
    )


def _pdf(data):
    try:
        from pypdf import PdfReader  # optional: `pip install pypdf`
    except ImportError:
        return ""
    return "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(data)).pages)


_PLAIN = {".txt", ".md", ".csv", ".json", ".xml", ".log", ".py", ".yaml", ".yml", ".vtt"}


def extract_text(name, data):
    """Best-effort text of a document named `name` with content `data`:
    .docx/.pptx/.xlsx (shared strings), .html, plain-text types, and .pdf if
    `pypdf` is installed. Unknown or unreadable types return "" — images are
    left to the caller (OCR/vision) on purpose."""
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""
    try:
        if ext == ".docx":
            return _docx(data).strip()
        if ext == ".pptx":
            return _pptx(data).strip()
        if ext == ".xlsx":
            return _xlsx(data).strip()
        if ext == ".pdf":
            return _pdf(data).strip()
        if ext in (".html", ".htm"):
            return html_to_text(data.decode("utf-8", "replace"))
        if ext in _PLAIN:
            return data.decode("utf-8", "replace")
    except (zipfile.BadZipFile, KeyError, ValueError):
        return ""
    return ""
