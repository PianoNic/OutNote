"""Turn what Graph hands back into plain text: HTML (Teams messages, OneNote
pages, mail bodies), Office files, PDFs and meeting transcripts. Stdlib only;
PDF needs the optional `pypdf` (`pip install outnote[pdf]`)."""

import io
import re
import zipfile
from html import unescape
from html.parser import HTMLParser

_BLOCK = {"p", "div", "tr", "table", "ul", "ol", "blockquote", "pre", "section", "body"}
_HEADINGS = {"title": 1, "h1": 1, "h2": 2, "h3": 3, "h4": 4, "h5": 5, "h6": 6}
_SKIP = {"script", "style", "attachment"}
_TEXT_EXT = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".log", ".yaml", ".yml", ".ini"}


class _HTMLText(HTMLParser):
    def __init__(self, markdown):
        super().__init__(convert_charrefs=True)
        self.md = markdown
        self.out = []
        self.skip = 0
        self.pre = 0
        self.href = []
        self.cells = 0

    def _nl(self, n=1):
        """Make sure the output ends with at least `n` newlines."""
        if not self.out:
            return
        tail = "".join(self.out[-4:])
        have = len(tail) - len(tail.rstrip("\n"))
        if have < n:
            self.out.append("\n" * (n - have))

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in _SKIP:
            self.skip += 1
        elif self.skip:
            return
        elif tag in _HEADINGS:
            self._nl(2)
            if self.md:
                self.out.append("#" * _HEADINGS[tag] + " ")
        elif tag == "br":
            self.out.append("\n")
        elif tag == "li":
            self._nl()
            self.out.append("- ")
        elif tag == "tr":
            self._nl()
            self.cells = 0
        elif tag in ("td", "th"):
            if self.cells:
                self.out.append(" | ")
            self.cells += 1
        elif tag == "pre":
            self._nl(2)
            self.pre += 1
            if self.md:
                self.out.append("```\n")
        elif tag in _BLOCK:
            self._nl(2 if tag == "p" else 1)
        elif tag in ("b", "strong") and self.md:
            self.out.append("**")
        elif tag in ("i", "em") and self.md:
            self.out.append("_")
        elif tag == "a":
            self.href.append(a.get("href"))
            if self.md and a.get("href"):
                self.out.append("[")
        elif tag == "img":
            alt = a.get("alt") or "image"
            src = a.get("data-fullres-src") or a.get("src") or ""
            self.out.append(f"![{alt}]({src})" if self.md else f"[{alt}]")
        elif tag == "emoji":
            self.out.append(a.get("alt") or "")
            self.skip += 1  # its children repeat the emoji

    def handle_endtag(self, tag):
        if tag in _SKIP or tag == "emoji":
            self.skip = max(0, self.skip - 1)
        elif self.skip:
            return
        elif tag in _HEADINGS or tag in _BLOCK or tag == "li":
            if tag == "pre":
                self.pre = max(0, self.pre - 1)
                if self.md:
                    self._nl()
                    self.out.append("```")
            self._nl(2 if tag in _HEADINGS or tag in ("p", "pre") else 1)
        elif tag in ("b", "strong") and self.md:
            self.out.append("**")
        elif tag in ("i", "em") and self.md:
            self.out.append("_")
        elif tag == "a":
            href = self.href.pop() if self.href else None
            if self.md and href:
                self.out.append(f"]({href})")

    def handle_data(self, data):
        if self.skip:
            return
        if not self.pre:
            data = re.sub(r"\s+", " ", data)
            if not self.out or self.out[-1].endswith(("\n", " ")):
                data = data.lstrip()
        self.out.append(data)


def html_to_text(html, markdown=False):
    """HTML -> readable plain text (or light Markdown with `markdown=True`):
    keeps paragraphs, headings, lists, table rows and @mention names; drops
    scripts, styles and Teams `<attachment>` stubs; images become `[alt]`."""
    if not html:
        return ""
    p = _HTMLText(markdown)
    p.feed(html)
    p.close()
    text = "".join(p.out)
    text = "\n".join(line.rstrip() for line in text.splitlines())
    return re.sub(r"\n{3,}", "\n\n", text).strip()


class _Images(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.found = []

    def handle_starttag(self, tag, attrs):
        if tag == "img":
            self.found.append(dict(attrs))


def html_images(html):
    """The attributes of every `<img>` in `html`, in order."""
    p = _Images()
    p.feed(html or "")
    p.close()
    return p.found


def _xml_text(xml, para, extra=()):
    for pat, rep in extra:
        xml = re.sub(pat, rep, xml)
    xml = re.sub(para, "\n", xml)
    return unescape(re.sub(r"<[^>]+>", "", xml))


def _docx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    text = _xml_text(
        xml,
        r"</w:p>",
        [(r"<w:tab/>", "\t"), (r"<w:(?:br|cr)/>", "\n"), (r"</w:tc>", "\t")],
    )
    return "\n".join(line.rstrip() for line in text.splitlines())


def _pptx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        slides = sorted(
            (n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.sub(r"\D", "", n)),
        )
        texts = [_xml_text(z.read(n).decode("utf-8"), r"</a:p>").strip() for n in slides]
    return "\n\n".join(t for t in texts if t)


def _pdf(data):
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""  # optional: pip install outnote[pdf]
    return "\n\n".join((page.extract_text() or "") for page in PdfReader(io.BytesIO(data)).pages)


def extract_text(name, data):
    """Best-effort text of a file from its `name` and `data` bytes: .docx,
    .pptx, .pdf (with the `pdf` extra), .html, and plain-text types. Returns
    "" for anything else or anything it cannot parse - never raises. Images
    are not OCR'd: pass their bytes to a vision model instead."""
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in (name or "") else ""
    try:
        if ext == ".docx":
            return _docx(data).strip()
        if ext == ".pptx":
            return _pptx(data)
        if ext == ".pdf":
            return _pdf(data).strip()
        if ext in (".html", ".htm"):
            return html_to_text(data.decode("utf-8", "replace"))
        if ext in _TEXT_EXT:
            return data.decode("utf-8-sig", "replace")
    except Exception:  # noqa: BLE001 - a corrupt file is "no text", not a crash
        return ""
    return ""


def vtt_to_text(vtt):
    """A Teams WebVTT transcript -> `Speaker: line` text, merging consecutive
    cues from the same speaker."""
    lines = []
    for block in re.split(r"\n\s*\n", vtt.replace("\r\n", "\n")):
        cue = [ln for ln in block.strip().splitlines() if ln and "-->" not in ln]
        if not cue or cue[0].startswith(("WEBVTT", "NOTE")):
            continue
        text = " ".join(cue)
        m = re.match(r"<v ([^>]+)>(.*?)(?:</v>)?$", text)
        speaker, said = (m.group(1), m.group(2)) if m else ("", text)
        said = unescape(re.sub(r"<[^>]+>", "", said)).strip()
        if lines and speaker and lines[-1][0] == speaker:
            lines[-1][1].append(said)
        else:
            lines.append((speaker, [said]))
    return "\n".join(f"{s}: {' '.join(t)}" if s else " ".join(t) for s, t in lines)
