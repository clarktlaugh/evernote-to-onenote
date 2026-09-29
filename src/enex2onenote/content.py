"""Convert Evernote ENML into OneNote-friendly XHTML and plain text."""

from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass
from typing import Callable

from bs4 import BeautifulSoup, Comment, Tag

from .enex import Note, Resource

log = logging.getLogger(__name__)

# Elements OneNote's input HTML either rejects or strips; drop them outright.
_DROP_TAGS = {"script", "style", "form", "input", "button", "select", "textarea",
              "noscript", "object", "embed", "applet", "meta", "link", "head"}
_KEEP_ATTRS = {"href", "src", "alt", "width", "height", "style", "colspan",
               "rowspan", "start", "type"}

# Image types OneNote renders inline; anything else is attached as a file.
ONENOTE_IMAGE_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/gif",
                       "image/bmp", "image/tiff"}


def _soup(enml: str) -> BeautifulSoup:
    enml = re.sub(r"<\?xml[^>]*\?>", "", enml)
    enml = re.sub(r"<!DOCTYPE[^>]*>", "", enml, flags=re.IGNORECASE)
    return BeautifulSoup(enml, "html.parser")


def note_text(note: Note) -> str:
    """Plain text of the note body (attachments not included)."""
    soup = _soup(note.enml)
    for t in soup.find_all(["en-crypt", "script", "style"]):
        t.decompose()
    for br in soup.find_all("br"):
        br.replace_with("\n")
    text = soup.get_text("\n")
    text = re.sub(r"[ \t ]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def pdf_text(data: bytes, max_pages: int = 30) -> tuple[str, int]:
    """Extract text from a PDF. Returns (text, page_count)."""
    try:
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                return "", 0
        pages = reader.pages
        parts = []
        for page in list(pages)[:max_pages]:
            try:
                parts.append(page.extract_text() or "")
            except Exception:  # malformed page streams are common in scans
                continue
        text = re.sub(r"[ \t]+", " ", "\n".join(parts))
        return re.sub(r"\n\s*\n+", "\n\n", text).strip(), len(pages)
    except Exception as e:
        log.debug("PDF text extraction failed: %s", e)
        return "", 0


def pdf_first_pages(data: bytes, pages: int) -> bytes | None:
    """Return a PDF containing only the first `pages` pages (for scanned PDFs)."""
    try:
        from pypdf import PdfReader, PdfWriter

        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            reader.decrypt("")
        if len(reader.pages) <= pages:
            return data
        writer = PdfWriter()
        for p in list(reader.pages)[:pages]:
            writer.add_page(p)
        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception as e:
        log.debug("PDF trimming failed: %s", e)
        return None


@dataclass
class ExtractedContent:
    text: str
    truncated: bool
    scanned_pdfs: list[Resource]  # PDFs with no extractable text
    images: list[Resource]


def extract_for_analysis(note: Note, max_chars: int) -> ExtractedContent:
    """Gather the text Claude will read: note body plus attachment text."""
    sections = []
    body = note_text(note)
    if body:
        sections.append(body)
    scanned: list[Resource] = []
    images: list[Resource] = []
    for r in note.resources:
        if r.is_pdf:
            text, npages = pdf_text(r.data)
            # Scanned PDFs have no text layer (or only a few stray characters).
            if len(text) < 20 * max(npages, 1):
                scanned.append(r)
            if text:
                sections.append(f"--- Attached PDF: {r.filename} ({npages} pages) ---\n{text}")
        elif r.is_image:
            images.append(r)
        elif r.mime.startswith("text/"):
            sections.append(f"--- Attached file: {r.filename} ---\n"
                            + r.data.decode("utf-8", errors="replace"))
    full = "\n\n".join(sections)
    truncated = len(full) > max_chars
    return ExtractedContent(full[:max_chars] if truncated else full, truncated, scanned, images)


def _clean(tag: Tag) -> None:
    for el in list(tag.find_all(True)):
        if el.decomposed:
            continue
        if el.name in _DROP_TAGS:
            el.decompose()
            continue
        el.attrs = {k: v for k, v in el.attrs.items() if k in _KEEP_ATTRS}
    for c in tag.find_all(string=lambda s: isinstance(s, Comment)):
        c.extract()


def enml_to_html(note: Note, render_resource: Callable[[Resource, dict], str]) -> tuple[str, list[Resource]]:
    """Convert a note body to OneNote input HTML.

    `render_resource(resource, attrs)` returns the HTML fragment to place where
    an <en-media> element appeared. Returns (html, unreferenced_resources) so the
    caller can append attachments the body never referenced.
    """
    soup = _soup(note.enml)
    root = soup.find("en-note") or soup
    referenced: set[str] = set()
    fragments: list[str] = []

    # Swap each <en-media> for a text marker first: _clean() strips attributes
    # and tags (like <object>) that our generated fragments need to keep.
    for media in root.find_all("en-media"):
        res = note.resource_by_hash(media.get("hash", ""))
        if res is None:
            media.decompose()
            continue
        referenced.add(res.hash)
        fragments.append(render_resource(res, dict(media.attrs)))
        media.replace_with(f"@@ENMEDIA{len(fragments) - 1}@@")

    for todo in root.find_all("en-todo"):
        todo.replace_with("\u2611 " if todo.get("checked") == "true" else "\u2610 ")
    for crypt in root.find_all("en-crypt"):
        p = soup.new_tag("p")
        p.string = "[Encrypted Evernote content was not migrated]"
        crypt.replace_with(p)

    _clean(root)
    if isinstance(root, Tag) and root.name == "en-note":
        root.name = "div"
        root.attrs = {}
        html = str(root)
    else:
        html = f"<div>{root}</div>"
    html = re.sub(r"@@ENMEDIA(\d+)@@", lambda m: fragments[int(m.group(1))], html)
    unreferenced = [r for r in note.resources if r.hash not in referenced]
    return html, unreferenced
