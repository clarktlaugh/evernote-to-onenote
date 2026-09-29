"""Streaming parser for Evernote .enex export files."""

from __future__ import annotations

import base64
import hashlib
import mimetypes
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from defusedxml.ElementTree import iterparse


@dataclass
class Resource:
    """A binary attachment (image, PDF, etc.) embedded in a note."""

    data: bytes
    mime: str
    filename: str
    hash: str  # md5 hex digest; ENML <en-media hash="..."> refers to this
    width: int | None = None
    height: int | None = None

    @property
    def size(self) -> int:
        return len(self.data)

    @property
    def is_image(self) -> bool:
        return self.mime.startswith("image/")

    @property
    def is_pdf(self) -> bool:
        return self.mime == "application/pdf"


@dataclass
class Note:
    title: str
    enml: str
    created: datetime | None = None
    updated: datetime | None = None
    tags: list[str] = field(default_factory=list)
    source_url: str | None = None
    author: str | None = None
    resources: list[Resource] = field(default_factory=list)
    index: int = 0  # position within the export file

    @property
    def id(self) -> str:
        """Stable identifier derived from note content.

        ENEX files carry no note GUIDs, so we fingerprint the note. The same
        note exported twice yields the same id, which lets re-runs skip notes
        that were already analyzed or uploaded.
        """
        h = hashlib.sha1()
        h.update(self.title.encode("utf-8"))
        h.update(b"\0")
        h.update((self.created.isoformat() if self.created else "").encode())
        h.update(b"\0")
        h.update(self.enml.encode("utf-8"))
        for r in self.resources:
            h.update(r.hash.encode())
        return h.hexdigest()[:16]

    def resource_by_hash(self, md5: str) -> Resource | None:
        md5 = md5.lower()
        for r in self.resources:
            if r.hash == md5:
                return r
        return None


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    value = value.strip()
    for fmt in ("%Y%m%dT%H%M%SZ", "%Y%m%dT%H%M%S"):
        try:
            return datetime.strptime(value, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def _text(el, tag: str) -> str | None:
    child = el.find(tag)
    if child is None or child.text is None:
        return None
    return child.text.strip() or None


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value else None
    except ValueError:
        return None


def _parse_resource(el, position: int) -> Resource | None:
    data_el = el.find("data")
    if data_el is None or not data_el.text:
        return None
    data = base64.b64decode("".join(data_el.text.split()))
    mime = _text(el, "mime") or "application/octet-stream"
    filename = None
    attrs = el.find("resource-attributes")
    if attrs is not None:
        filename = _text(attrs, "file-name")
    if not filename:
        ext = mimetypes.guess_extension(mime) or ".bin"
        filename = f"attachment-{position + 1}{ext}"
    return Resource(
        data=data,
        mime=mime,
        filename=filename,
        hash=hashlib.md5(data).hexdigest(),
        width=_int(_text(el, "width")),
        height=_int(_text(el, "height")),
    )


def _parse_note(el, index: int) -> Note:
    attrs = el.find("note-attributes")
    return Note(
        title=_text(el, "title") or "Untitled",
        enml=(el.find("content").text or "") if el.find("content") is not None else "",
        created=_parse_date(_text(el, "created")),
        updated=_parse_date(_text(el, "updated")),
        tags=[t.text.strip() for t in el.findall("tag") if t.text and t.text.strip()],
        source_url=_text(attrs, "source-url") if attrs is not None else None,
        author=_text(attrs, "author") if attrs is not None else None,
        resources=[
            r
            for i, res_el in enumerate(el.findall("resource"))
            if (r := _parse_resource(res_el, i)) is not None
        ],
        index=index,
    )


def iter_notes(path: str | Path) -> Iterator[Note]:
    """Yield notes one at a time so multi-gigabyte exports don't fill memory."""
    index = 0
    for event, el in iterparse(str(path), events=("end",), forbid_dtd=False):
        if el.tag == "note":
            yield _parse_note(el, index)
            index += 1
            el.clear()


def count_notes(path: str | Path) -> int:
    count = 0
    for _, el in iterparse(str(path), events=("end",), forbid_dtd=False):
        if el.tag == "note":
            count += 1
        if el.tag in ("note", "resource", "data"):
            el.clear()
    return count
