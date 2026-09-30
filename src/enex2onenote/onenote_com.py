"""Create OneNote pages through the OneNote desktop app's COM API (Windows only).

This needs no Azure app registration and has none of Microsoft Graph's request
size limits. OneNote copies attachments into the notebook itself and syncs
the notebook to OneDrive (if the notebook lives there) on its own.

Pages are written in OneNote's 2013 page XML schema:
https://learn.microsoft.com/office/client-developer/onenote/onenote-developer-reference
"""

from __future__ import annotations

import base64
import logging
import re
import struct
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timezone
from html import escape as html_escape
from pathlib import Path
from typing import Callable
from xml.sax.saxutils import escape as xml_escape, quoteattr

from bs4 import Comment, NavigableString, Tag

from .content import _soup, page_header_items
from .enex import Note, Resource

log = logging.getLogger(__name__)

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_ONE = f"{{{NS}}}"

# OneNote COM enums
HS_NOTEBOOKS = 2
HS_SECTIONS = 3
CFT_NOTEBOOK = 1
CFT_SECTION = 3
# Microsoft OneNote 15.0 Type Library (OneNote 2013 and later): GUID, LCID, major, minor
ONENOTE_TYPELIB = ("{0EA692EE-BB50-4E3C-AEF0-356D91732725}", 0, 1, 1)
TYPE_E_LIBNOTREGISTERED = -2147319779

# Image types OneNote can embed directly from <one:Data>.
COM_IMAGE_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/gif", "image/bmp"}
MAX_IMAGE_WIDTH_PT = 500  # keep big photos from overflowing the page
PRINTOUT_DPI = 110

_DROP = {"script", "style", "form", "input", "button", "select", "textarea", "noscript",
         "object", "embed", "applet", "meta", "link", "head", "title"}
_BLOCKS = {"div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "section",
           "article", "header", "footer", "center", "dl", "dt", "dd", "address", "figure",
           "en-note", "body", "html", "thead", "tbody", "tfoot", "caption"}
_TAG_STYLES = {
    "b": "font-weight:bold", "strong": "font-weight:bold",
    "i": "font-style:italic", "em": "font-style:italic",
    "u": "text-decoration:underline",
    "s": "text-decoration:line-through", "strike": "text-decoration:line-through",
    "del": "text-decoration:line-through",
    "code": "font-family:Consolas", "tt": "font-family:Consolas", "kbd": "font-family:Consolas",
    "sup": "vertical-align:super", "sub": "vertical-align:sub",
    "h1": "font-size:16pt;font-weight:bold", "h2": "font-size:14pt;font-weight:bold",
    "h3": "font-size:12pt;font-weight:bold", "h4": "font-weight:bold",
    "h5": "font-weight:bold", "h6": "font-weight:bold", "th": "font-weight:bold",
}
_KEEP_CSS = {"color", "background-color", "font-weight", "font-style", "text-decoration",
             "font-family", "font-size", "vertical-align"}


# --------------------------------------------------------------------------
# Page XML building
# --------------------------------------------------------------------------

@dataclass
class OE:
    """One outline element: a paragraph, image, table or file, plus nested children."""

    content: str
    bullet: bool = False
    children: list["OE"] = field(default_factory=list)

    def xml(self) -> str:
        lst = '<one:List><one:Bullet bullet="2"/></one:List>' if self.bullet else ""
        kids = ""
        if self.children:
            kids = f"<one:OEChildren>{''.join(c.xml() for c in self.children)}</one:OEChildren>"
        return f"<one:OE>{lst}{self.content}{kids}</one:OE>"


def _t(html: str) -> str:
    # <one:T> holds a small subset of HTML (spans, links) as text.
    return f"<one:T>{xml_escape(html)}</one:T>"


def text_oe(html: str) -> OE:
    return OE(_t(html))


def _merge_css(*styles: str) -> str:
    props: dict[str, str] = {}
    for style in styles:
        for decl in (style or "").split(";"):
            if ":" in decl:
                k, v = decl.split(":", 1)
                k, v = k.strip().lower(), v.strip()
                if k in _KEEP_CSS and v and "<" not in v and '"' not in v:
                    props[k] = v
    return ";".join(f"{k}:{v}" for k, v in props.items())


def _span(text: str, style: str, href: str | None) -> str:
    out = html_escape(text, quote=False)
    if style:
        out = f'<span style="{html_escape(style)}">{out}</span>'
    if href:
        out = f'<a href="{html_escape(href)}">{out}</a>'
    return out


def image_size(data: bytes) -> tuple[int, int] | None:
    """Pixel size of a PNG, GIF or JPEG, read from its header."""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", data[16:24])
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", data[6:10])
        if data[:2] == b"\xff\xd8":
            i = 2
            while i + 9 < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                length = struct.unpack(">H", data[i + 2:i + 4])[0]
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                i += 2 + length
    except struct.error:
        pass
    return None


def image_oe(data: bytes, width_hint: str | int | None = None) -> OE:
    size = ""
    px = image_size(data)
    if px and px[0] > 0:
        w_pt = px[0] * 0.75  # 96 dpi pixels -> points
        try:
            if width_hint:
                w_pt = float(str(width_hint).rstrip("px")) * 0.75
        except ValueError:
            pass
        w_pt = min(w_pt, MAX_IMAGE_WIDTH_PT)
        h_pt = w_pt * px[1] / px[0]
        size = f'<one:Size width="{w_pt:.1f}" height="{h_pt:.1f}" isSetByUser="true"/>'
    return OE(f"<one:Image>{size}<one:Data>{base64.b64encode(data).decode()}</one:Data></one:Image>")


def file_oe(path: Path, name: str) -> OE:
    return OE(f"<one:InsertedFile pathSource={quoteattr(str(path))} "
              f"preferredName={quoteattr(name)}/>")


def pdf_printout_pngs(data: bytes, max_pages: int) -> list[bytes]:
    """Render PDF pages to PNG with PyMuPDF (optional dependency)."""
    try:
        import pymupdf
    except ImportError:
        log.warning("PyMuPDF is not installed; PDFs are attached without page images. "
                    "Install it with: pip install pymupdf")
        return []
    try:
        with pymupdf.open(stream=data, filetype="pdf") as doc:
            return [page.get_pixmap(dpi=PRINTOUT_DPI).tobytes("png")
                    for page in list(doc)[:max_pages]]
    except Exception as e:
        log.warning("Could not render PDF pages: %s", e)
        return []


class _Converter:
    """Walks ENML and produces OneNote outline elements."""

    def __init__(self, note: Note, render_media: Callable[[Resource, dict], list[OE]]):
        self.note = note
        self.render_media = render_media
        self.referenced: set[str] = set()

    def convert(self, node: Tag, style: str = "", prefix: str = "") -> list[OE]:
        out: list[OE] = []
        buf: list[str] = [html_escape(prefix, quote=False)] if prefix else []

        def flush() -> None:
            html = "".join(buf).strip()
            if html:
                out.append(text_oe(html))
            buf.clear()

        def walk(n: Tag, style: str, href: str | None) -> None:
            for child in n.children:
                if isinstance(child, Comment):
                    continue
                if isinstance(child, NavigableString):
                    text = re.sub(r"\s+", " ", str(child))
                    if text.strip() or buf:
                        buf.append(_span(text, style, href))
                    continue
                if not isinstance(child, Tag):
                    continue
                name = child.name
                if name in _DROP:
                    continue
                if name == "en-media":
                    flush()
                    out.extend(self._media(child))
                elif name == "en-todo":
                    buf.append("☑ " if child.get("checked") == "true" else "☐ ")
                elif name == "en-crypt":
                    flush()
                    out.append(text_oe("[Encrypted Evernote content was not migrated]"))
                elif name == "br":
                    if buf:
                        flush()
                    else:
                        out.append(text_oe(""))  # keep deliberate blank lines
                elif name == "hr":
                    flush()
                elif name == "img":
                    src = child.get("src", "")
                    if src.startswith("http"):
                        buf.append(_span(f"[image: {child.get('alt') or src}]", style, src))
                elif name in ("ul", "ol"):
                    flush()
                    out.extend(self._list(child, name == "ol", style))
                elif name == "table":
                    flush()
                    table = self._table(child, style)
                    if table:
                        out.append(table)
                elif name == "pre":
                    flush()
                    mono = _merge_css(style, "font-family:Consolas")
                    for line in child.get_text().splitlines():
                        out.append(text_oe(_span(line, mono, None) if line else ""))
                elif name == "a":
                    walk(child, _merge_css(style, child.get("style", "")),
                         child.get("href") or href)
                else:
                    new_style = _merge_css(style, _TAG_STYLES.get(name, ""),
                                           child.get("style", ""))
                    if name in _BLOCKS:
                        flush()
                        walk(child, new_style, href)
                        flush()
                    else:
                        walk(child, new_style, href)

        walk(node, style, None)
        flush()
        return out

    def _media(self, el: Tag) -> list[OE]:
        res = self.note.resource_by_hash(el.get("hash", ""))
        if res is None:
            return []
        self.referenced.add(res.hash)
        return self.render_media(res, dict(el.attrs))

    def _list(self, el: Tag, ordered: bool, style: str) -> list[OE]:
        items = []
        for n, li in enumerate(el.find_all("li", recursive=False), start=1):
            oes = self.convert(li, style, prefix=f"{n}. " if ordered else "")
            if not oes:
                continue
            first = oes[0]
            first.bullet = not ordered
            first.children.extend(oes[1:])
            items.append(first)
        return items

    def _table(self, el: Tag, style: str) -> OE | None:
        rows = [tr for tr in el.find_all("tr") if tr.find_parent("table") is el]
        grid = [tr.find_all(["td", "th"], recursive=False) for tr in rows]
        ncols = max((len(r) for r in grid), default=0)
        if not ncols:
            return None
        width = max(60.0, 480.0 / ncols)
        cols = "".join(f'<one:Column index="{i}" width="{width:.1f}"/>' for i in range(ncols))
        rows_xml = []
        for cells in grid:
            cell_xml = []
            for i in range(ncols):
                oes: list[OE] = []
                if i < len(cells):
                    c = cells[i]
                    oes = self.convert(c, _merge_css(style, _TAG_STYLES.get(c.name, "")))
                oes = oes or [text_oe("")]
                cell_xml.append("<one:Cell><one:OEChildren>"
                                + "".join(o.xml() for o in oes)
                                + "</one:OEChildren></one:Cell>")
            rows_xml.append(f"<one:Row>{''.join(cell_xml)}</one:Row>")
        return OE(f'<one:Table bordersVisible="true"><one:Columns>{cols}</one:Columns>'
                  f"{''.join(rows_xml)}</one:Table>")


def build_page_xml(note: Note, page_id: str, title: str, tags: list[str], summary: str,
                   attach_dir: Path, pdf_printouts: bool = True, include_summary: bool = True,
                   printout_max_pages: int = 30) -> str:
    """Build the XML for UpdatePageContent.

    Attachments are written to `attach_dir`; OneNote copies them into the notebook
    during UpdatePageContent, so the directory can be deleted afterwards.
    """
    written: dict[str, Path] = {}

    def save(res: Resource) -> Path:
        if res.hash not in written:
            safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", res.filename).strip() or "attachment"
            path = attach_dir / f"{len(written)}-{safe}"
            path.write_bytes(res.data)
            written[res.hash] = path
        return written[res.hash]

    def render(res: Resource, attrs: dict) -> list[OE]:
        if res.mime in COM_IMAGE_TYPES:
            return [image_oe(res.data, attrs.get("width"))]
        oes = [file_oe(save(res), res.filename)]
        if res.is_pdf and pdf_printouts:
            oes += [image_oe(png) for png in pdf_printout_pngs(res.data, printout_max_pages)]
        return oes

    conv = _Converter(note, render)
    soup = _soup(note.enml)
    body = conv.convert(soup.find("en-note") or soup)
    unreferenced = [r for r in note.resources if r.hash not in conv.referenced]
    if unreferenced:
        body.append(text_oe('<span style="font-weight:bold">Attachments</span>'))
        for r in unreferenced:
            body.extend(render(r, {}))

    header: list[OE] = []
    items = page_header_items(note, title, tags)
    if items:
        header.append(text_oe(f'<span style="color:#595959;font-size:9pt">'
                              f'{" &#183; ".join(items)}</span>'))
    if include_summary and summary:
        header.append(text_oe(f'<span style="font-style:italic">'
                              f"{html_escape(summary, quote=False)}</span>"))

    created = ""
    if note.created:
        stamp = note.created.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
        created = f' dateTime="{stamp}"'
    outline = "".join(o.xml() for o in header + body) or text_oe("").xml()
    return (f'<?xml version="1.0"?>'
            f'<one:Page xmlns:one="{NS}" ID={quoteattr(page_id)}{created}>'
            f"<one:Title><one:OE>{_t(html_escape(title, quote=False))}</one:OE></one:Title>"
            f"<one:Outline><one:OEChildren>{outline}</one:OEChildren></one:Outline>"
            f"</one:Page>")


# --------------------------------------------------------------------------
# COM automation
# --------------------------------------------------------------------------

def _dispatch():
    try:
        import pywintypes
        import win32com.client
    except ImportError:
        if sys.platform != "win32":
            raise SystemExit("Uploading through the OneNote desktop app (--target com, the "
                             "default) needs Windows. On other systems use --target graph; "
                             "see README.md for its one-time setup.")
        raise SystemExit("--target com needs the OneNote desktop app and pywin32:\n"
                         '  pip install -e ".[com]"')
    # Early binding (makepy) is required: the API returns values through
    # out-parameters, which late-bound Dispatch can't handle. Generate the
    # wrapper from the type library first; EnsureDispatch(progid) needs the
    # running app to describe itself, which fails on some installs.
    app, last = None, None
    try:
        module = win32com.client.gencache.EnsureModule(*ONENOTE_TYPELIB)
        if module is not None:
            app = module.Application()
    except Exception as e:  # pywintypes.com_error
        last = e
    for progid in ("OneNote.Application", "OneNote.Application.15"):
        if app is not None:
            break
        try:
            app = win32com.client.gencache.EnsureDispatch(progid)
        except Exception as e:  # pywintypes.com_error, AttributeError from a stale cache
            last = e
    if app is None:
        raise SystemExit(f"Could not start OneNote through COM ({last}). Make sure the OneNote "
                         "desktop app (Microsoft 365 / OneNote 2016) is installed and has been "
                         "opened once. If this persists, delete the folder "
                         "%LOCALAPPDATA%\\Temp\\gen_py and try again.")
    try:  # the first real call surfaces broken registrations (see _libnotregistered_help)
        app.GetHierarchy("", HS_NOTEBOOKS)
    except pywintypes.com_error as e:
        if e.hresult == TYPE_E_LIBNOTREGISTERED:
            raise SystemExit(_libnotregistered_help()) from e
        raise SystemExit(f"OneNote started but did not respond to COM calls ({e}).") from e
    return app


def _libnotregistered_help() -> str:
    guid, _, major, minor = ONENOTE_TYPELIB
    base = f"HKCU:\\Software\\Classes\\TypeLib\\{guid}"
    exe = r"C:\Program Files\Microsoft Office\Root\Office16\ONENOTE.EXE\3"
    try:
        import winreg

        key = f"TypeLib\\{guid}\\{major}.{minor}\\0\\win32"
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key) as k:
            exe = winreg.QueryValue(k, None) or exe
    except OSError:
        pass
    # Click-to-Run Office registers the library as version 1.1 for Win32 only and
    # leaves an empty 1.0 key that OneNote itself looks up. Point both at the file.
    entries = [f"{major}.{minor}\\0\\win64", "1.0\\0\\win64", "1.0\\0\\win32"]
    fix = "\n".join(f"  New-Item -Force '{base}\\{e}' -Value '{exe}'" for e in entries)
    return (
        "OneNote's COM type library registration is incomplete, so OneNote can't answer "
        "automation calls (\"Library not registered\"). This is a known problem with "
        "Click-to-Run Office installs. To fix it for your Windows account only (no admin "
        "rights needed), run this in PowerShell, then fully restart OneNote:\n\n"
        f"{fix}\n\n"
        "To undo it later:\n\n"
        f"  Remove-Item '{base}\\{major}.{minor}\\0\\win64'\n"
        f"  Remove-Item -Recurse '{base}\\1.0'")


def _str(value) -> str:
    """COM out-parameters sometimes come back inside a tuple."""
    if isinstance(value, tuple):
        return next((v for v in value if isinstance(v, str) and v), "")
    return value or ""


def _any_time():
    """OLE date 0 (1899-12-30), which tells OneNote not to check a page's last-modified time.

    makepy generates a plain tuple as the default for these DATE parameters,
    which current pywin32 rejects, so the value has to be passed explicitly. It
    must be timezone-aware: a naive pre-1970 datetime fails local-time
    conversion on Windows (OSError 22).
    """
    try:
        import pywintypes
    except ImportError:  # only the test doubles run without pywin32
        return None
    return pywintypes.Time(datetime(1899, 12, 30, tzinfo=timezone.utc))


class OneNoteCom:
    def __init__(self, app=None):
        self.app = app if app is not None else _dispatch()

    def _hierarchy(self, start: str, scope: int) -> ET.Element:
        return ET.fromstring(_str(self.app.GetHierarchy(start, scope)))

    def _open(self, path: str, relative_to: str, create_type: int) -> str:
        try:
            return _str(self.app.OpenHierarchy(path, relative_to, cftIfNotExist=create_type))
        except TypeError:
            return _str(self.app.OpenHierarchy(path, relative_to, "", create_type))

    def notebooks(self) -> list[dict]:
        root = self._hierarchy("", HS_NOTEBOOKS)
        return [{"id": nb.get("ID"), "name": nb.get("nickname") or nb.get("name"),
                 "path": nb.get("path")} for nb in root.iter(f"{_ONE}Notebook")]

    def find_notebook(self, name: str) -> dict | None:
        """An open notebook by its display name (case-insensitive).

        Notebooks are never created here: COM can only create them from a folder,
        which OneNote then treats as local files. Notebooks made in OneNote with
        File > New > OneDrive are synced by OneNote itself.
        """
        name = name.strip().lower()
        return next((nb for nb in self.notebooks() if (nb["name"] or "").lower() == name), None)

    def ensure_section(self, notebook_id: str, name: str) -> dict:
        root = self._hierarchy(notebook_id, HS_SECTIONS)
        for sec in root.iter(f"{_ONE}Section"):
            if (sec.get("name") or "").lower() == name.lower():
                return {"id": sec.get("ID"), "name": sec.get("name")}
        return {"id": self._open(f"{name}.one", notebook_id, CFT_SECTION), "name": name}

    def create_page(self, section_id: str, build_xml: Callable[[str], str]) -> str:
        page_id = _str(self.app.CreateNewPage(section_id))
        try:
            self.app.UpdatePageContent(build_xml(page_id), _any_time())
        except Exception:
            try:  # don't leave an empty "Untitled page" behind
                self.app.DeleteHierarchy(page_id, _any_time())
            except Exception:
                pass
            raise
        return page_id

    def link(self, object_id: str) -> str | None:
        try:
            return _str(self.app.GetHyperlinkToObject(object_id, "")) or None
        except Exception:
            return None
