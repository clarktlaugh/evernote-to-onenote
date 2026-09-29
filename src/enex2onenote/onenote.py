"""Microsoft Graph client for creating OneNote notebooks, sections and pages."""

from __future__ import annotations

import json
import logging
import re
import time
import uuid
from dataclasses import dataclass, field
from html import escape
from pathlib import Path
from urllib.parse import quote

import requests

from .content import ONENOTE_IMAGE_TYPES, enml_to_html
from .enex import Note, Resource

log = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Notes.ReadWrite", "Files.ReadWrite"]

# Graph rejects requests over 4 MB and allows at most 6 parts (Presentation or
# Commands + 5 data parts). Stay under the size cap with some headroom.
MAX_REQUEST_BYTES = 3_700_000
MAX_DATA_PARTS = 5
MAX_RENDERED_PDFS = 5  # data-render-src images allowed per request
# Anything bigger than this can't go through the OneNote API at all, so it is
# uploaded to OneDrive and linked from the page instead.
MAX_INLINE_RESOURCE = 3_400_000

_BAD_NAME_CHARS = re.compile(r'[?*\\/:<>|&#"%~]')


def safe_name(name: str, limit: int = 50) -> str:
    name = _BAD_NAME_CHARS.sub("-", name).strip(" .-") or "Imported"
    return name[:limit].rstrip()


class GraphError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(f"Graph API {status}: {message}")
        self.status = status


# --------------------------------------------------------------------------
# Authentication
# --------------------------------------------------------------------------

class Auth:
    """MSAL device-code login with a persistent token cache."""

    def __init__(self, client_id: str, tenant: str, cache_path: Path):
        import msal

        self.cache_path = cache_path
        self.cache = msal.SerializableTokenCache()
        if cache_path.exists():
            self.cache.deserialize(cache_path.read_text())
        self.app = msal.PublicClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=self.cache)

    def _save(self) -> None:
        if self.cache.has_state_changed:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(self.cache.serialize())
            try:
                self.cache_path.chmod(0o600)
            except OSError:
                pass

    def token(self, interactive: bool = True) -> str:
        result = None
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            if not interactive:
                raise SystemExit("Not signed in to Microsoft. Run `enex2onenote login` first.")
            flow = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise SystemExit(f"Could not start Microsoft sign-in: {flow.get('error_description')}")
            print(flow["message"], flush=True)
            result = self.app.acquire_token_by_device_flow(flow)
        self._save()
        if "access_token" not in result:
            raise SystemExit(f"Microsoft sign-in failed: {result.get('error_description')}")
        return result["access_token"]

    def account_name(self) -> str | None:
        accounts = self.app.get_accounts()
        return accounts[0].get("username") if accounts else None


# --------------------------------------------------------------------------
# Graph client
# --------------------------------------------------------------------------

class Graph:
    def __init__(self, auth: Auth):
        self.auth = auth
        self.session = requests.Session()

    def request(self, method: str, url: str, *, retries: int = 6, **kw) -> requests.Response:
        if not url.startswith("http"):
            url = GRAPH + url
        delay = 2.0
        extra_headers = kw.pop("headers", None) or {}
        for attempt in range(retries + 1):
            headers = {**extra_headers, "Authorization": f"Bearer {self.auth.token()}"}
            try:
                resp = self.session.request(method, url, headers=headers, timeout=300, **kw)
            except requests.ConnectionError as e:
                if attempt == retries:
                    raise GraphError(0, str(e)) from e
                time.sleep(delay)
                delay *= 2
                continue
            if resp.status_code in (429, 500, 502, 503, 504) and attempt < retries:
                wait = float(resp.headers.get("Retry-After", delay))
                log.info("Graph %s; retrying in %.0fs", resp.status_code, wait)
                time.sleep(wait)
                delay = min(delay * 2, 120)
                continue
            if resp.status_code >= 400:
                raise GraphError(resp.status_code, _error_message(resp))
            return resp
        raise GraphError(0, "retries exhausted")

    # -- notebooks & sections -------------------------------------------------

    def _find(self, url: str, name: str) -> dict | None:
        items = self.request("GET", url, params={"$select": "id,displayName"}).json()["value"]
        return next((i for i in items if i["displayName"].lower() == name.lower()), None)

    def ensure_notebook(self, name: str) -> dict:
        name = safe_name(name, 128)
        nb = self._find("/me/onenote/notebooks", name)
        if nb:
            return nb
        log.info("Creating notebook %r", name)
        return self.request("POST", "/me/onenote/notebooks", json={"displayName": name}).json()

    def ensure_section(self, notebook_id: str, name: str) -> dict:
        name = safe_name(name)
        url = f"/me/onenote/notebooks/{notebook_id}/sections"
        sec = self._find(url, name)
        if sec:
            return sec
        log.info("Creating section %r", name)
        return self.request("POST", url, json={"displayName": name}).json()

    def list_notebooks(self) -> list[dict]:
        return self.request("GET", "/me/onenote/notebooks",
                            params={"$select": "id,displayName"}).json()["value"]

    # -- pages -----------------------------------------------------------------

    def create_page(self, section_id: str, html: str, parts: dict[str, Resource]) -> dict:
        files = [("Presentation", (None, html.encode("utf-8"), "text/html; charset=utf-8"))]
        files += [(name, (name, r.data, r.mime)) for name, r in parts.items()]
        return self.request("POST", f"/me/onenote/sections/{section_id}/pages", files=files).json()

    def patch_page(self, page_id: str, commands: list[dict], parts: dict[str, Resource]) -> None:
        files = [("Commands", (None, json.dumps(commands).encode(), "application/json"))]
        files += [(name, (name, r.data, r.mime)) for name, r in parts.items()]
        # A freshly created page can briefly 404 while OneNote finishes creating it.
        for attempt in range(5):
            try:
                self.request("PATCH", f"/me/onenote/pages/{page_id}/content", files=files)
                return
            except GraphError as e:
                if e.status != 404 or attempt == 4:
                    raise
                time.sleep(3 * (attempt + 1))

    # -- OneDrive (for attachments too large for the OneNote API) --------------

    def upload_to_onedrive(self, folder: str, resource: Resource) -> str:
        path = "/".join(quote(safe_name(p, 100)) for p in folder.split("/"))
        fname = quote(safe_name(resource.filename, 150))
        session = self.request(
            "POST", f"/me/drive/root:/{path}/{fname}:/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": "rename"}}).json()
        upload_url = session["uploadUrl"]
        chunk = 320 * 1024 * 16  # must be a multiple of 320 KiB
        data, total, item = resource.data, resource.size, None
        for start in range(0, total, chunk):
            end = min(start + chunk, total)
            resp = requests.put(upload_url, data=data[start:end], timeout=300, headers={
                "Content-Length": str(end - start),
                "Content-Range": f"bytes {start}-{end - 1}/{total}"})
            if resp.status_code >= 400:
                raise GraphError(resp.status_code, _error_message(resp))
            if resp.status_code in (200, 201):
                item = resp.json()
        if not item:
            raise GraphError(0, "OneDrive upload did not complete")
        return item.get("webUrl", "")


def _error_message(resp: requests.Response) -> str:
    try:
        body = resp.json()
        err = body.get("error", {})
        msg = err.get("message") or json.dumps(body)[:500]
        return msg
    except ValueError:
        return resp.text[:500]


# --------------------------------------------------------------------------
# Page building
# --------------------------------------------------------------------------

@dataclass
class PagePlan:
    """A page's HTML plus the resources it needs, split into Graph-sized requests."""

    html: str
    first_parts: dict[str, Resource]
    # Each follow-up batch: (commands, parts) for a PATCH request.
    patches: list[tuple[list[dict], dict[str, Resource]]] = field(default_factory=list)
    onedrive: list[tuple[str, Resource]] = field(default_factory=list)  # (placeholder id, res)


def _size_label(n: int) -> str:
    return f"{n / 1024 / 1024:.1f} MB" if n >= 1024 * 1024 else f"{max(1, n // 1024)} KB"


class _Batcher:
    """Assigns each resource to the first request that has room for it."""

    def __init__(self):
        self.batches: list[dict] = [self._new()]

    @staticmethod
    def _new() -> dict:
        return {"parts": {}, "bytes": 0, "renders": 0}

    def place(self, res: Resource, render: bool) -> int:
        for i, b in enumerate(self.batches):
            if (len(b["parts"]) < MAX_DATA_PARTS
                    and b["bytes"] + res.size <= MAX_REQUEST_BYTES
                    and (not render or b["renders"] < MAX_RENDERED_PDFS)):
                return i
        self.batches.append(self._new())
        return len(self.batches) - 1


def plan_page(note: Note, title: str, tags: list[str], summary: str,
              pdf_printouts: bool = True, include_summary: bool = True) -> PagePlan:
    batcher = _Batcher()
    part_names: dict[str, str] = {}  # resource hash -> part name
    deferred: list[tuple[int, str, str]] = []  # (batch, placeholder id, fragment)
    onedrive: list[tuple[str, Resource]] = []
    placed: dict[str, tuple[int, str]] = {}  # hash -> (batch, part name)

    def fragment_for(res: Resource, part: str, attrs: dict) -> str:
        if res.mime in ONENOTE_IMAGE_TYPES:
            size = "".join(f' {k}="{escape(str(attrs[k]))}"' for k in ("width",) if attrs.get(k))
            return f'<img src="name:{part}" alt="{escape(res.filename)}"{size}/>'
        html = (f'<object data-attachment="{escape(res.filename)}" data="name:{part}" '
                f'type="{escape(res.mime)}"/>')
        if res.is_pdf and pdf_printouts:
            html += f'<img data-render-src="name:{part}" alt="{escape(res.filename)}"/>'
        return html

    def render(res: Resource, attrs: dict) -> str:
        if res.size > MAX_INLINE_RESOURCE:
            pid = f"od-{len(onedrive)}"
            onedrive.append((pid, res))
            return (f'<p data-id="{pid}">\U0001F4CE {escape(res.filename)} '
                    f'({_size_label(res.size)}) - too large for OneNote, stored in OneDrive</p>')
        render_pdf = res.is_pdf and pdf_printouts
        if res.hash in placed:  # same file referenced twice: reuse its part
            batch, part = placed[res.hash]
        else:
            batch = batcher.place(res, render_pdf)
            part = f"part{len(placed)}"
            b = batcher.batches[batch]
            b["parts"][part] = res
            b["bytes"] += res.size
            b["renders"] += int(render_pdf)
            placed[res.hash] = (batch, part)
        frag = fragment_for(res, part, attrs)
        if batch == 0:
            return frag
        pid = f"res-{len(deferred)}"
        deferred.append((batch, pid, frag))
        return f'<p data-id="{pid}">\U0001F4CE {escape(res.filename)}</p>'

    body, unreferenced = enml_to_html(note, render)
    extra = "".join(render(r, {}) for r in unreferenced)
    if extra:
        body += f"<div><p><b>Attachments</b></p>{extra}</div>"

    meta = []
    if tags:
        meta.append("Tags: " + " ".join(f"#{escape(t)}" for t in tags))
    if note.title.strip() != title.strip():
        meta.append(f"Evernote title: {escape(note.title)}")
    if note.created:
        meta.append(f"Created {note.created.date().isoformat()}")
    if note.source_url:
        meta.append(f'<a href="{escape(note.source_url)}">Source</a>')
    header = ""
    if meta:
        header += f'<p style="color:#595959;font-size:9pt">{" &#183; ".join(meta)}</p>'
    if include_summary and summary:
        header += f"<p><i>{escape(summary)}</i></p>"
    created = (f'<meta name="created" content="{note.created.isoformat()}"/>'
               if note.created else "")
    html = (f"<!DOCTYPE html><html><head><title>{escape(title)}</title>{created}</head>"
            f"<body>{header}{body}</body></html>")

    patches = []
    for i, b in enumerate(batcher.batches[1:], start=1):
        cmds = [{"target": f"#{pid}", "action": "insert", "position": "after", "content": frag}
                for batch, pid, frag in deferred if batch == i]
        # A reused part must be sent in every request that references it.
        needed = set(re.findall(r'name:(part\d+)', "".join(c["content"] for c in cmds)))
        parts = {n: r for bb in batcher.batches for n, r in bb["parts"].items() if n in needed}
        if cmds:
            patches.append((cmds, parts))
    return PagePlan(html=html, first_parts=batcher.batches[0]["parts"],
                    patches=patches, onedrive=onedrive)


def upload_page(graph: Graph, section_id: str, plan: PagePlan, onedrive_folder: str) -> dict:
    try:
        page = graph.create_page(section_id, plan.html, plan.first_parts)
    except GraphError as e:
        # PDF printouts occasionally fail to render; retry with plain attachments.
        if e.status == 400 and "data-render-src" in plan.html:
            html = re.sub(r'<img data-render-src="name:part\d+"[^>]*/>', "", plan.html)
            page = graph.create_page(section_id, html, plan.first_parts)
        else:
            raise
    page_id = page["id"]
    for cmds, parts in plan.patches:
        try:
            graph.patch_page(page_id, cmds, parts)
        except GraphError as e:
            if e.status == 400 and any("data-render-src" in c["content"] for c in cmds):
                for c in cmds:
                    c["content"] = re.sub(r'<img data-render-src="name:part\d+"[^>]*/>', "",
                                          c["content"])
                graph.patch_page(page_id, cmds, parts)
            else:
                raise
    for pid, res in plan.onedrive:
        url = graph.upload_to_onedrive(onedrive_folder, res)
        graph.patch_page(page_id, [{
            "target": f"#{pid}", "action": "insert", "position": "after",
            "content": f'<p><a href="{escape(url)}">Open {escape(res.filename)} '
                       f'({_size_label(res.size)}) in OneDrive</a></p>'}], {})
    return page
