"""Command-line interface: enex2onenote analyze | upload | migrate | login | notebooks."""

from __future__ import annotations

import argparse
import logging
import os
import sys
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from datetime import datetime, timezone
from pathlib import Path

from . import __version__, layout
from .ai import AnalysisError, make_analyzer
from .enex import Note, count_notes, iter_notes
from .state import NotebookState, tag_vocabulary

DEFAULT_STATE_DIR = "migration-state"
DEFAULT_NOTEBOOK = "Evernote Import"
TOKEN_CACHE = Path.home() / ".enex2onenote" / "msal_token_cache.json"


def _say(msg: str) -> None:
    print(msg, flush=True)


def _short(s: str, n: int = 45) -> str:
    s = " ".join(s.split())
    return s if len(s) <= n else s[: n - 1] + "…"


# --------------------------------------------------------------------------
# analyze
# --------------------------------------------------------------------------

def run_analyze(args) -> dict[str, int]:
    enex = Path(args.enex)
    state = NotebookState(Path(args.state_dir), enex)
    instructions = args.instructions
    if args.instructions_file:
        instructions = Path(args.instructions_file).read_text(encoding="utf-8")
    analyzer = make_analyzer(args.backend, args.model, args.effort, args.max_chars, instructions)
    vocabulary = tag_vocabulary(Path(args.state_dir))
    total = count_notes(enex)
    limit = args.limit or total
    _say(f"Analyzing {min(limit, total)} of {total} notes in {enex.name} with "
         f"backend '{analyzer.name}'" + (f" (model {analyzer.model})" if analyzer.model else ""))

    def todo(note: Note) -> bool:
        rec = state.analyses.get(note.id)
        if args.reanalyze:
            return True
        return not rec or "analysis" not in rec  # re-try earlier failures

    def work(note: Note) -> tuple[Note, dict]:
        rec = {
            "note_id": note.id, "index": note.index, "original_title": note.title,
            "evernote_tags": note.tags,
            "created": note.created.isoformat() if note.created else "",
            "attachments": "; ".join(r.filename for r in note.resources),
            "analyzed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        }
        try:
            rec["analysis"] = analyzer.analyze(note, vocabulary).to_dict()
        except AnalysisError as e:
            rec["error"] = str(e)
        return note, rec

    done = failed = 0
    reset: set[str] = set()
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        pending = set()

        def drain(block_until_below: int) -> None:
            nonlocal done, failed, pending
            while len(pending) > block_until_below:
                finished, pending = wait(pending, return_when=FIRST_COMPLETED)
                for fut in finished:
                    note, rec = fut.result()
                    state.record_analysis(rec)
                    reset.add(note.id)
                    done += 1
                    prefix = f"[{note.index + 1}/{total}]"
                    if "analysis" in rec:
                        a = rec["analysis"]
                        tags = " ".join("#" + t for t in a["tags"])
                        _say(f"{prefix} {_short(note.title, 35)!r} -> {a['title']!r} {tags}")
                    else:
                        failed += 1
                        _say(f"{prefix} {_short(note.title, 35)!r} FAILED: {rec['error']}")

        for note in iter_notes(enex):
            if note.index >= limit:
                break
            if not todo(note):
                continue
            pending.add(pool.submit(work, note))
            drain(args.workers * 2)  # bound memory: few notes held at once
        drain(0)

    state.write_review(reset_ids=reset if args.reanalyze else set())
    _say(f"\nAnalyzed {done} note(s), {failed} failed. "
         f"{len(state.analyses) - done} already done from earlier runs.")
    _say(f"Review and edit titles/tags (or mark skip=y) in:\n  {state.review_path}")
    if failed:
        _say("Re-run the same command to retry failed notes.")
    return {"analyzed": done, "analysis failed": failed}


# --------------------------------------------------------------------------
# upload
# --------------------------------------------------------------------------

def _graph(args, interactive: bool = True):
    from .onenote import Auth, Graph

    client_id = args.client_id or os.environ.get("ENEX2ONENOTE_CLIENT_ID")
    if not client_id:
        raise SystemExit("Set ENEX2ONENOTE_CLIENT_ID (or pass --client-id) to your Azure app "
                         "registration's Application (client) ID. See README.md.")
    tenant = args.tenant or os.environ.get("ENEX2ONENOTE_TENANT", "common")
    auth = Auth(client_id, tenant, TOKEN_CACHE)
    auth.token(interactive=interactive)
    return Graph(auth)


def _final_metadata(note: Note, state: NotebookState, review: dict, keep_good: bool):
    """Returns (title, tags, summary, skip) or None if the note hasn't been analyzed."""
    rec = state.analyses.get(note.id)
    a = (rec or {}).get("analysis")
    row = review.get(note.id)
    if row and (a or row.get("final_title", "").strip()):
        skip = row.get("skip", "").strip().lower() in ("y", "yes", "x", "1", "true", "skip")
        title = row.get("final_title", "").strip() or (a or {}).get("title") or note.title
        tags = [t.strip() for t in row.get("final_tags", "").split(";") if t.strip()]
        return title, tags, row.get("summary", ""), skip
    if not a:
        return None
    title = note.title if (keep_good and a.get("original_title_ok")) else a["title"]
    return title, a.get("tags", []), a.get("summary", ""), False


class UploadError(Exception):
    pass


class GraphUploader:
    """Uploads through Microsoft Graph (any OS; needs an Azure app registration)."""

    def __init__(self, args, section_name: str):
        from .onenote import safe_name

        self.args = args
        self.section_name = section_name
        self.section_suffix = 1
        self.onedrive_folder = (f"Evernote Attachments/{safe_name(args.notebook, 100)}/"
                                f"{safe_name(section_name, 100)}")
        self.graph = self.section = None

    def _plan(self, note, title, tags, summary):
        from .onenote import plan_page

        return plan_page(note, title, tags, summary, pdf_printouts=not self.args.no_pdf_printouts,
                         include_summary=not self.args.no_summary)

    def connect(self) -> None:
        self.graph = _graph(self.args)
        notebook = self.graph.ensure_notebook(self.args.notebook)
        self.notebook_id = notebook["id"]
        self.section = self.graph.ensure_section(self.notebook_id, self.section_name)
        _say(f"Uploading via Microsoft Graph to notebook {notebook['displayName']!r}, "
             f"section {self.section['displayName']!r}")

    def preview(self, note, title, tags, summary, out_base: Path) -> str:
        plan = self._plan(note, title, tags, summary)
        Path(f"{out_base}.html").write_text(plan.html, encoding="utf-8")
        extra = sum(len(p) for _, p in plan.patches)
        return (f"{len(plan.first_parts)} inline + {extra} follow-up attachment(s), "
                f"{len(plan.onedrive)} to OneDrive")

    def upload(self, note, title, tags, summary) -> dict:
        from .onenote import GraphError, upload_page

        plan = self._plan(note, title, tags, summary)
        while True:
            try:
                page = upload_page(self.graph, self.section["id"], plan, self.onedrive_folder)
                break
            except GraphError as e:
                if e.status != 507 or self.section_suffix >= 100:
                    raise UploadError(str(e)) from e
                # Section reached OneNote's page limit: continue in the next one. When
                # several exports share a section, "(2)" may already be full as well.
                self.section_suffix += 1
                try:
                    self.section = self.graph.ensure_section(
                        self.notebook_id, f"{self.section_name[:44]} ({self.section_suffix})")
                except GraphError as e2:
                    raise UploadError(str(e2)) from e2
                _say(f"Section full; continuing in {self.section['displayName']!r}")
        return {"page_id": page.get("id"),
                "url": (page.get("links") or {}).get("oneNoteWebUrl", {}).get("href"),
                "section": self.section["displayName"]}


def _com_notebook(args):
    """Connect to OneNote and find --notebook, which must already be open there."""
    from .onenote_com import OneNoteCom

    onenote = OneNoteCom()
    notebook = onenote.find_notebook(args.notebook)
    if notebook is None:
        _missing_notebooks(onenote, [args.notebook])
    return onenote, notebook


def _missing_notebooks(onenote, missing: list[str]) -> None:
    names = "\n".join(f"  {nb['name']}" for nb in onenote.notebooks()) or "  (none)"
    wanted = "\n".join(f"  {n}" for n in missing)
    raise SystemExit(
        f"OneNote has no open notebook named:\n{wanted}\n"
        "Create each one in OneNote first with File > New > OneDrive, so that OneNote "
        "syncs it to OneDrive itself. Use exactly these names (for a folder of exports, "
        "the notebook folder names), then run this again.\n"
        f"Notebooks open in OneNote:\n{names}")


def _preflight(args, jobs: list[layout.Job]) -> None:
    """Check the upload target before any work, so a long run can't fail partway."""
    if args.dry_run:
        return
    if args.target == "graph":
        _graph(args)  # sign in; Graph creates missing notebooks itself
        return
    from .onenote_com import OneNoteCom

    onenote = OneNoteCom()
    missing = [n for n in layout.notebooks(jobs) or [args.notebook]
               if onenote.find_notebook(n) is None]
    if missing:
        _missing_notebooks(onenote, missing)


class ComUploader:
    """Drives the OneNote desktop app through COM (Windows only)."""

    def __init__(self, args, section_name: str):
        self.args = args
        self.section_name = section_name
        self.onenote = self.section = None

    def _xml(self, note, title, tags, summary, page_id: str, attach_dir: Path) -> str:
        from .onenote_com import build_page_xml

        return build_page_xml(note, page_id, title, tags, summary, attach_dir,
                              pdf_printouts=not self.args.no_pdf_printouts,
                              include_summary=not self.args.no_summary,
                              printout_max_pages=self.args.printout_max_pages)

    def connect(self) -> None:
        from .onenote import safe_name

        self.onenote, notebook = _com_notebook(self.args)
        self.section = self.onenote.ensure_section(notebook["id"], safe_name(self.section_name))
        _say(f"Uploading via the OneNote app (COM) to notebook {notebook['name']!r}, "
             f"section {self.section['name']!r}")

    def preview(self, note, title, tags, summary, out_base: Path) -> str:
        attach_dir = Path(f"{out_base}-files")
        attach_dir.mkdir(exist_ok=True)
        xml = self._xml(note, title, tags, summary, "{preview-page-id}", attach_dir)
        Path(f"{out_base}.xml").write_text(xml, encoding="utf-8")
        files = len(list(attach_dir.iterdir()))
        if not files:
            attach_dir.rmdir()
        return f"{files} attached file(s)"

    def upload(self, note, title, tags, summary) -> dict:
        import tempfile

        with tempfile.TemporaryDirectory(prefix="enex2onenote-") as tmp:
            try:
                page_id = self.onenote.create_page(
                    self.section["id"],
                    lambda pid: self._xml(note, title, tags, summary, pid, Path(tmp)))
            except Exception as e:  # pywintypes.com_error carries the HRESULT
                raise UploadError(f"OneNote COM error: {e}") from e
        return {"page_id": page_id, "url": self.onenote.link(page_id),
                "section": self.section["name"]}


def run_upload(args) -> dict[str, int]:
    from .onenote import safe_name

    enex = Path(args.enex)
    state = NotebookState(Path(args.state_dir), enex)
    review = {} if args.ignore_review else state.read_review()
    section_name = args.section or enex.stem
    total = count_notes(enex)
    limit = args.limit or total

    uploader = (ComUploader if args.target == "com" else GraphUploader)(args, section_name)
    if args.dry_run:
        state.preview_dir.mkdir(exist_ok=True)
        _say(f"Dry run ({args.target}): writing page previews to {state.preview_dir}")
    else:
        uploader.connect()

    uploaded = skipped = failed = missing = 0
    for note in iter_notes(enex):
        if note.index >= limit:
            break
        prefix = f"[{note.index + 1}/{total}]"
        if state.is_uploaded(note.id) and not args.dry_run:
            continue
        meta = _final_metadata(note, state, review, args.keep_good_titles)
        if meta is None:
            missing += 1
            _say(f"{prefix} {_short(note.title)!r} has no AI title yet - skipping "
                 "(re-run analyze, or type a final_title in review.csv)")
            continue
        title, tags, summary, skip = meta
        if skip:
            skipped += 1
            continue
        if args.dry_run:
            out_base = state.preview_dir / f"{note.index + 1:05d}-{safe_name(title, 60)}"
            _say(f"{prefix} {title!r}: {uploader.preview(note, title, tags, summary, out_base)}")
            uploaded += 1
            continue
        try:
            info = uploader.upload(note, title, tags, summary)
        except UploadError as e:
            failed += 1
            state.record_upload({"note_id": note.id, "status": "error", "error": str(e)})
            _say(f"{prefix} {title!r} FAILED: {e}")
            continue
        uploaded += 1
        state.record_upload({
            "note_id": note.id, "status": "ok", "title": title, "target": args.target, **info,
            "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds")})
        _say(f"{prefix} uploaded {title!r}")

    verb = "Previewed" if args.dry_run else "Uploaded"
    _say(f"\n{verb} {uploaded}, skipped {skipped}, failed {failed}, not analyzed {missing}.")
    if failed:
        _say("Re-run the same command to retry failed notes; uploaded notes are not duplicated.")
    if missing:
        _say("Run `enex2onenote analyze` first for notes that were not analyzed.")
    return {verb.lower(): uploaded, "skipped": skipped, "upload failed": failed,
            "not analyzed": missing}


# --------------------------------------------------------------------------
# argument parsing
# --------------------------------------------------------------------------

def _add_ai_args(p):
    g = p.add_argument_group("AI analysis")
    g.add_argument("--backend", default="auto", choices=["auto", "api", "claude-code", "none"],
                   help="api = Claude API key; claude-code = your Claude subscription via the "
                        "`claude` CLI; none = keep Evernote titles. auto picks api if "
                        "ANTHROPIC_API_KEY is set, else claude-code.")
    g.add_argument("--model", help="Claude model (api default: claude-opus-5-5; claude-code "
                                   "default: your Claude Code default model)")
    g.add_argument("--effort", choices=["low", "medium", "high", "xhigh", "max"],
                   help="Reasoning effort (api default: low)")
    g.add_argument("--workers", type=int, default=4, help="Notes analyzed in parallel")
    g.add_argument("--max-chars", type=int, default=20000,
                   help="Characters of note/attachment text sent to Claude per note")
    g.add_argument("--instructions", help="Extra guidance, e.g. 'Start titles with YYYY-MM'")
    g.add_argument("--instructions-file", help="Read extra guidance from a file")
    g.add_argument("--reanalyze", action="store_true",
                   help="Re-analyze notes that already have results")


def _add_upload_args(p):
    g = p.add_argument_group("OneNote upload")
    g.add_argument("--target", choices=["com", "graph"], default="com",
                   help="com (default) = drive the OneNote desktop app (Windows only); "
                        "graph = Microsoft Graph web API (any OS, needs an Azure app "
                        "registration)")
    g.add_argument("--notebook",
                   help="OneNote notebook for a single .enex file. com: must already be open "
                        "in OneNote (create it with File > New > OneDrive); graph: created if "
                        f"missing (default {DEFAULT_NOTEBOOK!r}). For a folder, the notebook "
                        "folder names are used instead.")
    g.add_argument("--section", help="OneNote section for a single .enex file (default: the "
                                     "file name)")
    g.add_argument("--dry-run", action="store_true",
                   help="Write page previews locally instead of uploading (OneNote XML for "
                        "com, HTML for graph)")
    g.add_argument("--ignore-review", action="store_true",
                   help="Ignore review.csv edits and use the AI results directly")
    g.add_argument("--keep-good-titles", action="store_true",
                   help="Keep the Evernote title when Claude judged it accurate")
    g.add_argument("--no-pdf-printouts", action="store_true",
                   help="Attach PDFs only; don't render their pages as images on the page")
    g.add_argument("--no-summary", action="store_true", help="Don't add the AI summary line")
    g.add_argument("--printout-max-pages", type=int, default=30,
                   help="com only: PDF pages rendered as images per attachment")


def _add_ms_args(p):
    g = p.add_argument_group("Microsoft account")
    g.add_argument("--client-id", help="Azure app client ID (or ENEX2ONENOTE_CLIENT_ID)")
    g.add_argument("--tenant", help="common (default), consumers, organizations, or tenant ID "
                                    "(or ENEX2ONENOTE_TENANT)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="enex2onenote",
        description="Migrate Evernote .enex exports to OneNote, with Claude-generated titles and tags.")
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--state-dir", default=DEFAULT_STATE_DIR,
                        help=f"Where progress and review files live (default ./{DEFAULT_STATE_DIR})")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("login", help="Sign in to your Microsoft account")
    _add_ms_args(p)

    p = sub.add_parser("notebooks", help="List your OneNote notebooks")
    p.add_argument("--target", choices=["com", "graph"], default="com")
    _add_ms_args(p)

    def add_source(p):
        p.add_argument("enex", metavar="PATH",
                       help="An .enex file, or a folder of notebook folders "
                            "(<folder>/<Notebook>/<Section>.enex), or one notebook folder")
        p.add_argument("--limit", type=int,
                       help="Only process the first N notes of each .enex file (for trying it out)")

    p = sub.add_parser("analyze", help="Step 1: generate titles/tags and write review.csv")
    add_source(p)
    _add_ai_args(p)

    p = sub.add_parser("upload", help="Step 2: create OneNote pages from the reviewed results")
    add_source(p)
    _add_upload_args(p)
    _add_ms_args(p)

    p = sub.add_parser("migrate", help="Analyze and upload in one go (no review step)")
    add_source(p)
    _add_ai_args(p)
    _add_upload_args(p)
    _add_ms_args(p)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        # Titles can hold any character. When output goes to a file or pipe, Windows
        # would otherwise encode it with cp1252 and crash on the first unusual one.
        if hasattr(stream, "reconfigure") and not stream.isatty():
            stream.reconfigure(encoding="utf-8", errors="replace")
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format="%(levelname)s %(name)s: %(message)s")
    if not args.verbose:
        # pypdf warns about every quirk of a damaged PDF ("Ignoring wrong pointing
        # object ...") while still extracting its text. Show those only with -v.
        logging.getLogger("pypdf").setLevel(logging.ERROR)

    if args.command in ("analyze", "upload", "migrate"):
        run_jobs(args, layout.plan(Path(args.enex)))
    elif args.command == "login":
        graph = _graph(args)
        _say(f"Signed in as {graph.auth.account_name()}")
    elif args.command == "notebooks":
        if args.target == "com":
            from .onenote_com import OneNoteCom

            for nb in OneNoteCom().notebooks():
                _say(f"{nb['name']}  ({nb['path']})")
        else:
            for nb in _graph(args).list_notebooks():
                _say(nb["displayName"])
    return 0


def run_jobs(args, jobs: list[layout.Job]) -> None:
    """Run analyze/upload/migrate for one .enex file, or for each file in a folder."""
    analyze = args.command in ("analyze", "migrate")
    upload = args.command in ("upload", "migrate")
    folder = jobs[0].notebook is not None
    if upload:
        if folder and (args.notebook or args.section):
            raise SystemExit("--notebook and --section apply to a single .enex file. For a "
                             "folder, notebooks are named after the notebook folders and "
                             "sections after the .enex files.")
        args.notebook = args.notebook or DEFAULT_NOTEBOOK
    if folder:
        names = layout.notebooks(jobs)
        _say(f"{len(jobs)} export(s) in {len(names)} notebook(s):")
        for name in names:
            sections = [j.section for j in jobs if j.notebook == name]
            _say(f"  {name}: " + ", ".join(sections))
    if upload:
        _preflight(args, jobs)  # before analysis, which can take hours

    totals: dict[str, int] = {}
    for i, job in enumerate(jobs, start=1):
        job_args = argparse.Namespace(**vars(args))
        job_args.enex = str(job.enex)
        if folder:
            job_args.notebook, job_args.section = job.notebook, job.section
            _say(f"\n=== {job.label} ({i} of {len(jobs)}) ===")
        counts = {}
        if analyze:
            counts.update(run_analyze(job_args))
        if upload:
            if args.command == "migrate":
                job_args.ignore_review = False
            counts.update(run_upload(job_args))
        for k, v in counts.items():
            totals[k] = totals.get(k, 0) + v

    if folder:
        _say(f"\nAll {len(jobs)} exports: "
             + ", ".join(f"{k} {v}" for k, v in totals.items()) + ".")


if __name__ == "__main__":
    sys.exit(main())
