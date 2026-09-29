"""Command-line interface: enex2onenote analyze | upload | migrate | login | notebooks."""

from __future__ import annotations

import argparse
import logging
import os
import sys
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from datetime import datetime, timezone
from pathlib import Path

from . import __version__
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

def run_analyze(args) -> NotebookState:
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
    return state


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
    if row:
        skip = row.get("skip", "").strip().lower() in ("y", "yes", "x", "1", "true", "skip")
        title = row.get("final_title", "").strip() or note.title
        tags = [t.strip() for t in row.get("final_tags", "").split(";") if t.strip()]
        return title, tags, row.get("summary", ""), skip
    if not a:
        return None
    title = note.title if (keep_good and a.get("original_title_ok")) else a["title"]
    return title, a.get("tags", []), a.get("summary", ""), False


def run_upload(args) -> None:
    from .onenote import GraphError, plan_page, safe_name, upload_page

    enex = Path(args.enex)
    state = NotebookState(Path(args.state_dir), enex)
    review = {} if args.ignore_review else state.read_review()
    section_name = args.section or enex.stem
    total = count_notes(enex)
    limit = args.limit or total

    graph = section = None
    if args.dry_run:
        state.preview_dir.mkdir(exist_ok=True)
        _say(f"Dry run: writing page previews to {state.preview_dir}")
    else:
        graph = _graph(args)
        notebook = graph.ensure_notebook(args.notebook)
        section = graph.ensure_section(notebook["id"], section_name)
        _say(f"Uploading to OneNote notebook {notebook['displayName']!r}, "
             f"section {section['displayName']!r}")
    section_suffix = 1
    onedrive_folder = f"Evernote Attachments/{safe_name(args.notebook, 100)}/{safe_name(section_name, 100)}"

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
            _say(f"{prefix} {_short(note.title)!r} not analyzed yet - skipping")
            continue
        title, tags, summary, skip = meta
        if skip:
            skipped += 1
            continue
        plan = plan_page(note, title, tags, summary, pdf_printouts=not args.no_pdf_printouts,
                         include_summary=not args.no_summary)
        if args.dry_run:
            out = state.preview_dir / f"{note.index + 1:05d}-{safe_name(title, 60)}.html"
            out.write_text(plan.html, encoding="utf-8")
            extra = sum(len(p) for _, p in plan.patches)
            _say(f"{prefix} {title!r}: {len(plan.first_parts)} inline + {extra} follow-up "
                 f"attachment(s), {len(plan.onedrive)} to OneDrive")
            uploaded += 1
            continue
        try:
            try:
                page = upload_page(graph, section["id"], plan, onedrive_folder)
            except GraphError as e:
                if e.status != 507:
                    raise
                # Section reached OneNote's page limit: continue in a new section.
                section_suffix += 1
                section = graph.ensure_section(
                    graph.ensure_notebook(args.notebook)["id"],
                    f"{section_name[:44]} ({section_suffix})")
                _say(f"Section full; continuing in {section['displayName']!r}")
                page = upload_page(graph, section["id"], plan, onedrive_folder)
        except GraphError as e:
            failed += 1
            state.record_upload({"note_id": note.id, "status": "error", "error": str(e)})
            _say(f"{prefix} {title!r} FAILED: {e}")
            continue
        uploaded += 1
        state.record_upload({
            "note_id": note.id, "status": "ok", "title": title, "page_id": page.get("id"),
            "url": (page.get("links") or {}).get("oneNoteWebUrl", {}).get("href"),
            "section": section["displayName"],
            "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds")})
        _say(f"{prefix} uploaded {title!r}")

    verb = "Previewed" if args.dry_run else "Uploaded"
    _say(f"\n{verb} {uploaded}, skipped {skipped}, failed {failed}, not analyzed {missing}.")
    if failed:
        _say("Re-run the same command to retry failed notes; uploaded notes are not duplicated.")
    if missing:
        _say("Run `enex2onenote analyze` first for notes that were not analyzed.")


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
    g.add_argument("--notebook", default=DEFAULT_NOTEBOOK,
                   help=f"OneNote notebook (created if missing; default {DEFAULT_NOTEBOOK!r})")
    g.add_argument("--section", help="OneNote section (default: the .enex file name)")
    g.add_argument("--dry-run", action="store_true",
                   help="Write HTML previews locally instead of uploading")
    g.add_argument("--ignore-review", action="store_true",
                   help="Ignore review.csv edits and use the AI results directly")
    g.add_argument("--keep-good-titles", action="store_true",
                   help="Keep the Evernote title when Claude judged it accurate")
    g.add_argument("--no-pdf-printouts", action="store_true",
                   help="Attach PDFs only; don't render their pages as images on the page")
    g.add_argument("--no-summary", action="store_true", help="Don't add the AI summary line")


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
    _add_ms_args(p)

    p = sub.add_parser("analyze", help="Step 1: generate titles/tags and write review.csv")
    p.add_argument("enex")
    p.add_argument("--limit", type=int, help="Only process the first N notes (for trying it out)")
    _add_ai_args(p)

    p = sub.add_parser("upload", help="Step 2: create OneNote pages from the reviewed results")
    p.add_argument("enex")
    p.add_argument("--limit", type=int, help="Only process the first N notes")
    _add_upload_args(p)
    _add_ms_args(p)

    p = sub.add_parser("migrate", help="Analyze and upload in one go (no review step)")
    p.add_argument("enex")
    p.add_argument("--limit", type=int, help="Only process the first N notes")
    _add_ai_args(p)
    _add_upload_args(p)
    _add_ms_args(p)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format="%(levelname)s %(name)s: %(message)s")
    if getattr(args, "enex", None) and not Path(args.enex).is_file():
        raise SystemExit(f"File not found: {args.enex}")

    if args.command == "login":
        graph = _graph(args)
        _say(f"Signed in as {graph.auth.account_name()}")
    elif args.command == "notebooks":
        for nb in _graph(args).list_notebooks():
            _say(nb["displayName"])
    elif args.command == "analyze":
        run_analyze(args)
    elif args.command == "upload":
        run_upload(args)
    elif args.command == "migrate":
        if not args.dry_run:
            _graph(args)  # sign in up front, before spending time on analysis
        run_analyze(args)
        args.ignore_review = False
        run_upload(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
