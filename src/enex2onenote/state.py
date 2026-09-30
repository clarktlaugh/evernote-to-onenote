"""On-disk progress tracking so every run can be stopped and resumed safely."""

from __future__ import annotations

import csv
import json
import os
import threading
from collections import Counter
from pathlib import Path

REVIEW_FIELDS = ["note_id", "index", "skip", "final_title", "final_tags", "ai_title",
                 "original_title", "evernote_tags", "summary", "confidence", "created",
                 "attachments"]


def _read_jsonl(path: Path) -> dict[str, dict]:
    records: dict[str, dict] = {}
    if path.exists():
        with path.open(encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    rec = json.loads(line)
                    records[rec["note_id"]] = rec
    return records


def state_dir_for(root: Path, enex_path: Path) -> Path:
    """<root>/<folder>/<file name>: exports with the same name in different folders
    (one folder per OneNote notebook) must not share progress or review sheets."""
    enex_path = Path(enex_path).resolve()
    return root / enex_path.parent.name / enex_path.stem


class NotebookState:
    """Per-.enex working directory: analysis cache, review sheet, upload log."""

    def __init__(self, root: Path, enex_path: Path):
        self.root = root
        self.dir = state_dir_for(root, enex_path)
        self.dir.mkdir(parents=True, exist_ok=True)
        self._claim(Path(enex_path).resolve())
        self.analysis_path = self.dir / "analysis.jsonl"
        self.uploads_path = self.dir / "uploads.jsonl"
        self.review_path = self.dir / "review.csv"
        self.preview_dir = self.dir / "preview"
        self._lock = threading.Lock()
        self.analyses = _read_jsonl(self.analysis_path)
        self.uploads = _read_jsonl(self.uploads_path)

    def _claim(self, enex_path: Path) -> None:
        """Record which export this directory belongs to, and refuse a different one.

        Two folders with the same name in different places would otherwise map to
        the same directory. A moved or renamed export takes the directory over,
        since the file it was recorded for no longer exists.
        """
        source = self.dir / "source.txt"
        if source.exists():
            previous = Path(source.read_text(encoding="utf-8").strip())
            if (os.path.normcase(str(previous)) != os.path.normcase(str(enex_path))
                    and previous.is_file()):
                raise SystemExit(
                    f"{self.dir} already holds the progress for\n  {previous}\n"
                    f"so it can't also be used for\n  {enex_path}\n"
                    "Rename one of the two folders or files, or pass a different --state-dir.")
        source.write_text(str(enex_path) + "\n", encoding="utf-8")

    def _append(self, path: Path, rec: dict) -> None:
        with self._lock, path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    def record_analysis(self, rec: dict) -> None:
        self.analyses[rec["note_id"]] = rec
        self._append(self.analysis_path, rec)

    def record_upload(self, rec: dict) -> None:
        self.uploads[rec["note_id"]] = rec
        self._append(self.uploads_path, rec)

    def is_uploaded(self, note_id: str) -> bool:
        return self.uploads.get(note_id, {}).get("status") == "ok"

    # -- review sheet ------------------------------------------------------------

    def read_review(self) -> dict[str, dict]:
        if not self.review_path.exists():
            return {}
        with self.review_path.open(encoding="utf-8-sig", newline="") as f:
            return {row["note_id"]: row for row in csv.DictReader(f) if row.get("note_id")}

    def write_review(self, reset_ids: set[str] = frozenset()) -> None:
        """Write review.csv, keeping edits the person already made to existing rows."""
        existing = self.read_review()
        rows = []
        for rec in sorted(self.analyses.values(), key=lambda r: r.get("index", 0)):
            a = rec.get("analysis") or {}
            row = {
                "note_id": rec["note_id"],
                "index": rec.get("index", ""),
                "skip": "",
                # Left blank when analysis failed: the note isn't uploaded until it
                # is analyzed or the person types a title here themselves.
                "final_title": a.get("title", ""),
                "final_tags": "; ".join(a.get("tags", [])),
                "ai_title": a.get("title", "") if a else f"ERROR: {rec.get('error', '')}",
                "original_title": rec.get("original_title", ""),
                "evernote_tags": "; ".join(rec.get("evernote_tags", [])),
                "summary": a.get("summary", ""),
                "confidence": a.get("confidence", ""),
                "created": rec.get("created", ""),
                "attachments": rec.get("attachments", ""),
            }
            old = existing.get(rec["note_id"])
            if old and rec["note_id"] not in reset_ids:
                for k in ("skip", "final_title", "final_tags"):
                    row[k] = old.get(k, row[k])
            rows.append(row)
        tmp = self.review_path.with_suffix(".tmp")
        # utf-8-sig so Excel opens accented characters correctly.
        with tmp.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=REVIEW_FIELDS)
            w.writeheader()
            w.writerows(rows)
        tmp.replace(self.review_path)


def tag_vocabulary(root: Path, limit: int = 250) -> list[str]:
    """Most-used tags across every notebook processed so far, for consistency."""
    counts: Counter[str] = Counter()
    for path in root.glob("**/analysis.jsonl"):
        for rec in _read_jsonl(path).values():
            counts.update((rec.get("analysis") or {}).get("tags", []))
        review = path.parent / "review.csv"
        if review.exists():
            with review.open(encoding="utf-8-sig", newline="") as f:
                for row in csv.DictReader(f):
                    counts.update(t.strip() for t in (row.get("final_tags") or "").split(";")
                                  if t.strip())
    return sorted(t for t, _ in counts.most_common(limit))
