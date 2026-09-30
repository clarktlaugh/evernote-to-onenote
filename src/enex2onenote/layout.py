"""Map a folder of Evernote exports onto OneNote notebooks and sections.

One level only:

    Exports/                  <- pass this: every notebook folder inside
      Finances/               <- OneNote notebook "Finances"
        Receipts.enex         <- section "Receipts", one page per note
        Taxes.enex
      Family/
        Recipes.enex

or pass a single notebook folder (Exports/Finances) to do just that notebook.
Anything else (.enex files deeper down, or loose .enex files next to notebook
folders) is refused rather than skipped, so no export is silently left behind.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from .onenote import safe_name


@dataclass
class Job:
    enex: Path
    notebook: str | None = None  # None: a single file, so --notebook decides
    section: str | None = None

    @property
    def label(self) -> str:
        return f"{self.notebook} / {self.section}" if self.notebook else self.enex.name


def _enex_files(folder: Path) -> list[Path]:
    return sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".enex"),
                  key=lambda p: p.name.lower())


def _deeper_enex(folder: Path) -> list[Path]:
    """.enex files inside subfolders of `folder`, at any depth."""
    return sorted(p for sub in folder.iterdir() if sub.is_dir()
                  for p in sub.rglob("*") if p.is_file() and p.suffix.lower() == ".enex")


def _refuse(problem: str, paths: list[Path], root: Path) -> None:
    listed = "\n".join(f"  {p.relative_to(root)}" for p in paths[:20])
    more = f"\n  ... and {len(paths) - 20} more" if len(paths) > 20 else ""
    raise SystemExit(
        f"{problem}\n{listed}{more}\n"
        "Expected one level of folders: <folder>/<Notebook>/<Section>.enex "
        "(or a single notebook folder containing .enex files).")


def plan(path: Path) -> list[Job]:
    """The exports to process, in order. A single .enex file is one job."""
    path = Path(path)
    if path.is_file():
        return [Job(path)]
    if not path.is_dir():
        raise SystemExit(f"File or folder not found: {path}")

    loose = _enex_files(path)
    if loose:  # a single notebook folder
        deeper = _deeper_enex(path)
        if deeper:
            _refuse(f"{path.name} contains .enex files and also folders with .enex files in "
                    "them. Use either a notebook folder or a folder of notebook folders:",
                    deeper, path)
        jobs = [Job(p, path.name, p.stem) for p in loose]
    else:
        jobs, too_deep = [], []
        for sub in sorted((d for d in path.iterdir() if d.is_dir()), key=lambda d: d.name.lower()):
            jobs += [Job(p, sub.name, p.stem) for p in _enex_files(sub)]
            too_deep += _deeper_enex(sub)
        if too_deep:
            _refuse("These .enex files are more than one folder deep:", too_deep, path)
        if not jobs:
            raise SystemExit(f"No .enex files found in {path} or its folders.")

    _check_sections(jobs)
    return jobs


def _check_sections(jobs: list[Job]) -> None:
    """Two exports whose names clean up to the same section would be merged."""
    by_section: dict[tuple[str, str], list[Job]] = defaultdict(list)
    for job in jobs:
        if job.notebook:
            by_section[(job.notebook.lower(), safe_name(job.section).lower())].append(job)
    clashes = [group for group in by_section.values() if len(group) > 1]
    if clashes:
        lines = "\n".join(
            f"  {g[0].notebook}: " + ", ".join(j.enex.name for j in g)
            + f" -> section {safe_name(g[0].section)!r}" for g in clashes)
        raise SystemExit("These exports would end up in the same OneNote section. Rename "
                         f"them so their names differ in the first 50 characters:\n{lines}")


def notebooks(jobs: list[Job]) -> list[str]:
    """Distinct notebook names, in order."""
    seen: dict[str, str] = {}
    for job in jobs:
        if job.notebook:
            seen.setdefault(job.notebook.lower(), job.notebook)
    return list(seen.values())
