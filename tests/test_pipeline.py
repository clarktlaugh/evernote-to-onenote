import csv
import json
import re

import pytest

from enex2onenote import ai, cli
from enex2onenote.content import extract_for_analysis, note_text
from enex2onenote.enex import Resource, count_notes, iter_notes
from enex2onenote.onenote import MAX_DATA_PARTS, plan_page


def test_parse_enex(sample_enex):
    notes = list(iter_notes(sample_enex))
    assert count_notes(sample_enex) == 3
    assert [n.title for n in notes] == ["Scan_0012.pdf", "Untitled", "Recipe & notes"]
    assert notes[0].tags == ["Bills"]
    assert notes[0].created.year == 2021
    assert notes[0].source_url == "https://example.com/x"
    assert notes[1].resources[0].filename == "attachment-1.png"
    assert len({n.id for n in notes}) == 3
    # ids are stable across parses
    assert [n.id for n in iter_notes(sample_enex)] == [n.id for n in notes]


def test_extracts_pdf_text_and_flags_scans(sample_enex):
    first, second, _ = iter_notes(sample_enex)
    ex = extract_for_analysis(first, 20000)
    assert "Comcast invoice March 2021" in ex.text
    assert not ex.scanned_pdfs
    ex2 = extract_for_analysis(second, 20000)
    assert [r.filename for r in ex2.scanned_pdfs] == ["unreferenced.pdf"]
    assert len(ex2.images) == 1
    assert "Photo" in note_text(second)


def test_truncation_is_flagged(sample_enex):
    first = next(iter_notes(sample_enex))
    ex = extract_for_analysis(first, 10)
    assert ex.truncated and len(ex.text) == 10


def test_plan_page_html(sample_enex):
    _, second, _ = iter_notes(sample_enex)
    plan = plan_page(second, "Kitchen photo", ["home", "photos"], "A photo.")
    html = plan.html
    assert "<title>Kitchen photo</title>" in html
    assert 'name="created" content="2021-03-14T10:15:00+00:00"' in html
    assert "#home #photos" in html and "Evernote title: Untitled" in html
    assert "<script" not in html and "en-media" not in html
    assert re.search(r'<img src="name:part0" alt="attachment-1.png" width="300"/>', html)
    assert "☑ done" in html and "☐ open" in html
    # unreferenced PDF is attached at the end, with a printout
    assert 'data-attachment="unreferenced.pdf"' in html
    assert '<img data-render-src="name:part1"' in html
    assert set(plan.first_parts) == {"part0", "part1"}
    assert not plan.patches and not plan.onedrive


def test_plan_page_splits_large_notes():
    from enex2onenote.enex import Note

    resources = [Resource(data=bytes([i]) * 1_000_000, mime="image/jpeg",
                          filename=f"p{i}.jpg", hash=f"{i:032x}") for i in range(8)]
    resources.append(Resource(data=b"x" * 5_000_000, mime="application/pdf",
                              filename="huge.pdf", hash="f" * 32))
    body = "".join(f'<en-media hash="{r.hash}" type="{r.mime}"/>' for r in resources)
    note = Note(title="Many photos", enml=f"<en-note>{body}</en-note>", resources=resources)
    plan = plan_page(note, "Many photos", [], "")
    assert len(plan.first_parts) == 3  # 3 MB fits under the 4 MB request cap
    moved = [p for _, parts in plan.patches for p in parts]
    assert len(moved) == 5
    for cmds, parts in plan.patches:
        assert len(parts) <= MAX_DATA_PARTS
        assert sum(r.size for r in parts.values()) < 4_000_000
        for c in cmds:
            assert c["action"] == "insert" and f'data-id="{c["target"][1:]}"' in plan.html
    assert [r.filename for _, r in plan.onedrive] == ["huge.pdf"]
    assert 'data-id="od-0"' in plan.html


def test_analysis_normalization():
    a = ai.Analysis.from_dict({"title": ' "Comcast bill - March 2021." ',
                               "tags": ["#Bills", "Comcast", "bills", "home internet"],
                               "summary": "s", "original_title_ok": False, "confidence": "high"})
    assert a.title == "Comcast bill - March 2021"
    assert a.tags == ["bills", "comcast", "home-internet"]
    with pytest.raises(ai.AnalysisError):
        ai.Analysis.from_dict({"title": "  "})


def test_prompt_includes_vocabulary_and_escapes(sample_enex):
    note = list(iter_notes(sample_enex))[2]
    p = ai.build_prompt(note, "a <b> c", False, ["bills", "recipes"], "Prefix titles with year")
    assert "bills, recipes" in p and "Prefix titles with year" in p
    assert "Recipe &amp; notes" in p and "a &lt;b&gt; c" in p


class FakeAnalyzer(ai.Analyzer):
    name = "fake"

    def __init__(self):
        super().__init__("fake-model")
        self.seen_vocab = None

    def analyze(self, note, vocabulary):
        self.seen_vocab = vocabulary
        if note.title == "Untitled":
            raise ai.AnalysisError("boom")
        return ai.Analysis(title=f"AI {note.title}", tags=["bills"], summary="sum",
                           original_title_ok=False, confidence="high", model="fake")


def test_analyze_review_and_dry_run(sample_enex, tmp_path, monkeypatch, capsys):
    fake = FakeAnalyzer()
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: fake)
    state_dir = tmp_path / "state"
    base = ["--state-dir", str(state_dir)]

    cli.main(base + ["analyze", str(sample_enex), "--workers", "2"])
    review_path = state_dir / "Household Bills" / "review.csv"
    rows = list(csv.DictReader(review_path.open(encoding="utf-8-sig")))
    assert [r["final_title"] for r in rows] == ["AI Scan_0012.pdf", "", "AI Recipe & notes"]
    assert rows[1]["ai_title"].startswith("ERROR")

    # Person edits the sheet: rename one note, skip another.
    rows[0]["final_title"] = "My edited title"
    rows[2]["skip"] = "y"
    with review_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)

    # Re-running retries only the failure and keeps the edits.
    cli.main(base + ["analyze", str(sample_enex)])
    rows = list(csv.DictReader(review_path.open(encoding="utf-8-sig")))
    assert rows[0]["final_title"] == "My edited title" and rows[2]["skip"] == "y"
    assert fake.seen_vocab == ["bills"]
    out = capsys.readouterr().out
    assert "Analyzed 1 note(s), 1 failed" in out

    cli.main(base + ["upload", str(sample_enex), "--dry-run"])
    previews = sorted((state_dir / "Household Bills" / "preview").glob("*.html"))
    assert len(previews) == 1
    assert "<title>My edited title</title>" in previews[0].read_text()
    out = capsys.readouterr().out
    assert "Previewed 1, skipped 1, failed 0, not analyzed 1" in out

    # Typing a title for the failed note in the sheet makes it uploadable.
    rows[1]["final_title"] = "Kitchen photo"
    with review_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)
    cli.main(base + ["upload", str(sample_enex), "--dry-run"])
    assert "Previewed 2, skipped 1" in capsys.readouterr().out


def test_claude_code_analyzer_parses_output(sample_enex, monkeypatch):
    import subprocess

    calls = {}

    def fake_run(cmd, input, **kw):
        calls["cmd"], calls["input"] = cmd, input
        out = {"type": "result", "subtype": "success", "is_error": False, "result": "",
               "structured_output": {"title": "Comcast bill - March 2021", "tags": ["bills"],
                                     "summary": "x", "original_title_ok": False,
                                     "confidence": "high"},
               "modelUsage": {"claude-sonnet-5-5": {}}}
        return subprocess.CompletedProcess(cmd, 0, json.dumps(out), "")

    monkeypatch.setattr(ai.shutil, "which", lambda _: "/usr/bin/claude")
    monkeypatch.setattr(ai.subprocess, "run", fake_run)
    analyzer = ai.ClaudeCodeAnalyzer(None)
    result = analyzer.analyze(next(iter_notes(sample_enex)), [])
    assert result.title == "Comcast bill - March 2021" and result.model == "claude-sonnet-5-5"
    assert "--json-schema" in calls["cmd"] and calls["cmd"][calls["cmd"].index("--tools") + 1] == ""
    assert "Comcast invoice" in calls["input"]
