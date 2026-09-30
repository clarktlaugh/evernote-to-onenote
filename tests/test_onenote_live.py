"""End-to-end tests against the real OneNote desktop app (Windows only, opt-in).

Each test opens a throwaway local notebook in a temporary folder (the tool
itself never creates notebooks), uploads through the CLI with `--target com`,
reads the pages back from OneNote, then closes the notebook and deletes its
files. Your other notebooks are not touched.

    $env:ENEX2ONENOTE_LIVE_ONENOTE = "1"     # PowerShell
    pytest -m onenote
"""

import os
import shutil
import sys
import time
import uuid
import xml.etree.ElementTree as ET

import pytest

from enex2onenote import cli
from enex2onenote.onenote_com import CFT_NOTEBOOK, HS_SECTIONS, NS, OneNoteCom, _str

pytestmark = [
    pytest.mark.onenote,
    pytest.mark.skipif(sys.platform != "win32", reason="OneNote COM needs Windows"),
    pytest.mark.skipif(not os.environ.get("ENEX2ONENOTE_LIVE_ONENOTE"),
                       reason="set ENEX2ONENOTE_LIVE_ONENOTE=1 to test against the real OneNote app"),
]

ONE = f"{{{NS}}}"
HS_PAGES = 4


@pytest.fixture
def live(tmp_path, monkeypatch):
    """A OneNote connection plus a unique notebook name, cleaned up afterwards."""
    from test_pipeline import FakeAnalyzer

    pytest.importorskip("win32com.client")
    try:
        com = OneNoteCom()
    except SystemExit as e:  # the tool's own diagnosis, e.g. how to fix the registry
        pytest.fail(str(e), pytrace=False)
    # The tool never creates notebooks, so make a throwaway local one here. OneNote
    # only creates a notebook in a folder that already exists.
    name = f"enex2onenote-test-{uuid.uuid4().hex[:8]}"
    path = tmp_path / "notebooks" / name
    path.mkdir(parents=True)
    nb_id = com._open(str(path), "", CFT_NOTEBOOK)
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: FakeAnalyzer())

    def run(*argv):
        cli.main(["--state-dir", str(tmp_path / "state"), *argv])

    def upload(enex, *extra):
        run("analyze", str(enex))
        run("upload", str(enex), "--target", "com", "--notebook", name, *extra)

    yield com, name, upload

    com.app.CloseNotebook(nb_id)
    for _ in range(10):  # OneNote can hold the .one files briefly after closing
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists():
            break
        time.sleep(0.5)


def notebook_pages(com, name) -> dict[str, list[tuple[str, ET.Element]]]:
    """{section name: [(page title, page XML), ...]} read back from OneNote."""
    nb = next(nb for nb in com.notebooks() if nb["name"] == name)
    out = {}
    for sec in com._hierarchy(nb["id"], HS_SECTIONS).iter(f"{ONE}Section"):
        pages = com._hierarchy(sec.get("ID"), HS_PAGES).iter(f"{ONE}Page")
        out[sec.get("name")] = [
            (p.get("name"), ET.fromstring(_str(com.app.GetPageContent(p.get("ID")))))
            for p in pages]
    return out


def page_text(page: ET.Element) -> str:
    return "\n".join(t.text or "" for t in page.iter(f"{ONE}T"))


def test_upload_creates_real_pages(live, sample_enex, capsys):
    com, name, upload = live
    upload(sample_enex)
    assert "Uploaded 2, skipped 0, failed 0, not analyzed 1" in capsys.readouterr().out

    pages = notebook_pages(com, name)
    assert list(pages) == ["Household Bills"]
    titles = [title for title, _ in pages["Household Bills"]]
    assert titles == ["AI Scan_0012.pdf", "AI Recipe & notes"]

    scan = pages["Household Bills"][0][1]
    assert scan.get("dateTime", "").startswith("2021-03-14")  # Evernote creation date kept
    text = page_text(scan)
    assert "#bills" in text and "Evernote title: Scan_0012.pdf" in text and "sum" in text
    assert "See attached" in text
    files = [f.get("preferredName") for f in scan.iter(f"{ONE}InsertedFile")]
    assert files == ["Scan_0012.pdf"]  # OneNote copied the attachment into the notebook
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        pass
    else:
        assert scan.find(f".//{ONE}Image") is not None  # PDF printout

    # Re-running uploads nothing new.
    upload(sample_enex)
    assert "Uploaded 0" in capsys.readouterr().out
    assert len(notebook_pages(com, name)["Household Bills"]) == 2


def test_merge_exports_into_one_section(live, sample_enex, second_enex):
    com, name, upload = live
    upload(sample_enex, "--section", "Everything")
    upload(second_enex, "--section", "Everything")

    pages = notebook_pages(com, name)
    assert list(pages) == ["Everything"]
    assert [t for t, _ in pages["Everything"]] == [
        "AI Scan_0012.pdf", "AI Recipe & notes", "AI IMG_2231", "AI Note"]


def test_folder_of_exports(live, sample_enex, second_enex, tmp_path):
    """<folder>/<Notebook>/<Section>.enex: the folder name picks the real notebook."""
    com, name, _ = live
    notebook_dir = tmp_path / "Exports" / name
    notebook_dir.mkdir(parents=True)
    for enex in (sample_enex, second_enex):
        enex.rename(notebook_dir / enex.name)
    base = ["--state-dir", str(tmp_path / "state")]
    cli.main(base + ["migrate", str(tmp_path / "Exports")])

    pages = notebook_pages(com, name)
    assert {s: [t for t, _ in p] for s, p in pages.items()} == {
        "Household Bills": ["AI Scan_0012.pdf", "AI Recipe & notes"],
        "Recipes": ["AI IMG_2231", "AI Note"],
    }
