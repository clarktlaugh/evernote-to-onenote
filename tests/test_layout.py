"""Folder layout: <folder>/<Notebook>/<Section>.enex, one level only."""

import pytest

from conftest import note_xml, write_enex

from enex2onenote import layout


def make(root, *paths):
    for p in paths:
        path = root / p
        path.parent.mkdir(parents=True, exist_ok=True)
        write_enex(path, [note_xml(path.stem, "<p>x</p>")])


def test_folder_of_notebook_folders(tmp_path):
    make(tmp_path, "Finances/Taxes.enex", "Finances/receipts.ENEX", "Family/Recipes.enex")
    (tmp_path / "Finances" / "notes.txt").write_text("not an export")
    (tmp_path / "Empty").mkdir()
    jobs = layout.plan(tmp_path)
    assert [(j.notebook, j.section) for j in jobs] == [
        ("Family", "Recipes"), ("Finances", "receipts"), ("Finances", "Taxes")]
    assert layout.notebooks(jobs) == ["Family", "Finances"]
    assert jobs[0].label == "Family / Recipes"


def test_single_notebook_folder(tmp_path):
    make(tmp_path, "Finances/Taxes.enex", "Finances/Receipts.enex")
    jobs = layout.plan(tmp_path / "Finances")
    assert [(j.notebook, j.section) for j in jobs] == [
        ("Finances", "Receipts"), ("Finances", "Taxes")]


def test_single_file(tmp_path):
    make(tmp_path, "Taxes.enex")
    (job,) = layout.plan(tmp_path / "Taxes.enex")
    assert job.notebook is None and job.section is None
    assert layout.notebooks([job]) == []


def test_refuses_exports_more_than_one_level_deep(tmp_path):
    make(tmp_path, "Finances/Taxes.enex", "Finances/2019/Receipts.enex")
    with pytest.raises(SystemExit, match="more than one folder deep") as exc:
        layout.plan(tmp_path)
    assert "Receipts.enex" in str(exc.value)


def test_refuses_loose_files_next_to_notebook_folders(tmp_path):
    make(tmp_path, "Stray.enex", "Finances/Taxes.enex")
    with pytest.raises(SystemExit, match="also folders with .enex files") as exc:
        layout.plan(tmp_path)
    assert "Taxes.enex" in str(exc.value)


def test_refuses_empty_or_missing(tmp_path):
    (tmp_path / "Empty").mkdir()
    with pytest.raises(SystemExit, match="No .enex files"):
        layout.plan(tmp_path)
    with pytest.raises(SystemExit, match="not found"):
        layout.plan(tmp_path / "nope")


def test_refuses_exports_that_map_to_the_same_section(tmp_path):
    long = "x" * 50
    make(tmp_path, "Finances/Q&A.enex", "Finances/Q#A.enex",
         f"Family/{long}-one.enex", f"Family/{long}-two.enex", "Travel/Q&A.enex")
    with pytest.raises(SystemExit, match="same OneNote section") as exc:
        layout.plan(tmp_path)
    msg = str(exc.value)
    assert "Finances: Q#A.enex, Q&A.enex -> section 'Q-A'" in msg
    assert f"Family: {long}-one.enex, {long}-two.enex" in msg
    assert "Travel" not in msg  # the same name in another notebook is fine
