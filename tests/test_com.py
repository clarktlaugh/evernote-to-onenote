import xml.etree.ElementTree as ET

import pytest

from enex2onenote import cli, onenote_com
from enex2onenote.enex import Note, iter_notes
from enex2onenote.onenote_com import NS, OneNoteCom, build_page_xml, image_size
from enex2onenote.state import state_dir_for

ONE = f"{{{NS}}}"


def texts(root):
    return [t.text or "" for t in root.iter(f"{ONE}T")]


def test_page_xml_structure(sample_enex, tmp_path):
    _, second, _ = iter_notes(sample_enex)
    xml = build_page_xml(second, "{page-1}", "Kitchen photo", ["home"], "A photo.", tmp_path,
                         pdf_printouts=False)
    root = ET.fromstring(xml)
    assert root.tag == f"{ONE}Page" and root.get("ID") == "{page-1}"
    assert root.get("dateTime") == "2021-03-14T10:15:00.000Z"
    # Title must come before the outline, per the OneNote schema
    assert [c.tag for c in root] == [f"{ONE}Title", f"{ONE}Outline"]
    assert root.find(f"{ONE}Title/{ONE}OE/{ONE}T").text == "Kitchen photo"
    t = texts(root)
    assert any("#home" in x and "Evernote title: Untitled" in x for x in t)
    assert any("A photo." in x for x in t)
    assert any("Photo" in x and 'font-weight:bold">here' in x for x in t)
    assert not any("alert" in x for x in t)
    assert any("\u2611 done" in x and "\u2610 open" in x for x in t)
    # PNG is embedded; width hint 300px -> 225pt
    img = root.find(f".//{ONE}Image")
    assert img.find(f"{ONE}Size").get("width") == "225.0"
    assert img.find(f"{ONE}Data").text
    # the unreferenced PDF is written to disk and inserted as a file
    f = root.find(f".//{ONE}InsertedFile")
    assert f.get("preferredName") == "unreferenced.pdf"
    assert (tmp_path / "0-unreferenced.pdf").exists()
    assert f.get("pathSource") == str(tmp_path / "0-unreferenced.pdf")
    # every OE holds at most one content element
    for oe in root.iter(f"{ONE}OE"):
        content = [c for c in oe if c.tag not in (f"{ONE}List", f"{ONE}OEChildren")]
        assert len(content) == 1


def test_lists_tables_and_escaping(tmp_path):
    enml = ("<en-note><h1>Head &amp; <i>more</i></h1>"
            "<ul><li>one<ul><li>nested</li></ul></li><li>two</li></ul>"
            "<ol><li>first</li><li>second</li></ol>"
            "<table><tr><th>A</th><td><b>B</b></td></tr><tr><td>1</td></tr></table>"
            "<div><a href=\"https://x.test/?a=1&amp;b=2\">link</a> ]]&gt; &lt;tag&gt;</div>"
            "<pre>line1\nline2</pre></en-note>")
    note = Note(title="t", enml=enml)
    root = ET.fromstring(build_page_xml(note, "{p}", "t", [], "", tmp_path))
    t = texts(root)
    assert '<span style="font-size:16pt;font-weight:bold">Head &amp; </span>' in t[1]
    bullets = [oe for oe in root.iter(f"{ONE}OE") if oe.find(f"{ONE}List") is not None]
    assert [b.find(f"{ONE}T").text for b in bullets] == ["one", "nested", "two"]
    assert bullets[0].find(f"{ONE}OEChildren/{ONE}OE/{ONE}T").text == "nested"
    assert "1. first" in t and "2. second" in t
    table = root.find(f".//{ONE}Table")
    assert len(table.findall(f"{ONE}Columns/{ONE}Column")) == 2
    rows = table.findall(f"{ONE}Row")
    assert len(rows) == 2 and all(len(r.findall(f"{ONE}Cell")) == 2 for r in rows)
    assert '<span style="font-weight:bold">A</span>' in t
    assert any('<a href="https://x.test/?a=1&amp;b=2">link</a> ]]&gt; &lt;tag&gt;' in x for x in t)
    assert '<span style="font-family:Consolas">line2</span>' in t


def test_image_size_parsing(sample_enex):
    from conftest import make_png

    assert image_size(make_png()) == (1, 1)
    jpeg = (b"\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
            b"\xff\xc0\x00\x11\x08\x00\x20\x00\x40\x03\x01\x22\x00\x02\x11\x01\x03\x11\x01")
    assert image_size(jpeg) == (64, 32)
    assert image_size(b"GIF89a\x10\x00\x08\x00") == (16, 8)
    assert image_size(b"junk") is None


def test_pdf_printouts(sample_enex, tmp_path):
    pytest.importorskip("pymupdf")
    first = next(iter_notes(sample_enex))
    root = ET.fromstring(build_page_xml(first, "{p}", "t", [], "", tmp_path))
    assert root.find(f".//{ONE}InsertedFile") is not None
    assert len(root.findall(f".//{ONE}Image")) == 1  # one-page PDF rendered


class FakeOneNote:
    """Mimics the parts of OneNote.Application that the tool calls."""

    def __init__(self):
        base = "https://d.docs.live.net/0123456789ABCDEF/Documents"
        self.notebooks = {"{nb-1}": ("Existing", f"{base}/Existing/", {}),
                          "{nb-2}": ("Family", f"{base}/Family/", {})}
        self.pages = {}
        self.page_sections = {}
        self.fail_next_update = False

    def GetHierarchy(self, start, scope):
        if scope == onenote_com.HS_NOTEBOOKS:
            inner = "".join(f'<one:Notebook name="{n}" nickname="{n}" ID="{i}" path="{p}"/>'
                            for i, (n, p, _) in self.notebooks.items())
            return f'<one:Notebooks xmlns:one="{NS}">{inner}</one:Notebooks>'
        name, path, sections = self.notebooks[start]
        inner = "".join(f'<one:Section name="{n}" ID="{i}"/>' for i, n in sections.items())
        return f'<one:Notebook xmlns:one="{NS}" name="{name}" ID="{start}">{inner}</one:Notebook>'

    def OpenHierarchy(self, path, relative_to, cftIfNotExist=0):
        assert cftIfNotExist == onenote_com.CFT_SECTION, "the tool must never create notebooks"
        sections = self.notebooks[relative_to][2]
        sec_id = f"{{{relative_to}-sec-{len(sections) + 1}}}"  # unique across notebooks
        sections[sec_id] = path[:-4]
        return sec_id

    def CreateNewPage(self, section_id):
        page_id = f"{{page-{len(self.pages) + 1}}}"
        self.pages[page_id] = None
        self.page_sections[page_id] = section_id
        return page_id

    def UpdatePageContent(self, xml, date_expected_last_modified):
        if self.fail_next_update:
            self.fail_next_update = False
            raise RuntimeError("0x80042001 invalid XML")
        page_id = ET.fromstring(xml).get("ID")
        self.pages[page_id] = xml

    def DeleteHierarchy(self, object_id, date_expected_last_modified):
        del self.pages[object_id]

    def GetHyperlinkToObject(self, object_id, content_id):
        return f"onenote:#page&id={object_id}"


def test_com_hierarchy_helpers():
    app = FakeOneNote()
    com = OneNoteCom(app)
    assert com.find_notebook(" existing ")["id"] == "{nb-1}"
    assert com.find_notebook("New One") is None
    sec = com.ensure_section("{nb-2}", "Bills")
    assert com.ensure_section("{nb-2}", "bills")["id"] == sec["id"]

    app.fail_next_update = True
    with pytest.raises(RuntimeError):
        com.create_page(sec["id"], lambda pid: f'<one:Page xmlns:one="{NS}" ID="{pid}"/>')
    assert app.pages == {}  # the empty page was cleaned up


def test_cli_upload_via_com(sample_enex, tmp_path, monkeypatch, capsys):
    from test_pipeline import FakeAnalyzer

    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: FakeAnalyzer())
    base = ["--state-dir", str(tmp_path / "state")]
    cli.main(base + ["analyze", str(sample_enex)])

    cli.main(base + ["upload", str(sample_enex), "--target", "com", "--dry-run"])
    previews = sorted((state_dir_for(tmp_path / "state", sample_enex) / "preview").glob("*.xml"))
    assert len(previews) == 2

    cli.main(base + ["upload", str(sample_enex), "--target", "com", "--notebook", "Existing"])
    out = capsys.readouterr().out
    assert "section 'Household Bills'" in out
    assert "Uploaded 2, skipped 0, failed 0, not analyzed 1" in out
    titles = [ET.fromstring(x).find(f"{ONE}Title/{ONE}OE/{ONE}T").text
              for x in app.pages.values()]
    assert titles == ["AI Scan_0012.pdf", "AI Recipe &amp; notes"]  # T text is HTML

    # Re-running uploads nothing new.
    cli.main(base + ["upload", str(sample_enex), "--target", "com", "--notebook", "Existing"])
    assert "Uploaded 0" in capsys.readouterr().out and len(app.pages) == 2


@pytest.mark.parametrize("section, expected", [
    (None, {"Household Bills": 2, "Recipes": 2}),  # default: one section per export
    ("Everything", {"Everything": 4}),
])
def test_merge_exports_into_one_notebook(sample_enex, second_enex, tmp_path, monkeypatch,
                                         capsys, section, expected):
    from test_pipeline import FakeAnalyzer

    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: FakeAnalyzer())
    base = ["--state-dir", str(tmp_path / "state")]
    target = ["--target", "com", "--notebook", "Family"] + (["--section", section] if section else [])
    for enex in (sample_enex, second_enex):
        cli.main(base + ["analyze", str(enex)])
        cli.main(base + ["upload", str(enex)] + target)

    sections = app.notebooks["{nb-2}"][2]
    counts = {name: sum(1 for s in app.page_sections.values() if s == sid)
              for sid, name in sections.items()}
    assert counts == expected
    assert len(app.notebooks) == 2


@pytest.mark.parametrize("command", ["upload", "migrate"])
def test_missing_notebook_is_not_created(sample_enex, tmp_path, monkeypatch, command):
    from test_pipeline import FakeAnalyzer

    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    analyzer = FakeAnalyzer()
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: analyzer)
    base = ["--state-dir", str(tmp_path / "state")]
    with pytest.raises(SystemExit) as exc:
        cli.main(base + [command, str(sample_enex), "--notebook", "Finances"])
    msg = str(exc.value)
    assert "no open notebook named:\n  Finances\n" in msg and "File > New > OneDrive" in msg
    assert "  Existing\n  Family" in msg  # lists what is open
    assert len(app.notebooks) == 2 and not app.pages
    assert analyzer.seen_vocab is None  # migrate stops before analyzing anything


def test_dispatch_explains_32_bit_only_type_library(monkeypatch):
    """Click-to-Run Office can leave OneNote's type library without a win64 entry."""
    import sys
    import types

    class ComError(Exception):
        def __init__(self, hresult, msg):
            super().__init__(hresult, msg)
            self.hresult = hresult

    class App:
        def GetHierarchy(self, start, scope):
            raise ComError(onenote_com.TYPE_E_LIBNOTREGISTERED, "Library not registered.")

    gencache = types.SimpleNamespace(
        EnsureModule=lambda *a: types.SimpleNamespace(Application=App))
    client = types.ModuleType("win32com.client")
    client.gencache = gencache
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setitem(sys.modules, "pywintypes", types.SimpleNamespace(com_error=ComError))

    with pytest.raises(SystemExit) as exc:
        onenote_com._dispatch()
    msg = str(exc.value)
    assert "Library not registered" in msg
    for entry in (r"\1.1\0\win64", r"\1.0\0\win64", r"\1.0\0\win32"):
        assert f"{entry}' -Value '" in msg
    assert "Remove-Item -Recurse" in msg


def test_com_is_the_default_target():
    parser = cli.build_parser()
    for argv in (["upload", "x.enex"], ["migrate", "x.enex"], ["notebooks"]):
        assert parser.parse_args(argv).target == "com"


def test_migrate_checks_onenote_before_analyzing(sample_enex, tmp_path, monkeypatch):
    def no_onenote():
        raise SystemExit("Could not start OneNote through COM")

    def must_not_analyze(*a, **k):
        raise AssertionError("analysis started before OneNote was checked")

    monkeypatch.setattr(onenote_com, "_dispatch", no_onenote)
    monkeypatch.setattr(cli, "make_analyzer", must_not_analyze)
    with pytest.raises(SystemExit, match="Could not start OneNote"):
        cli.main(["--state-dir", str(tmp_path / "state"), "migrate", str(sample_enex)])


def test_same_file_name_in_two_notebook_folders(tmp_path, monkeypatch):
    """Finances/Receipts.enex and Family/Receipts.enex keep separate progress."""
    import csv

    from conftest import note_xml, write_enex
    from test_pipeline import FakeAnalyzer

    shared = note_xml("Scan_0001.pdf", "<p>Same note exported from both notebooks</p>")
    exports = {}
    for folder, own in (("Existing", "Hardware store receipt"), ("Family", "Zoo tickets")):
        (tmp_path / folder).mkdir()
        exports[folder] = write_enex(tmp_path / folder / "Receipts.enex",
                                     [shared, note_xml(own, f"<p>{own}</p>")])

    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    analyzer = FakeAnalyzer()
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: analyzer)
    state = tmp_path / "state"
    for folder, enex in exports.items():
        cli.main(["--state-dir", str(state), "analyze", str(enex)])
        cli.main(["--state-dir", str(state), "upload", str(enex), "--notebook", folder])

    for folder, own in (("Existing", "Hardware store receipt"), ("Family", "Zoo tickets")):
        review = state / folder / "Receipts" / "review.csv"
        titles = [r["original_title"] for r in csv.DictReader(review.open(encoding="utf-8-sig"))]
        assert titles == ["Scan_0001.pdf", own]
    assert analyzer.seen_vocab == ["bills"]  # tags from the first folder reach the second
    # The shared note is uploaded into both notebooks, not skipped the second time.
    for nb_id in ("{nb-1}", "{nb-2}"):
        (sec_id,) = app.notebooks[nb_id][2]
        assert sum(1 for s in app.page_sections.values() if s == sec_id) == 2


def make_export_folder(root):
    """Exports/<Notebook>/<Section>.enex with notebooks Existing and Family."""
    from conftest import note_xml, write_enex

    for rel, titles in {"Existing/Receipts.enex": ["Hardware store", "Pharmacy"],
                        "Existing/Taxes.enex": ["W-2 form"],
                        "Family/Recipes.enex": ["Pancakes", "Apple pie", "Chili"]}.items():
        path = root / "Exports" / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        write_enex(path, [note_xml(t, f"<p>{t}</p>") for t in titles])
    return root / "Exports"


def pages_by_section(app):
    """{(notebook, section): [page titles]} from the fake OneNote."""
    out = {}
    for nb_name, _, sections in app.notebooks.values():
        for sec_id, sec_name in sections.items():
            out[(nb_name, sec_name)] = [
                ET.fromstring(app.pages[p]).find(f"{ONE}Title/{ONE}OE/{ONE}T").text
                for p, s in app.page_sections.items() if s == sec_id]
    return out


@pytest.mark.parametrize("two_steps", [True, False])
def test_folder_upload_maps_folders_to_notebooks(tmp_path, monkeypatch, capsys, two_steps):
    from test_pipeline import FakeAnalyzer

    exports = make_export_folder(tmp_path)
    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: FakeAnalyzer())
    base = ["--state-dir", str(tmp_path / "state")]
    if two_steps:
        cli.main(base + ["analyze", str(exports)])
        cli.main(base + ["upload", str(exports)])
    else:
        cli.main(base + ["migrate", str(exports)])

    assert pages_by_section(app) == {
        ("Existing", "Receipts"): ["AI Hardware store", "AI Pharmacy"],
        ("Existing", "Taxes"): ["AI W-2 form"],
        ("Family", "Recipes"): ["AI Pancakes", "AI Apple pie", "AI Chili"],
    }
    out = capsys.readouterr().out
    assert "3 export(s) in 2 notebook(s):\n  Existing: Receipts, Taxes\n  Family: Recipes" in out
    assert "=== Family / Recipes (3 of 3) ===" in out
    assert "All 3 exports: " in out and "uploaded 6" in out
    for review in ("Existing/Receipts", "Existing/Taxes", "Family/Recipes"):
        assert (tmp_path / "state" / review / "review.csv").exists()

    # Re-running the folder uploads nothing new.
    cli.main(base + ["upload", str(exports)])
    assert "uploaded 0" in capsys.readouterr().out and len(app.pages) == 6


@pytest.mark.parametrize("command", ["upload", "migrate"])
def test_folder_upload_checks_every_notebook_first(tmp_path, monkeypatch, command):
    from conftest import note_xml, write_enex
    from test_pipeline import FakeAnalyzer

    exports = make_export_folder(tmp_path)
    for name in ("Travel", "Finances"):
        (exports / name).mkdir()
        write_enex(exports / name / "Misc.enex", [note_xml("t", "<p>x</p>")])
    app = FakeOneNote()
    monkeypatch.setattr(onenote_com, "_dispatch", lambda: app)
    analyzer = FakeAnalyzer()
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: analyzer)
    with pytest.raises(SystemExit) as exc:
        cli.main(["--state-dir", str(tmp_path / "state"), command, str(exports)])
    assert "no open notebook named:\n  Finances\n  Travel\n" in str(exc.value)
    assert not app.pages and analyzer.seen_vocab is None  # nothing uploaded or analyzed


def test_folder_rejects_notebook_and_section_options(tmp_path):
    exports = make_export_folder(tmp_path)
    for opt in (["--notebook", "X"], ["--section", "Y"]):
        with pytest.raises(SystemExit, match="apply to a single .enex file"):
            cli.main(["--state-dir", str(tmp_path / "state"), "upload", str(exports), *opt])


def test_folder_dry_run_needs_no_onenote(tmp_path, monkeypatch, capsys):
    from test_pipeline import FakeAnalyzer

    def no_onenote():
        raise AssertionError("dry run must not connect to OneNote")

    exports = make_export_folder(tmp_path)
    monkeypatch.setattr(onenote_com, "_dispatch", no_onenote)
    monkeypatch.setattr(cli, "make_analyzer", lambda *a, **k: FakeAnalyzer())
    base = ["--state-dir", str(tmp_path / "state")]
    cli.main(base + ["analyze", str(exports / "Family"), "--limit", "2"])  # one notebook folder
    cli.main(base + ["upload", str(exports / "Family"), "--dry-run"])
    previews = list((tmp_path / "state" / "Family" / "Recipes" / "preview").glob("*.xml"))
    assert len(previews) == 2  # only the notes --limit let through were analyzed
    out = capsys.readouterr().out
    assert "previewed 2" in out and "not analyzed 1" in out
