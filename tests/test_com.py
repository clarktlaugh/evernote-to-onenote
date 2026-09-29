import xml.etree.ElementTree as ET

import pytest

from enex2onenote import cli, onenote_com
from enex2onenote.enex import Note, iter_notes
from enex2onenote.onenote_com import NS, OneNoteCom, build_page_xml, image_size

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
        self.notebooks = {"{nb-1}": ("Existing", "C:\\Notebooks\\Existing", {})}
        self.pages = {}
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
        if cftIfNotExist == onenote_com.CFT_NOTEBOOK:
            nb_id = f"{{nb-{len(self.notebooks) + 1}}}"
            self.notebooks[nb_id] = (path.rsplit("\\", 1)[-1].rsplit("/", 1)[-1], path, {})
            return nb_id
        sections = self.notebooks[relative_to][2]
        sec_id = f"{{sec-{len(sections) + 1}}}"
        sections[sec_id] = path[:-4]
        return sec_id

    def GetSpecialLocation(self, loc):
        return "C:\\Notebooks"

    def CreateNewPage(self, section_id):
        page_id = f"{{page-{len(self.pages) + 1}}}"
        self.pages[page_id] = None
        return page_id

    def UpdatePageContent(self, xml):
        if self.fail_next_update:
            self.fail_next_update = False
            raise RuntimeError("0x80042001 invalid XML")
        page_id = ET.fromstring(xml).get("ID")
        self.pages[page_id] = xml

    def DeleteHierarchy(self, object_id):
        del self.pages[object_id]

    def GetHyperlinkToObject(self, object_id, content_id):
        return f"onenote:#page&id={object_id}"


def test_com_hierarchy_helpers():
    app = FakeOneNote()
    com = OneNoteCom(app)
    nb, created = com.ensure_notebook("existing")
    assert nb["id"] == "{nb-1}" and not created
    nb2, created = com.ensure_notebook("New One")
    assert created and nb2["path"].endswith("New One")
    sec = com.ensure_section(nb2["id"], "Bills")
    assert com.ensure_section(nb2["id"], "bills")["id"] == sec["id"]

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
    previews = sorted((tmp_path / "state" / "Household Bills" / "preview").glob("*.xml"))
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
