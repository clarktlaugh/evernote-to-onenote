import base64
import hashlib
import io
import struct
import zlib

import pytest


def make_png() -> bytes:
    def chunk(kind, data):
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))
    raw = b"\x00\xff\x00\x00"  # one red pixel
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def make_text_pdf(text: str) -> bytes:
    """A minimal one-page PDF whose text pypdf can extract."""
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = io.BytesIO()
    out.write(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(out.tell())
        out.write(b"%d 0 obj\n" % i + body + b"\nendobj\n")
    xref = out.tell()
    out.write(b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1))
    for off in offsets:
        out.write(b"%010d 00000 n \n" % off)
    out.write(b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
              % (len(objs) + 1, xref))
    return out.getvalue()


def make_blank_pdf() -> bytes:
    from pypdf import PdfWriter

    w = PdfWriter()
    w.add_blank_page(612, 792)
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue()


def resource_xml(data: bytes, mime: str, filename: str | None) -> str:
    fn = f"<resource-attributes><file-name>{filename}</file-name></resource-attributes>" if filename else ""
    return (f"<resource><data encoding=\"base64\">{base64.b64encode(data).decode()}</data>"
            f"<mime>{mime}</mime>{fn}</resource>")


def note_xml(title, body, resources=(), tags=(), created="20210314T101500Z"):
    tag_xml = "".join(f"<tag>{t}</tag>" for t in tags)
    res = "".join(resources)
    return (f"<note><title>{title}</title><content><![CDATA[<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            f"<!DOCTYPE en-note SYSTEM \"http://xml.evernote.com/pub/enml2.dtd\">"
            f"<en-note>{body}</en-note>]]></content><created>{created}</created>"
            f"<updated>{created}</updated>{tag_xml}"
            f"<note-attributes><source-url>https://example.com/x</source-url></note-attributes>"
            f"{res}</note>")


def md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def write_enex(path, notes):
    path.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export4.dtd">\n'
        f'<en-export export-date="20240101T000000Z">{"".join(notes)}</en-export>',
        encoding="utf-8")
    return path


@pytest.fixture
def sample_enex(tmp_path):
    png = make_png()
    pdf = make_text_pdf("Comcast invoice March 2021 amount due 89.99")
    blank = make_blank_pdf()
    notes = [
        note_xml("Scan_0012.pdf",
                 f"<div>See attached</div><en-media hash=\"{md5(pdf)}\" type=\"application/pdf\"/>",
                 [resource_xml(pdf, "application/pdf", "Scan_0012.pdf")], tags=["Bills"]),
        note_xml("Untitled",
                 f"<div>Photo <b>here</b><script>alert(1)</script>:</div>"
                 f"<en-media hash=\"{md5(png)}\" type=\"image/png\" width=\"300\"/>"
                 f"<div><en-todo checked=\"true\"/>done <en-todo/>open</div>",
                 [resource_xml(png, "image/png", None),
                  resource_xml(blank, "application/pdf", "unreferenced.pdf")]),
        note_xml("Recipe &amp; notes", "<p>Plain text note about pancakes.</p>"),
    ]
    return write_enex(tmp_path / "Household Bills.enex", notes)


@pytest.fixture
def second_enex(tmp_path):
    """A second export, for merging several .enex files into one notebook."""
    return write_enex(tmp_path / "Recipes.enex", [
        note_xml("IMG_2231", "<p>Buttermilk pancakes: flour, buttermilk, eggs.</p>",
                 created="20190502T080000Z"),
        note_xml("Note", "<p>Grandma Ruth's apple pie.</p>", created="20180101T120000Z"),
    ])
