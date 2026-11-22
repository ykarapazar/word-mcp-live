"""Unit tests for extended macOS helpers in word_document_server.core.word_mac_ext."""

from __future__ import annotations

import os
import zipfile
from lxml import etree as ET
import pytest

from word_document_server.core import word_mac_ext as ext


def test_hex_to_as_rgb():
    assert ext._hex_to_as_rgb("#000000") == "{0, 0, 0}"
    assert ext._hex_to_as_rgb("FFFFFF") == "{65535, 65535, 65535}"
    assert ext._hex_to_as_rgb("#C0C0C0") == "{49344, 49344, 49344}"
    with pytest.raises(ValueError):
        ext._hex_to_as_rgb("#FFF")


def test_latex_to_unicodemath():
    sym = {r"\alpha": "\u03b1", r"\beta": "\u03b2", r"\pi": "\u03c0"}
    assert ext._latex_to_unicodemath(r"\frac{\alpha}{\beta}", sym) == "(\u03b1)/(\u03b2)"
    assert ext._latex_to_unicodemath(r"\sqrt{x^2+1}", sym) == "\u221a(x^2+1)"
    assert ext._latex_to_unicodemath(r"\sqrt[3]{27}", sym) == "\u221a(3&27)"
    assert ext._latex_to_unicodemath(r"e^{i\pi} + x_{10}", sym) == "e^(i\u03c0) + x_(10)"


def test_undo_history_tracker():
    ext._MAC_UNDO_HISTORY.clear()
    ext.record_mac_undo("Skripsi.docx", "Format table 1")
    ext.record_mac_undo("Skripsi.docx", "Insert equation")
    key = ext._norm_doc_key("Skripsi.docx")
    assert [e["action"] for e in ext._MAC_UNDO_HISTORY[key]] == [
        "Insert equation",
        "Format table 1",
    ]
    ext.pop_mac_undo("skripsi.docx", 1)
    assert [e["action"] for e in ext._MAC_UNDO_HISTORY[key]] == ["Format table 1"]


def test_rewrite_docx_zip_preserves_inode(tmp_path):
    docx = tmp_path / "sample.docx"
    with zipfile.ZipFile(docx, "w") as zf:
        zf.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>')
        zf.writestr("word/document.xml", "<w:document/>")

    orig_inode = os.stat(docx).st_ino
    ext._rewrite_docx_zip(str(docx), {"word/document.xml": b"<w:document><w:body/></w:document>"})
    assert os.stat(docx).st_ino == orig_inode

    with zipfile.ZipFile(docx, "r") as zf:
        assert zf.read("word/document.xml") == b"<w:document><w:body/></w:document>"


def test_ensure_comments_extended_wiring_preserves_namespaces(tmp_path):
    docx = tmp_path / "wired.docx"
    ct_orig = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "</Types>"
    )
    rels_orig = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://example.com/a" Target="styles.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(docx, "w") as zf:
        zf.writestr("[Content_Types].xml", ct_orig)
        zf.writestr("word/_rels/document.xml.rels", rels_orig)

    updates = {}
    with zipfile.ZipFile(docx, "r") as zf:
        ext._ensure_comments_extended_wiring(zf, updates)

    ct_new = updates["[Content_Types].xml"].decode("utf-8")
    rels_new = updates["word/_rels/document.xml.rels"].decode("utf-8")
    assert 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"' in ct_new
    assert "/word/commentsExtended.xml" in ct_new
    assert "ns0:" not in ct_new
    assert 'Id="rId2"' in rels_new
    assert 'Target="commentsExtended.xml"' in rels_new
    assert "ns0:" not in rels_new


def test_lxml_preserves_mc_ignorable_prefixes():
    raw = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        b'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        b'mc:Ignorable="w14 w15">'
        b'<w15:commentEx w15:paraId="12345678" w15:done="0"/>'
        b"</w15:commentsEx>"
    )
    root = ET.fromstring(raw)
    root[0].set(f"{{{ext._W15_NS}}}done", "1")
    out = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    assert 'xmlns:w14=' in out
    assert 'xmlns:mc=' in out
    assert 'w15:done="1"' in out


def test_stage_image_for_word_sandbox(tmp_path, monkeypatch):
    fake_gc = tmp_path / "GroupContainer"
    fake_gc.mkdir()
    monkeypatch.setattr(ext, "_OFFICE_GROUP_CONTAINER", str(fake_gc))

    src_img = tmp_path / "photo.png"
    src_img.write_bytes(b"\x89PNG\r\n\x1a\nFAKE")

    staged, cleanup = ext._stage_image_for_word_sandbox(str(src_img))
    assert cleanup is True
    assert staged.startswith(str(fake_gc / "MCPStaging"))
    assert os.path.exists(staged)
    assert open(staged, "rb").read() == b"\x89PNG\r\n\x1a\nFAKE"
    os.unlink(staged)
