"""Tests for P9.3 text boxes — insert_text_box."""

from __future__ import annotations

import zipfile

import pytest

from docx_mcp.document import DocxDocument, W, W14, WP
from docx_mcp.document.errors import DocxMcpError, ErrCode

# WPS namespace constant
_WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"


def _make_docx(tmp_path):
    """Create a minimal valid DOCX with one body paragraph."""
    path = tmp_path / "test.docx"
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p w14:paraId="AABBCC01" w14:textId="77777777">
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""
    ct_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""
    doc_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml.strip())
        zf.writestr("_rels/.rels", rels_xml.strip())
        zf.writestr("word/document.xml", doc_xml.strip())
        zf.writestr("word/_rels/document.xml.rels", doc_rels_xml.strip())
    return path


class TestInsertTextBox:
    def test_creates_drawing_in_new_paragraph(self, tmp_path):
        """insert_text_box inserts a new w:p with w:drawing after the reference para."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            doc.insert_text_box("AABBCC01", "Sample text")
            tree = doc._tree("word/document.xml")
            drawings = tree.findall(f".//{W}drawing")
            assert len(drawings) == 1
        finally:
            doc.close()

    def test_textbox_contains_text(self, tmp_path):
        """The w:txbxContent inside the drawing contains the supplied text."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            doc.insert_text_box("AABBCC01", "My box text")
            tree = doc._tree("word/document.xml")
            # Find txbxContent
            txbx_contents = tree.findall(f".//{W}txbxContent")
            assert len(txbx_contents) == 1
            # Find w:t inside
            t_els = txbx_contents[0].findall(f".//{W}t")
            assert any("My box text" in (el.text or "") for el in t_els)
        finally:
            doc.close()

    def test_returns_correct_keys(self, tmp_path):
        """Result dict has para_id, text, width_cm, height_cm keys."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            result = doc.insert_text_box("AABBCC01", "Hello", width_cm=4.0, height_cm=3.0)
            assert set(result.keys()) >= {"para_id", "text", "width_cm", "height_cm"}
            assert result["text"] == "Hello"
            assert result["width_cm"] == 4.0
            assert result["height_cm"] == 3.0
        finally:
            doc.close()

    def test_returns_new_para_id(self, tmp_path):
        """result['para_id'] is the NEW paragraph's id, not the reference para_id."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            result = doc.insert_text_box("AABBCC01", "Test")
            assert result["para_id"] != "AABBCC01"
            assert result["para_id"]  # non-empty
        finally:
            doc.close()

    def test_width_and_height_in_emu(self, tmp_path):
        """wp:extent cx and cy reflect the cm-to-EMU conversion (1 cm = 360000 EMU)."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            doc.insert_text_box("AABBCC01", "x", width_cm=2.0, height_cm=1.0)
            tree = doc._tree("word/document.xml")
            extent = tree.find(f".//{WP}extent")
            assert extent is not None
            assert extent.get("cx") == "720000"   # 2 * 360000
            assert extent.get("cy") == "360000"   # 1 * 360000
        finally:
            doc.close()

    def test_para_not_found_raises(self, tmp_path):
        """DocxMcpError(PARA_NOT_FOUND) raised when reference para_id is missing."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            with pytest.raises(DocxMcpError) as exc_info:
                doc.insert_text_box("DOESNOTEXIST", "Boom")
            assert exc_info.value.code == ErrCode.PARA_NOT_FOUND
        finally:
            doc.close()

    def test_two_paragraphs_exist_after_insert(self, tmp_path):
        """Doc starts with 1 body para; after insert_text_box there are 2+ paras."""
        path = _make_docx(tmp_path)
        doc = DocxDocument(str(path))
        doc.open()
        try:
            tree = doc._tree("word/document.xml")
            body = tree.find(f"{W}body")
            # Count non-sectPr children that are w:p before insert
            before = [c for c in body if c.tag == f"{W}p"]
            assert len(before) == 1

            doc.insert_text_box("AABBCC01", "Second")
            tree2 = doc._tree("word/document.xml")
            body2 = tree2.find(f"{W}body")
            after = [c for c in body2 if c.tag == f"{W}p"]
            assert len(after) >= 2
        finally:
            doc.close()
