"""Tests for P8.8: Accessibility — set_alt_text, get_alt_text, check_accessibility."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.base import W, W14, WP


# ── Helper: insert a wp:docPr-bearing drawing into a paragraph ───────────────

def _insert_drawing(doc: DocxDocument, para_id: str, image_name: str = "TestImage") -> None:
    """Insert a minimal wp:docPr-bearing drawing into a paragraph for testing."""
    tree = doc._tree("word/document.xml")
    para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
    run = etree.SubElement(para, f"{W}r")
    drawing = etree.SubElement(run, f"{W}drawing")
    inline = etree.SubElement(drawing, f"{WP}inline")
    doc_pr = etree.SubElement(inline, f"{WP}docPr")
    doc_pr.set("id", "99")
    doc_pr.set("name", image_name)
    doc._mark("word/document.xml")


# ── Tests: set_alt_text ──────────────────────────────────────────────────────

class TestSetAltText:
    def test_set_descr_returns_correct_dict(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        result = doc.set_alt_text(0, "A cat")
        assert result == {"image_index": 0, "alt_text": "A cat"}
        doc.close()

    def test_set_descr_and_title(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        doc.set_alt_text(0, "A cat", title="Cat photo")
        # Verify by reading back
        result = doc.get_alt_text(0)
        assert result["alt_text"] == "A cat"
        assert result["title"] == "Cat photo"
        doc.close()

    def test_set_descr_removes_title_when_empty(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        # First set with title
        doc.set_alt_text(0, "A cat", title="Cat photo")
        # Now set without title — should remove it
        doc.set_alt_text(0, "A cat")
        result = doc.get_alt_text(0)
        assert result["title"] == ""
        doc.close()

    def test_set_alt_text_marks_document_dirty(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        doc._modified.discard("word/document.xml")
        doc.set_alt_text(0, "A cat")
        assert "word/document.xml" in doc._modified
        doc.close()

    def test_set_alt_text_out_of_range_raises(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        with pytest.raises(IndexError, match="Image index 99 out of range"):
            doc.set_alt_text(99, "x")
        doc.close()


# ── Tests: get_alt_text ──────────────────────────────────────────────────────

class TestGetAltText:
    def test_get_returns_correct_dict(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        doc.set_alt_text(0, "A dog", title="Dog photo")
        result = doc.get_alt_text(0)
        assert result == {"image_index": 0, "alt_text": "A dog", "title": "Dog photo"}
        doc.close()

    def test_get_no_descr_returns_empty_string(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        # The fixture has an image with no descr attribute initially
        result = doc.get_alt_text(0)
        assert result["alt_text"] == ""
        assert result["title"] == ""
        doc.close()

    def test_get_alt_text_out_of_range_raises(self, test_docx):
        doc = DocxDocument(str(test_docx))
        doc.open()
        with pytest.raises(IndexError, match="Image index 99 out of range"):
            doc.get_alt_text(99)
        doc.close()


# ── Tests: check_accessibility ───────────────────────────────────────────────

class TestCheckAccessibility:
    def test_no_issues_when_all_ok(self, test_docx, tmp_path):
        """Doc with image that HAS alt text and table with header — no issues."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        # Set alt text on the existing image
        doc.set_alt_text(0, "A descriptive alt text")
        # The fixture table has no header row — add a header to suppress that issue.
        # We need a clean doc; build a new one in tmp_path that has a table with header.
        doc.close()

        # Build a minimal doc with no images and no tables
        import zipfile
        no_img_doc = tmp_path / "no_img.docx"
        _CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
        _TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""
        _DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""
        _DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
        with zipfile.ZipFile(no_img_doc, "w") as zf:
            zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
            zf.writestr("_rels/.rels", _TOP_RELS.strip())
            zf.writestr("word/document.xml", _DOCUMENT_XML.strip())
            zf.writestr("word/_rels/document.xml.rels", _DOC_RELS.strip())

        doc2 = DocxDocument(str(no_img_doc))
        doc2.open()
        result = doc2.check_accessibility()
        assert result == {"issue_count": 0, "issues": []}
        doc2.close()

    def test_reports_missing_alt_text(self, test_docx):
        """Image without alt text should appear in issues."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        # Don't set alt text — fixture image has no descr
        result = doc.check_accessibility()
        missing = [i for i in result["issues"] if i["type"] == "missing_alt_text"]
        assert len(missing) >= 1
        assert missing[0]["image_index"] == 0
        assert "no alt text" in missing[0]["description"]
        doc.close()

    def test_no_missing_alt_when_alt_set(self, test_docx):
        """Image with alt text should NOT be reported as missing."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        doc.set_alt_text(0, "A descriptive text")
        result = doc.check_accessibility()
        missing = [i for i in result["issues"] if i["type"] == "missing_alt_text"]
        assert len(missing) == 0
        doc.close()

    def test_reports_table_no_header(self, test_docx):
        """Table without tblHeader on first row should appear in issues."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        # Set alt text so images don't pollute; focus on table issue
        doc.set_alt_text(0, "Alt text")
        result = doc.check_accessibility()
        table_issues = [i for i in result["issues"] if i["type"] == "table_no_header"]
        assert len(table_issues) >= 1
        assert table_issues[0]["table_index"] == 0
        assert "no header row" in table_issues[0]["description"]
        doc.close()

    def test_issue_count_matches_issues_list(self, test_docx):
        """issue_count must equal len(issues)."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        result = doc.check_accessibility()
        assert result["issue_count"] == len(result["issues"])
        doc.close()

    def test_inserted_drawing_missing_alt_detected(self, test_docx):
        """Dynamically inserted drawing with no descr must be detected."""
        doc = DocxDocument(str(test_docx))
        doc.open()
        # Insert second drawing (no alt text)
        _insert_drawing(doc, "00000001", "Chart1")
        # Set alt text on the first image to isolate
        doc.set_alt_text(0, "First image alt")
        result = doc.check_accessibility()
        missing = [i for i in result["issues"] if i["type"] == "missing_alt_text"]
        # The second image (index 1) has no alt text
        indices = [i["image_index"] for i in missing]
        assert 1 in indices
        doc.close()
