"""Tests for StatisticsMixin: get_word_count and get_statistics."""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""


def _build_doc(tmp_path: Path, body_xml: str) -> DocxDocument:
    doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
{body_xml}
  </w:body>
</w:document>"""
    path = tmp_path / "test.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _TOP_RELS.strip())
        zf.writestr("word/document.xml", doc_xml.strip())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS.strip())
    doc = DocxDocument(str(path))
    doc.open()
    return doc


class TestGetWordCount:

    def test_empty_doc_returns_zero(self, tmp_path: Path) -> None:
        doc = _build_doc(tmp_path, "")
        assert doc.get_word_count() == 0

    def test_single_paragraph_two_words(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:r><w:t>hello world</w:t></w:r>
    </w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_word_count() == 2

    def test_multiple_paragraphs_sum(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:r><w:t>one two three</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002">
      <w:r><w:t>four five</w:t></w:r>
    </w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_word_count() == 5

    def test_split_runs_counted_correctly(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:r><w:t xml:space="preserve">First </w:t></w:r>
      <w:r><w:t>second</w:t></w:r>
    </w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_word_count() == 2


class TestGetStatistics:

    def test_returns_all_six_keys(self, tmp_path: Path) -> None:
        doc = _build_doc(tmp_path, "")
        stats = doc.get_statistics()
        assert set(stats.keys()) == {
            "word_count",
            "character_count",
            "paragraph_count",
            "table_count",
            "image_count",
            "section_count",
        }

    def test_paragraph_count(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001"><w:r><w:t>Alpha</w:t></w:r></w:p>
    <w:p w14:paraId="00000002"><w:r><w:t>Beta</w:t></w:r></w:p>
    <w:p w14:paraId="00000003"><w:r><w:t>Gamma</w:t></w:r></w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_statistics()["paragraph_count"] == 3

    def test_table_count(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001"><w:r><w:t>Before</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p w14:paraId="00000002"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:tbl>
      <w:tr><w:tc><w:p w14:paraId="00000003"><w:r><w:t>Cell2</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>"""
        doc = _build_doc(tmp_path, body)
        stats = doc.get_statistics()
        assert stats["table_count"] == 2

    def test_character_count(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:r><w:t>hello world</w:t></w:r>
    </w:p>"""
        doc = _build_doc(tmp_path, body)
        stats = doc.get_statistics()
        assert stats["character_count"] == len("hello world")

    def test_image_count(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Picture 1"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_statistics()["image_count"] == 1

    def test_section_count_default_one(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001"><w:r><w:t>Only</w:t></w:r></w:p>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_statistics()["section_count"] == 1

    def test_section_count_with_explicit_sectpr(self, tmp_path: Path) -> None:
        body = """<w:p w14:paraId="00000001">
      <w:pPr><w:sectPr/></w:pPr>
      <w:r><w:t>Page 1</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002"><w:r><w:t>Page 2</w:t></w:r></w:p>
    <w:sectPr/>"""
        doc = _build_doc(tmp_path, body)
        assert doc.get_statistics()["section_count"] == 2
