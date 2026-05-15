"""Tests for find_replace_formatted: find/replace with character formatting."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002" w14:textId="77777777">
      <w:r><w:t>foo bar foo</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000003" w14:textId="77777777">
      <w:r><w:t>No match here</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000004" w14:textId="77777777">
      <w:r><w:t>alpha</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000005" w14:textId="77777777">
      <w:r><w:t>beta</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

_TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>"""

_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"""


def _build_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _TOP_RELS.strip())
        zf.writestr("word/document.xml", _DOCUMENT_XML.strip())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS.strip())
        zf.writestr("word/styles.xml", _STYLES_XML.strip())


def _open_doc(path: Path) -> DocxDocument:
    server._doc = None
    server.open_document(str(path))
    return server._doc


def _get_runs(doc: DocxDocument, para_id: str) -> list:
    root = doc._tree("word/document.xml")
    para = doc._find_para(root, para_id)
    return para.findall(f"{W}r")


class TestFindReplaceFormatted:
    def test_basic_replace_returns_count_one(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.find_replace_formatted("Hello", "Hi")
        assert result == {"find": "Hello", "replace": "Hi", "count": 1}

    def test_replace_applies_bold_to_replacement_run(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("world", "earth", bold=True)
        runs = _get_runs(doc, "00000001")
        fmt_run = next(
            r for r in runs if r.find(f"{W}t") is not None and r.find(f"{W}t").text == "earth"
        )  # noqa: E501
        rpr = fmt_run.find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}b") is not None

    def test_replace_applies_italic_to_replacement_run(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("Hello", "Hi", italic=True)
        runs = _get_runs(doc, "00000001")
        fmt_run = next(
            r for r in runs if r.find(f"{W}t") is not None and r.find(f"{W}t").text == "Hi"
        )  # noqa: E501
        rpr = fmt_run.find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}i") is not None

    def test_replace_applies_color_to_replacement_run(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("Hello", "Hi", color="FF0000")
        runs = _get_runs(doc, "00000001")
        fmt_run = next(
            r for r in runs if r.find(f"{W}t") is not None and r.find(f"{W}t").text == "Hi"
        )  # noqa: E501
        rpr = fmt_run.find(f"{W}rPr")
        assert rpr is not None
        color_el = rpr.find(f"{W}color")
        assert color_el is not None
        assert color_el.get(f"{W}val") == "FF0000"

    def test_replace_with_size_pt_sets_sz_in_half_points(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("Hello", "Hi", size_pt=12.0)
        runs = _get_runs(doc, "00000001")
        fmt_run = next(
            r for r in runs if r.find(f"{W}t") is not None and r.find(f"{W}t").text == "Hi"
        )  # noqa: E501
        rpr = fmt_run.find(f"{W}rPr")
        assert rpr is not None
        sz = rpr.find(f"{W}sz")
        assert sz is not None
        assert sz.get(f"{W}val") == "24"
        sz_cs = rpr.find(f"{W}szCs")
        assert sz_cs is not None
        assert sz_cs.get(f"{W}val") == "24"

    def test_multiple_occurrences_returns_correct_count(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.find_replace_formatted("foo", "baz")
        assert result["count"] == 2

    def test_find_not_present_returns_count_zero(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.find_replace_formatted("xyzzy", "qwerty")
        assert result == {"find": "xyzzy", "replace": "qwerty", "count": 0}

    def test_empty_find_raises_value_error(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        with pytest.raises(ValueError, match="non-empty"):
            doc.find_replace_formatted("", "something")

    def test_bold_false_adds_b_val_zero(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("Hello", "Hi", bold=False)
        runs = _get_runs(doc, "00000001")
        fmt_run = next(
            r for r in runs if r.find(f"{W}t") is not None and r.find(f"{W}t").text == "Hi"
        )  # noqa: E501
        rpr = fmt_run.find(f"{W}rPr")
        assert rpr is not None
        b_el = rpr.find(f"{W}b")
        assert b_el is not None
        assert b_el.get(f"{W}val") == "0"

    def test_replace_splits_run_preserving_surrounding_text(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.find_replace_formatted("world", "earth", bold=True)
        runs = _get_runs(doc, "00000001")
        texts = [r.find(f"{W}t").text for r in runs if r.find(f"{W}t") is not None]
        assert "Hello " in texts
        assert "earth" in texts

    def test_result_structure_contains_find_and_replace_keys(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.find_replace_formatted("alpha", "ALPHA")
        assert result["find"] == "alpha"
        assert result["replace"] == "ALPHA"
        assert result["count"] == 1
