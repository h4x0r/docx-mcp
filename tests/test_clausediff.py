"""Tests for ClauseDiffMixin.compare_contracts (Phase 7.2)."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document.errors import DocxMcpError, ErrCode

# ── Minimal DOCX builders ────────────────────────────────────────────────────

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


def _make_doc_xml(*sections: tuple[str, str, str]) -> str:
    """Build document XML with heading+body sections.
    Each section: (style_val, heading_text, body_text)
    style_val e.g. "Heading 1" or "Heading 2"
    """
    paras = []
    para_id = 1
    for style_val, heading_text, body_text in sections:
        paras.append(f"""    <w:p w14:paraId="{para_id:08X}">
      <w:pPr><w:pStyle w:val="{style_val}"/></w:pPr>
      <w:r><w:t>{heading_text}</w:t></w:r>
    </w:p>""")
        para_id += 1
        paras.append(f"""    <w:p w14:paraId="{para_id:08X}">
      <w:r><w:t>{body_text}</w:t></w:r>
    </w:p>""")
        para_id += 1
    body = "\n".join(paras)
    return f"""<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
{body}
  </w:body>
</w:document>"""


def _build_docx(path: Path, doc_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _TOP_RELS.strip())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS.strip())
        zf.writestr("word/document.xml", doc_xml.strip())


# ── Tests ────────────────────────────────────────────────────────────────────


class TestCompareContracts:
    @pytest.fixture(autouse=True)
    def _open_base(self, tmp_path: Path):
        base_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
            ("Heading 2", "Section Two", "Body text two."),
        )
        base = tmp_path / "base.docx"
        _build_docx(base, base_xml)
        server.open_document(str(base))
        self.tmp = tmp_path

    def test_matched_clause_unchanged(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
            ("Heading 2", "Section Two", "Body text two."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_compared"] >= 2
        assert result["clauses_changed"] == 0

    def test_changed_clause_detected(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "DIFFERENT body text."),
            ("Heading 2", "Section Two", "Body text two."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_changed"] >= 1

    def test_added_clause_detected(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
            ("Heading 2", "Section Two", "Body text two."),
            ("Heading 1", "Section Three", "New clause."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_changed"] >= 1

    def test_deleted_clause_detected(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_changed"] >= 1

    def test_renamed_clause_detected(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One Modified", "Body text one."),
            ("Heading 2", "Section Two", "Body text two."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_changed"] >= 1

    def test_missing_other_raises(self, tmp_path: Path):
        with pytest.raises(DocxMcpError) as exc_info:
            server._doc.compare_contracts(str(tmp_path / "nonexistent.docx"))
        assert exc_info.value.code == ErrCode.PART_NOT_FOUND

    def test_output_path_returned(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "result.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert "output_path" in result
        assert result["output_path"] == str(out)
        assert Path(result["output_path"]).exists()

    def test_reordered_clause_tracked(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section Two", "Body text two."),
            ("Heading 2", "Section One", "Body text one."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["reordered"] >= 1

    def test_default_output_path_used(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        result = server._doc.compare_contracts(str(other))
        assert "output_path" in result
        assert Path(result["output_path"]).exists()

    def test_server_tool_returns_json(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 1", "Section One", "Body text one."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        raw = server.compare_contracts(str(other), str(out))
        data = json.loads(raw)
        assert "clauses_compared" in data
        assert "clauses_changed" in data
        assert "reordered" in data
        assert "output_path" in data

    def test_heading2_style_detected(self, tmp_path: Path):
        other_xml = _make_doc_xml(
            ("Heading 2", "Section One", "Body text one."),
            ("Heading 2", "Section Two", "Body text two."),
        )
        other = tmp_path / "other.docx"
        _build_docx(other, other_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_compared"] >= 2

    def test_heading_style_case_insensitive(self, tmp_path: Path):
        doc_xml = """<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:pPr><w:pStyle w:val="heading1"/></w:pPr>
      <w:r><w:t>Section A</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002">
      <w:r><w:t>Body A.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
        other = tmp_path / "other.docx"
        _build_docx(other, doc_xml)
        out = tmp_path / "out.docx"
        result = server._doc.compare_contracts(str(other), str(out))
        assert result["clauses_compared"] >= 1
