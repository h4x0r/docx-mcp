"""Tests for field operations: delete_field, get_field, insert_date_field, insert_page_number_field."""
from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14
from docx_mcp.document.errors import DocxMcpError

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


def _make_doc(tmp_path: Path) -> tuple[DocxDocument, str]:
    """Create a fresh document and return (doc, first_para_id)."""
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    tree = doc._tree("word/document.xml")
    para_id = None
    for p in tree.iter(f"{W}p"):
        pid = p.get(f"{W14}paraId")
        if pid is not None:
            para_id = pid
            break
    assert para_id is not None
    return doc, para_id


class TestInsertDateField:
    def test_insert_date_field_creates_field_runs(self, tmp_path: Path):
        """insert_date_field appends begin/instrText/separate/end runs to paragraph."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_date_field(para_id)

        tree = doc._tree("word/document.xml")
        para = next(
            p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id
        )
        fld_chars = list(para.iter(f"{W}fldChar"))
        types = [fc.get(f"{W}fldCharType") for fc in fld_chars]
        assert "begin" in types
        assert "separate" in types
        assert "end" in types
        assert types.index("begin") < types.index("separate") < types.index("end")

        instr_els = list(para.iter(f"{W}instrText"))
        assert len(instr_els) >= 1

    def test_insert_date_field_instrText_contains_DATE(self, tmp_path: Path):
        """instrText of inserted date field contains 'DATE'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_date_field(para_id)

        tree = doc._tree("word/document.xml")
        para = next(
            p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id
        )
        instr_els = list(para.iter(f"{W}instrText"))
        assert any("DATE" in (el.text or "") for el in instr_els)

    def test_insert_date_field_returns_dict(self, tmp_path: Path):
        """insert_date_field returns dict with para_id and field_type=DATE."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_date_field(para_id)
        assert result["para_id"] == para_id
        assert result["field_type"] == "DATE"
        assert "format" in result

    def test_insert_date_field_para_not_found(self, tmp_path: Path):
        """insert_date_field raises ValueError for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises((ValueError, DocxMcpError)):
            doc.insert_date_field("DEADBEEF")


class TestInsertPageNumberField:
    def test_insert_page_number_field_contains_PAGE(self, tmp_path: Path):
        """instrText of inserted page number field contains 'PAGE'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_page_number_field(para_id)

        tree = doc._tree("word/document.xml")
        para = next(
            p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id
        )
        instr_els = list(para.iter(f"{W}instrText"))
        assert any("PAGE" in (el.text or "") for el in instr_els)

    def test_insert_page_number_field_returns_dict(self, tmp_path: Path):
        """insert_page_number_field returns dict with para_id and field_type=PAGE."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_page_number_field(para_id)
        assert result["para_id"] == para_id
        assert result["field_type"] == "PAGE"


class TestGetField:
    def test_get_field_returns_details(self, tmp_path: Path):
        """get_field returns a dict matching list_fields entry for the same field."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_date_field(para_id)

        fields = doc.list_fields()
        assert len(fields) >= 1
        field_id = fields[0]["field_id"]

        result = doc.get_field(field_id)
        assert result["field_id"] == field_id
        assert "type" in result or "code" in result
        assert "instruction" in result or "code" in result

    def test_get_field_not_found_raises(self, tmp_path: Path):
        """get_field raises ValueError for unknown field_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises((ValueError, DocxMcpError)):
            doc.get_field("nonexistent_field_id")


class TestDeleteField:
    def test_delete_field_removes_runs(self, tmp_path: Path):
        """delete_field removes all field runs (begin through end) from paragraph."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_date_field(para_id)

        fields = doc.list_fields()
        assert len(fields) >= 1
        field_id = fields[0]["field_id"]

        result = doc.delete_field(field_id)
        assert result["deleted"] is True
        assert result["field_id"] == field_id

        # Paragraph should no longer have fldChar runs
        tree = doc._tree("word/document.xml")
        para = next(
            p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id
        )
        fld_chars = list(para.iter(f"{W}fldChar"))
        assert len(fld_chars) == 0

    def test_delete_field_not_found_raises(self, tmp_path: Path):
        """delete_field raises ValueError for unknown field_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises((ValueError, DocxMcpError)):
            doc.delete_field("nonexistent_field_id")
