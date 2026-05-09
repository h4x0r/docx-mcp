"""Tests for FieldsMixin — field insertion (add_field, update_fields, list_fields)."""
from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14
from docx_mcp.document.errors import DocxMcpError, ErrCode

# ── Namespace strings for assertions ────────────────────────────────────────

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


# ── Helper ───────────────────────────────────────────────────────────────────


def _make_doc(tmp_path: Path) -> tuple[DocxDocument, str]:
    """Create a fresh document and return (doc, first_para_id)."""
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    # Get first body paragraph para_id
    tree = doc._tree("word/document.xml")
    paras = [p for p in tree.iter(f"{W}p")]
    para_id = None
    for p in paras:
        pid = p.get(f"{W14}paraId")
        if pid is not None:
            para_id = pid
            break
    assert para_id is not None, "No paragraph with w14:paraId found"
    return doc, para_id


# ── Tests ─────────────────────────────────────────────────────────────────────


class TestFields:
    def test_add_page_field(self, tmp_path: Path):
        """add_field inserts PAGE field and returns correct dict."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_field(para_id, "PAGE")
        assert result["code"] == "PAGE"
        assert result["cached_value"] == "0"
        assert result["para_id"] == para_id

    def test_add_seq_field(self, tmp_path: Path):
        """add_field inserts SEQ field."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_field(para_id, "SEQ Figure")
        assert result["code"] == "SEQ Figure"
        assert result["para_id"] == para_id

    def test_add_ref_field(self, tmp_path: Path):
        """add_field inserts REF MyBookmark field."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_field(para_id, "REF MyBookmark", cached_value="See p. 3")
        assert result["code"] == "REF MyBookmark"
        assert result["cached_value"] == "See p. 3"
        assert result["para_id"] == para_id

    def test_field_structure_begin_separate_end(self, tmp_path: Path):
        """Complex field has begin/instrText/separate/cached/end structure."""
        doc, para_id = _make_doc(tmp_path)
        doc.add_field(para_id, "PAGE", cached_value="1")

        tree = doc._tree("word/document.xml")

        # Find the paragraph
        para = None
        for p in tree.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                para = p
                break
        assert para is not None

        # Collect fldChar types in order
        fld_chars = list(para.iter(f"{W}fldChar"))
        types = [fc.get(f"{W}fldCharType") for fc in fld_chars]
        assert "begin" in types
        assert "separate" in types
        assert "end" in types
        assert types.index("begin") < types.index("separate") < types.index("end")

        # begin fldChar must have dirty="true"
        begin_fc = fld_chars[types.index("begin")]
        assert begin_fc.get(f"{W}dirty") == "true"

        # instrText exists
        instr_els = list(para.iter(f"{W}instrText"))
        assert len(instr_els) >= 1
        assert "PAGE" in instr_els[0].text

        # cached value run exists between separate and end
        t_els = [t for t in para.iter(f"{W}t")]
        cached_texts = [t.text for t in t_els]
        assert "1" in cached_texts

    def test_update_fields_sets_dirty(self, tmp_path: Path):
        """update_fields sets w:dirty=true on all begin fldChar elements."""
        doc, para_id = _make_doc(tmp_path)
        # Add two fields
        doc.add_field(para_id, "PAGE")
        doc.add_field(para_id, "NUMPAGES")

        # Manually clear dirty flag on one to test update_fields re-sets it
        tree = doc._tree("word/document.xml")
        for fc in tree.iter(f"{W}fldChar"):
            if fc.get(f"{W}fldCharType") == "begin":
                fc.attrib.pop(f"{W}dirty", None)

        result = doc.update_fields()
        assert isinstance(result["updated_count"], int)
        assert result["updated_count"] >= 2

        # Verify all begin fldChars now have dirty=true
        for fc in tree.iter(f"{W}fldChar"):
            if fc.get(f"{W}fldCharType") == "begin":
                assert fc.get(f"{W}dirty") == "true"

    def test_list_fields(self, tmp_path: Path):
        """list_fields returns all fields with code and cached_value."""
        doc, para_id = _make_doc(tmp_path)
        # Start with no fields
        assert doc.list_fields() == []

        doc.add_field(para_id, "PAGE", cached_value="5")
        doc.add_field(para_id, "NUMPAGES", cached_value="10")

        fields = doc.list_fields()
        assert len(fields) == 2

        codes = {f["code"] for f in fields}
        assert "PAGE" in codes
        assert "NUMPAGES" in codes

        page_field = next(f for f in fields if f["code"] == "PAGE")
        assert page_field["cached_value"] == "5"
        assert page_field["para_id"] == para_id

        numpages_field = next(f for f in fields if f["code"] == "NUMPAGES")
        assert numpages_field["cached_value"] == "10"

    def test_add_field_para_not_found(self, tmp_path: Path):
        """add_field raises DocxMcpError(PARA_NOT_FOUND) for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.add_field("DEADBEEF", "PAGE")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND
