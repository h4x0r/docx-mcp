"""Tests for P9.1: advanced field codes — insert_if_field, insert_sequence_field, insert_merge_field."""  # noqa: E501
from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.document.errors import DocxMcpError, ErrCode


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


class TestInsertIfField:
    def test_creates_field_runs(self, tmp_path: Path):
        """insert_if_field appends begin/instrText/separate/end runs."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_if_field(para_id, "x > 0", "Yes", "No")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        fld_chars = list(para.iter(f"{W}fldChar"))
        types = [fc.get(f"{W}fldCharType") for fc in fld_chars]
        assert "begin" in types
        assert "separate" in types
        assert "end" in types
        assert types.index("begin") < types.index("separate") < types.index("end")

    def test_instrText_contains_IF(self, tmp_path: Path):
        """instrText of IF field contains 'IF'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_if_field(para_id, "x > 0", "Yes", "No")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        assert any("IF" in (el.text or "") for el in instr_els)

    def test_instrText_contains_condition_and_texts(self, tmp_path: Path):
        """instrText contains condition, true_text, and false_text."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_if_field(para_id, "x > 0", "Yes", "No")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        combined = " ".join(el.text or "" for el in instr_els)
        assert "x > 0" in combined
        assert "Yes" in combined
        assert "No" in combined

    def test_returns_correct_dict(self, tmp_path: Path):
        """insert_if_field returns dict with expected keys and values."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_if_field(para_id, "x > 0", "Yes", "No")
        assert result["para_id"] == para_id
        assert result["field_type"] == "IF"
        assert result["condition"] == "x > 0"
        assert result["true_text"] == "Yes"
        assert result["false_text"] == "No"

    def test_cached_value_is_true_text(self, tmp_path: Path):
        """Cached result run for IF field contains true_text."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_if_field(para_id, "x > 0", "Yes", "No")
        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        state = "before"
        cached_texts = []
        for child in para:
            fc = child.find(f"{{{W_NS}}}fldChar")
            if fc is not None:
                ftype = fc.get(f"{{{W_NS}}}fldCharType")
                if ftype == "separate":
                    state = "cached"
                    continue
                elif ftype == "end":
                    break
            if state == "cached":
                for t in child.iter(f"{{{W_NS}}}t"):
                    if t.text:
                        cached_texts.append(t.text)
        assert "Yes" in "".join(cached_texts)

    def test_para_not_found_raises(self, tmp_path: Path):
        """insert_if_field raises DocxMcpError(PARA_NOT_FOUND) for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.insert_if_field("DEADBEEF", "x > 0", "Yes", "No")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND


class TestInsertSequenceField:
    def test_instrText_contains_SEQ(self, tmp_path: Path):
        """instrText of SEQ field contains 'SEQ'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_sequence_field(para_id, "Figure")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        assert any("SEQ" in (el.text or "") for el in instr_els)

    def test_instrText_contains_seq_name(self, tmp_path: Path):
        """instrText contains the seq_name."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_sequence_field(para_id, "Figure")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        combined = " ".join(el.text or "" for el in instr_els)
        assert "Figure" in combined

    def test_reset_true_adds_reset_switch(self, tmp_path: Path):
        """reset=True adds \\r 1 switch to instrText."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_sequence_field(para_id, "Figure", reset=True)

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        combined = " ".join(el.text or "" for el in instr_els)
        assert r"\r" in combined or "\\r" in combined

    def test_reset_false_no_reset_switch(self, tmp_path: Path):
        """reset=False (default) does NOT add \\r switch."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_sequence_field(para_id, "Table")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        combined = " ".join(el.text or "" for el in instr_els)
        assert r"\r" not in combined

    def test_returns_correct_dict(self, tmp_path: Path):
        """insert_sequence_field returns dict with expected keys."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_sequence_field(para_id, "Figure")
        assert result["para_id"] == para_id
        assert result["field_type"] == "SEQ"
        assert result["seq_name"] == "Figure"
        assert result["reset"] is False

    def test_returns_correct_dict_reset(self, tmp_path: Path):
        """insert_sequence_field with reset=True returns reset=True in dict."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_sequence_field(para_id, "Figure", reset=True)
        assert result["reset"] is True

    def test_cached_value_is_1(self, tmp_path: Path):
        """Cached result run for SEQ field contains '1'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_sequence_field(para_id, "Figure")
        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        state = "before"
        cached_texts = []
        for child in para:
            fc = child.find(f"{{{W_NS}}}fldChar")
            if fc is not None:
                ftype = fc.get(f"{{{W_NS}}}fldCharType")
                if ftype == "separate":
                    state = "cached"
                    continue
                elif ftype == "end":
                    break
            if state == "cached":
                for t in child.iter(f"{{{W_NS}}}t"):
                    if t.text:
                        cached_texts.append(t.text)
        assert "1" in "".join(cached_texts)

    def test_para_not_found_raises(self, tmp_path: Path):
        """insert_sequence_field raises DocxMcpError(PARA_NOT_FOUND) for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.insert_sequence_field("DEADBEEF", "Figure")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND


class TestInsertMergeField:
    def test_instrText_contains_MERGEFIELD(self, tmp_path: Path):
        """instrText of MERGEFIELD contains 'MERGEFIELD'."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_merge_field(para_id, "FirstName")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        assert any("MERGEFIELD" in (el.text or "") for el in instr_els)

    def test_instrText_contains_field_name(self, tmp_path: Path):
        """instrText contains the field_name."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_merge_field(para_id, "FirstName")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        instr_els = list(para.iter(f"{W}instrText"))
        combined = " ".join(el.text or "" for el in instr_els)
        assert "FirstName" in combined

    def test_cached_value_is_guillemet(self, tmp_path: Path):
        """Cached value run contains «field_name»."""
        doc, para_id = _make_doc(tmp_path)
        doc.insert_merge_field(para_id, "FirstName")

        tree = doc._tree("word/document.xml")
        para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)
        # The cached value is the w:t after the separate fldChar
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        state = "before"
        cached_texts = []
        for child in para:
            fc = child.find(f"{{{W_NS}}}fldChar")
            if fc is not None:
                ftype = fc.get(f"{{{W_NS}}}fldCharType")
                if ftype == "separate":
                    state = "cached"
                    continue
                elif ftype == "end":
                    break
            if state == "cached":
                for t in child.iter(f"{{{W_NS}}}t"):
                    if t.text:
                        cached_texts.append(t.text)
        cached = "".join(cached_texts)
        assert "«FirstName»" in cached

    def test_returns_correct_dict(self, tmp_path: Path):
        """insert_merge_field returns dict with expected keys."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_merge_field(para_id, "FirstName")
        assert result["para_id"] == para_id
        assert result["field_type"] == "MERGEFIELD"
        assert result["field_name"] == "FirstName"

    def test_para_not_found_raises(self, tmp_path: Path):
        """insert_merge_field raises DocxMcpError(PARA_NOT_FOUND) for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.insert_merge_field("DEADBEEF", "FirstName")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND
