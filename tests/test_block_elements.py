"""Tests for P9.4: block elements — insert_blockquote, insert_code_block."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.document.errors import DocxMcpError, ErrCode


def _make_doc(tmp_path: Path) -> tuple[DocxDocument, str]:
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    tree = doc._tree("word/document.xml")
    para_id = next(
        p.get(f"{W14}paraId") for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") is not None
    )
    return doc, para_id


class TestInsertBlockquote:
    def test_inserts_paragraph_after_reference(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_blockquote(para_id, "A wise quote")
        tree = doc._tree("word/document.xml")
        paras = list(tree.iter(f"{W}p"))
        new_para = next((p for p in paras if p.get(f"{W14}paraId") == result["para_id"]), None)
        assert new_para is not None

    def test_has_left_indent_720(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_blockquote(para_id, "A wise quote")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        ppr = new_para.find(f"{W}pPr")
        assert ppr is not None
        ind = ppr.find(f"{W}ind")
        assert ind is not None
        assert ind.get(f"{W}left") == "720"

    def test_text_is_italic(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_blockquote(para_id, "A wise quote")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        r = new_para.find(f"{W}r")
        assert r is not None
        rpr = r.find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}i") is not None

    def test_contains_text(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_blockquote(para_id, "A wise quote")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        texts = [t.text or "" for t in new_para.iter(f"{W}t")]
        assert "A wise quote" in "".join(texts)

    def test_returns_new_para_id(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_blockquote(para_id, "A wise quote")
        assert result["para_id"] != para_id
        assert result["text"] == "A wise quote"

    def test_para_not_found_raises(self, tmp_path: Path):
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.insert_blockquote("DEADBEEF", "Quote")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND


class TestInsertCodeBlock:
    def test_inserts_paragraph_after_reference(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "print('hello')")
        tree = doc._tree("word/document.xml")
        new_para = next(
            (p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"]), None
        )  # noqa: E501
        assert new_para is not None

    def test_has_courier_new_font(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "print('hello')")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        r = new_para.find(f"{W}r")
        rpr = r.find(f"{W}rPr")
        fonts = rpr.find(f"{W}rFonts")
        assert fonts is not None
        assert fonts.get(f"{W}ascii") == "Courier New"

    def test_has_10pt_font_size(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "x = 1")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        r = new_para.find(f"{W}r")
        rpr = r.find(f"{W}rPr")
        sz = rpr.find(f"{W}sz")
        assert sz is not None
        assert sz.get(f"{W}val") == "20"

    def test_has_gray_shading(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "x = 1")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        ppr = new_para.find(f"{W}pPr")
        shd = ppr.find(f"{W}shd")
        assert shd is not None
        assert shd.get(f"{W}fill") == "F2F2F2"

    def test_contains_text(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "print('hello')")
        tree = doc._tree("word/document.xml")
        new_para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId") == result["para_id"])
        texts = [t.text or "" for t in new_para.iter(f"{W}t")]
        assert "print('hello')" in "".join(texts)

    def test_language_in_return(self, tmp_path: Path):
        doc, para_id = _make_doc(tmp_path)
        result = doc.insert_code_block(para_id, "x = 1", language="python")
        assert result["language"] == "python"
        assert result["text"] == "x = 1"

    def test_para_not_found_raises(self, tmp_path: Path):
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.insert_code_block("DEADBEEF", "x = 1")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND
