"""Tests for style extended: get_style, copy_style, apply_style_to_range."""
from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError


def _make_doc(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


class TestGetStyle:
    def test_get_existing_style_by_name(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.get_style("Normal")
        assert "style_id" in result
        assert "name" in result
        assert "type" in result
        assert "base_style" in result
        assert "next_style" in result
        assert result["name"] == "Normal"

    def test_get_style_by_id(self, tmp_path):
        doc = _make_doc(tmp_path)
        # Create a style with a known id
        doc.create_style("FindMeById", "paragraph")
        result = doc.get_style("FindMeById")
        assert result["style_id"] == "FindMeById"
        assert result["name"] == "FindMeById"

    def test_get_style_not_found_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="not found"):
            doc.get_style("NoSuchStyleXYZ")

    def test_get_style_case_insensitive(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.get_style("normal")
        assert result["name"] == "Normal"

    def test_get_style_base_style_empty_when_absent(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("StandaloneStyle", "paragraph")
        result = doc.get_style("StandaloneStyle")
        assert result["base_style"] == ""

    def test_get_style_next_style_empty_when_absent(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("NoNextStyle", "paragraph")
        result = doc.get_style("NoNextStyle")
        assert result["next_style"] == ""

    def test_get_style_base_style_populated(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("DerivedStyle", "paragraph", based_on="Normal")
        result = doc.get_style("DerivedStyle")
        assert result["base_style"] == "Normal"


class TestCopyStyle:
    def test_copy_style_creates_new_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.copy_style("Normal", "MyNormal")
        styles = doc.get_styles()
        ids = [s["id"] for s in styles]
        assert "MyNormal" in ids

    def test_copy_style_returns_correct_keys(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.copy_style("Normal", "CopyKeys")
        assert "style_id" in result
        assert "name" in result
        assert "type" in result
        assert result["name"] == "CopyKeys"

    def test_copy_style_has_correct_type(self, tmp_path):
        doc = _make_doc(tmp_path)
        # Normal is a paragraph style
        original = doc.get_style("Normal")
        result = doc.copy_style("Normal", "TypeCheck")
        assert result["type"] == original["type"]

    def test_copy_style_new_id_has_no_spaces(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.copy_style("Normal", "My Copy Style")
        assert result["style_id"] == "MyCopyStyle"
        # Verify it's findable by that id
        found = doc.get_style("MyCopyStyle")
        assert found["name"] == "My Copy Style"

    def test_copy_style_source_not_found_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="not found"):
            doc.copy_style("NonExistentSource", "NewStyle")

    def test_copy_style_target_already_exists_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.copy_style("Normal", "DupeTarget")
        with pytest.raises(ValueError, match="already exists"):
            doc.copy_style("Normal", "DupeTarget")

    def test_copy_style_by_style_id(self, tmp_path):
        doc = _make_doc(tmp_path)
        # Should be able to copy by styleId too
        result = doc.copy_style("Normal", "CopiedById")
        assert result["style_id"] == "CopiedById"


class TestApplyStyleToRange:
    def _get_first_para_id(self, doc: DocxDocument) -> str:
        from docx_mcp.document.base import W, W14
        tree = doc._require("word/document.xml")
        body = tree.find(f"{W}body")
        for p in body:
            pid = p.get(f"{W14}paraId")
            if pid:
                return pid
        raise RuntimeError("No paragraphs with paraId found")

    def test_apply_style_to_single_para(self, tmp_path):
        doc = _make_doc(tmp_path)
        para_id = self._get_first_para_id(doc)
        result = doc.apply_style_to_range([para_id], "Normal")
        assert result["applied"] == 1
        assert result["para_ids"] == [para_id]

    def test_apply_style_to_multiple_paras(self, tmp_path):
        doc = _make_doc(tmp_path)
        from docx_mcp.document.base import W, W14
        tree = doc._require("word/document.xml")
        body = tree.find(f"{W}body")
        para_ids = []
        for p in body:
            pid = p.get(f"{W14}paraId")
            if pid:
                para_ids.append(pid)
            if len(para_ids) >= 2:
                break
        if len(para_ids) < 2:
            # Insert a second paragraph
            doc.insert_paragraph(para_ids[0], "Second para", position="after")
            tree = doc._require("word/document.xml")
            body = tree.find(f"{W}body")
            para_ids = []
            for p in body:
                pid = p.get(f"{W14}paraId")
                if pid:
                    para_ids.append(pid)
                if len(para_ids) >= 2:
                    break
        result = doc.apply_style_to_range(para_ids, "Normal")
        assert result["applied"] == len(para_ids)

    def test_apply_style_returns_correct_keys(self, tmp_path):
        doc = _make_doc(tmp_path)
        para_id = self._get_first_para_id(doc)
        result = doc.apply_style_to_range([para_id], "Normal")
        assert "applied" in result
        assert "style_id" in result
        assert "para_ids" in result

    def test_apply_style_sets_pstyle_in_xml(self, tmp_path):
        doc = _make_doc(tmp_path)
        para_id = self._get_first_para_id(doc)
        # Create a custom style and apply it
        doc.create_style("TestApplyStyle", "paragraph")
        doc.apply_style_to_range([para_id], "TestApplyStyle")
        from docx_mcp.document.base import W, W14
        tree = doc._require("word/document.xml")
        body = tree.find(f"{W}body")
        for p in body.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                ppr = p.find(f"{W}pPr")
                assert ppr is not None
                pstyle = ppr.find(f"{W}pStyle")
                assert pstyle is not None
                assert pstyle.get(f"{W}val") == "TestApplyStyle"
                return
        pytest.fail("Paragraph not found")

    def test_apply_style_not_found_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        para_id = self._get_first_para_id(doc)
        with pytest.raises(ValueError, match="not found"):
            doc.apply_style_to_range([para_id], "NoSuchStyleXYZ")

    def test_apply_style_invalid_para_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError):
            doc.apply_style_to_range(["DEADBEEF"], "Normal")
