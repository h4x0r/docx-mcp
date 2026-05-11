"""Tests for Style CRUD: create_style, update_style, delete_style."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument


def _make_doc(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


class TestCreateStyle:
    def test_create_paragraph_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.create_style("MyCustom", "paragraph")
        assert result["style_id"] == "MyCustom"
        assert result["name"] == "MyCustom"
        assert result["type"] == "paragraph"
        styles = doc.get_styles()
        ids = [s["id"] for s in styles]
        assert "MyCustom" in ids

    def test_create_character_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.create_style("BoldEmphasis", "character")
        assert result["type"] == "character"
        styles = doc.get_styles()
        match = next((s for s in styles if s["id"] == "BoldEmphasis"), None)
        assert match is not None
        assert match["type"] == "character"

    def test_create_style_with_based_on(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.create_style("Derived", "paragraph", based_on="Normal")
        assert result["style_id"] == "Derived"
        styles = doc.get_styles()
        match = next(s for s in styles if s["id"] == "Derived")
        assert match["base_style"] == "Normal"

    def test_create_style_with_next_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("SectionTitle", "paragraph", next_style="Normal")
        from docx_mcp.document.base import W
        tree = doc._tree("word/styles.xml")
        style_el = next(
            s for s in tree.findall(f"{W}style")
            if s.get(f"{W}styleId") == "SectionTitle"
        )
        next_el = style_el.find(f"{W}next")
        assert next_el is not None
        assert next_el.get(f"{W}val") == "Normal"

    def test_create_duplicate_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("UniqueStyle", "paragraph")
        with pytest.raises(ValueError, match="already exists"):
            doc.create_style("UniqueStyle", "paragraph")

    def test_create_table_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        result = doc.create_style("CustomTable", "table")
        assert result["type"] == "table"


class TestUpdateStyle:
    def test_update_based_on(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("StyleA", "paragraph")
        result = doc.update_style("StyleA", based_on="Normal")
        assert result["style_id"] == "StyleA"
        styles = doc.get_styles()
        match = next(s for s in styles if s["id"] == "StyleA")
        assert match["base_style"] == "Normal"

    def test_update_next_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("StyleB", "paragraph")
        doc.update_style("StyleB", next_style="Normal")
        from docx_mcp.document.base import W
        tree = doc._tree("word/styles.xml")
        style_el = next(
            s for s in tree.findall(f"{W}style")
            if s.get(f"{W}styleId") == "StyleB"
        )
        next_el = style_el.find(f"{W}next")
        assert next_el is not None
        assert next_el.get(f"{W}val") == "Normal"

    def test_update_replaces_existing_based_on(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("StyleC", "paragraph", based_on="Normal")
        doc.update_style("StyleC", based_on="Heading1")
        styles = doc.get_styles()
        match = next(s for s in styles if s["id"] == "StyleC")
        assert match["base_style"] == "Heading1"

    def test_update_nonexistent_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="not found"):
            doc.update_style("NoSuchStyle", based_on="Normal")

    def test_update_case_insensitive(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("CamelStyle", "paragraph")
        result = doc.update_style("camelstyle", based_on="Normal")
        assert result["style_id"] == "CamelStyle"


class TestDeleteStyle:
    def test_delete_style(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.create_style("ToDelete", "paragraph")
        result = doc.delete_style("ToDelete")
        assert result["deleted"] == "ToDelete"
        styles = doc.get_styles()
        ids = [s["id"] for s in styles]
        assert "ToDelete" not in ids

    def test_delete_nonexistent_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="not found"):
            doc.delete_style("GhostStyle")

    def test_delete_default_paragraph_raises(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="default"):
            doc.delete_style("Normal")
