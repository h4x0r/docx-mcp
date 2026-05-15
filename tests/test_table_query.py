"""Tests for table query: get_table, get_cell_text, copy_table."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument


def _make_doc_with_table(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    # get first para_id and add a 2x3 table
    tree = doc._tree("word/document.xml")
    from docx_mcp.document.base import W14, W

    para = next(p for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId"))
    para_id = para.get(f"{W14}paraId")
    doc.add_table(para_id, rows=2, cols=3)
    # populate some cells
    doc.modify_cell(0, 0, 0, "Alpha")
    doc.modify_cell(0, 0, 1, "Beta")
    doc.modify_cell(0, 1, 0, "Gamma")
    return doc


class TestGetTable:
    def test_get_table_returns_expected_shape(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.get_table(0)
        assert result["index"] == 0
        assert result["row_count"] == 2
        assert result["col_count"] == 3
        assert "cells" in result

    def test_get_table_cells_content(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.get_table(0)
        assert result["cells"][0][0] == "Alpha"
        assert result["cells"][0][1] == "Beta"
        assert result["cells"][1][0] == "Gamma"

    def test_get_table_out_of_range_raises(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.get_table(99)


class TestGetCellText:
    def test_get_cell_text_returns_text(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.get_cell_text(0, 0, 0)
        assert result["text"] == "Alpha"

    def test_get_cell_text_empty_cell(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.get_cell_text(0, 0, 2)
        assert result["text"] == ""

    def test_get_cell_text_returns_keys(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.get_cell_text(0, 0, 0)
        assert "table_index" in result
        assert "row_index" in result
        assert "col_index" in result
        assert "text" in result
        assert result["table_index"] == 0
        assert result["row_index"] == 0
        assert result["col_index"] == 0

    def test_get_cell_text_invalid_table(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.get_cell_text(99, 0, 0)

    def test_get_cell_text_invalid_row(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.get_cell_text(0, 99, 0)

    def test_get_cell_text_invalid_col(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.get_cell_text(0, 0, 99)


class TestCopyTable:
    def test_copy_table_creates_second_table(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        assert len(doc.get_tables()) == 1
        doc.copy_table(0)
        assert len(doc.get_tables()) == 2

    def test_copy_table_content_matches(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        doc.copy_table(0)
        tables = doc.get_tables()
        assert tables[0]["cells"] == tables[1]["cells"]

    def test_copy_table_returns_indices(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        result = doc.copy_table(0)
        assert result == {"source_index": 0, "new_index": 1}

    def test_copy_table_paraids_regenerated(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        doc.copy_table(0)
        from docx_mcp.document.base import W14, W

        tree = doc._tree("word/document.xml")
        para_ids = [p.get(f"{W14}paraId") for p in tree.iter(f"{W}p") if p.get(f"{W14}paraId")]
        # All paraIds should be unique (no duplicates between original and copy)
        assert len(para_ids) == len(set(para_ids)), "Duplicate paraIds found after copy_table"

    def test_copy_table_invalid_index(self, tmp_path):
        doc = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.copy_table(99)
