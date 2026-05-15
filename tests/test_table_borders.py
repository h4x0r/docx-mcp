"""Tests for Task #32: set_table_borders, set_cell_shading, set_table_style."""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, DocxDocument, W


def _make_doc_with_table(tmp_path, rows=2, cols=2):
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc, 0


class TestSetTableBorders:
    def test_creates_tblBorders_with_all_six_sides(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tbl_pr = tbl.find(f"{W}tblPr")
        borders = tbl_pr.find(f"{W}tblBorders")
        assert borders is not None
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = borders.find(f"{W}{side}")
            assert el is not None, f"Missing border side: {side}"

    def test_sets_correct_val_on_each_side(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx, border_style="double")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        borders = tbl.find(f"{W}tblPr").find(f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = borders.find(f"{W}{side}")
            assert el.get(f"{W}val") == "double", f"Wrong val on {side}"

    def test_sets_correct_color_on_each_side(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx, color="FF0000")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        borders = tbl.find(f"{W}tblPr").find(f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = borders.find(f"{W}{side}")
            assert el.get(f"{W}color") == "FF0000", f"Wrong color on {side}"

    def test_sets_correct_size_on_each_side(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx, size=8)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        borders = tbl.find(f"{W}tblPr").find(f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = borders.find(f"{W}{side}")
            assert el.get(f"{W}sz") == "8", f"Wrong size on {side}"

    def test_sets_space_zero_on_each_side(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        borders = tbl.find(f"{W}tblPr").find(f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = borders.find(f"{W}{side}")
            assert el.get(f"{W}space") == "0", f"Wrong space on {side}"

    def test_returns_correct_dict(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_table_borders(idx, border_style="single", color="000000", size=4)
        assert result == {"table_idx": idx, "border_style": "single", "color": "000000", "size": 4}

    def test_overwrites_existing_tblBorders(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_borders(idx, color="FF0000")
        doc.set_table_borders(idx, color="00FF00")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tbl_pr = tbl.find(f"{W}tblPr")
        all_borders = tbl_pr.findall(f"{W}tblBorders")
        assert len(all_borders) == 1
        borders = all_borders[0]
        top = borders.find(f"{W}top")
        assert top.get(f"{W}color") == "00FF00"

    def test_out_of_range_raises_index_error(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_table_borders(99)


class TestSetCellShading:
    def test_sets_shd_with_fill_color(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_cell_shading(idx, 0, 0, fill_color="FFFF00")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tc = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")[0]
        tc_pr = tc.find(f"{W}tcPr")
        shd = tc_pr.find(f"{W}shd")
        assert shd is not None
        assert shd.get(f"{W}fill") == "FFFF00"

    def test_sets_shd_val_pattern(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_cell_shading(idx, 0, 0, fill_color="AABBCC", pattern="solid")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tc = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")[0]
        shd = tc.find(f"{W}tcPr").find(f"{W}shd")
        assert shd.get(f"{W}val") == "solid"

    def test_sets_color_auto(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_cell_shading(idx, 0, 0, fill_color="112233")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tc = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")[0]
        shd = tc.find(f"{W}tcPr").find(f"{W}shd")
        assert shd.get(f"{W}color") == "auto"

    def test_returns_correct_dict(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_cell_shading(idx, 1, 1, fill_color="CCCCCC")
        assert result == {"table_idx": idx, "row_idx": 1, "col_idx": 1, "fill_color": "CCCCCC"}

    def test_out_of_range_row_raises_index_error(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_cell_shading(idx, 99, 0, fill_color="000000")

    def test_out_of_range_col_raises_index_error(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_cell_shading(idx, 0, 99, fill_color="000000")


class TestSetTableStyle:
    def test_creates_tblStyle_with_correct_val(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_style(idx, "LightShading-Accent1")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tbl_pr = tbl.find(f"{W}tblPr")
        tbl_style = tbl_pr.find(f"{W}tblStyle")
        assert tbl_style is not None
        assert tbl_style.get(f"{W}val") == "LightShading-Accent1"

    def test_returns_correct_dict(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_table_style(idx, "TableGrid")
        assert result == {"table_idx": idx, "style_name": "TableGrid"}

    def test_overwrites_existing_tblStyle(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_table_style(idx, "TableGrid")
        doc.set_table_style(idx, "PlainTable1")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tbl_pr = tbl.find(f"{W}tblPr")
        all_styles = tbl_pr.findall(f"{W}tblStyle")
        assert len(all_styles) == 1
        assert all_styles[0].get(f"{W}val") == "PlainTable1"
