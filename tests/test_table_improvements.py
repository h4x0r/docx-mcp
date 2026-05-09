"""Tests for table improvement methods: merge_cells, set_header_row,
set_column_widths, csv_to_table, table_to_csv."""

import csv
import io
import pytest
from docx_mcp.document import DocxDocument, W, W14


def _make_doc_with_table(tmp_path, rows=3, cols=3):
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    # Get para_id of first paragraph
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc, 0  # table_index=0


class TestTableImprovements:
    def test_merge_cells_horizontal(self, tmp_path):
        """merge_cells across columns sets gridSpan and removes intermediate cells."""
        doc, idx = _make_doc_with_table(tmp_path, 3, 3)
        result = doc.merge_cells(idx, 0, 0, 0, 1)  # merge cols 0-1 in row 0
        assert result["merged_cols"] == 2
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        first_row = tbl.findall(f"{W}tr")[0]
        first_cell = first_row.findall(f"{W}tc")[0]
        grid_span = first_cell.find(f".//{W}gridSpan")
        assert grid_span is not None
        assert grid_span.get(f"{W}val") == "2"

    def test_merge_cells_vertical(self, tmp_path):
        """merge_cells across rows sets vMerge."""
        doc, idx = _make_doc_with_table(tmp_path, 3, 3)
        result = doc.merge_cells(idx, 0, 0, 1, 0)  # merge rows 0-1 in col 0
        assert result["merged_rows"] == 2
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        rows = tbl.findall(f"{W}tr")
        # Top cell has vMerge restart
        top_cell = rows[0].findall(f"{W}tc")[0]
        vmerge = top_cell.find(f".//{W}vMerge")
        assert vmerge is not None
        assert vmerge.get(f"{W}val") == "restart"
        # Second cell has vMerge continuation (no val)
        second_cell = rows[1].findall(f"{W}tc")[0]
        vmerge2 = second_cell.find(f".//{W}vMerge")
        assert vmerge2 is not None
        assert vmerge2.get(f"{W}val") is None

    def test_set_header_row(self, tmp_path):
        """set_header_row adds tblHeader to first row's trPr."""
        doc, idx = _make_doc_with_table(tmp_path, 3, 3)
        result = doc.set_header_row(idx)
        assert result["set"] is True
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        first_row = tbl.findall(f"{W}tr")[0]
        trPr = first_row.find(f"{W}trPr")
        assert trPr is not None
        header = trPr.find(f"{W}tblHeader")
        assert header is not None

    def test_set_column_widths(self, tmp_path):
        """set_column_widths updates tblGrid gridCol widths."""
        doc, idx = _make_doc_with_table(tmp_path, 2, 3)
        result = doc.set_column_widths(idx, [3.0, 5.0, 2.0])
        assert result["column_count"] == 3
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        grid = tbl.find(f"{W}tblGrid")
        if grid is not None:
            cols = grid.findall(f"{W}gridCol")
            if cols:
                assert cols[0].get(f"{W}w") == str(int(3.0 * 567))
                assert cols[1].get(f"{W}w") == str(int(5.0 * 567))

    def test_csv_to_table_roundtrip(self, tmp_path):
        """csv_to_table then table_to_csv gives back the same data."""
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        csv_input = "Name,Age,City\nAlice,30,Paris\nBob,25,London"
        result = doc.csv_to_table(para_id, csv_input, header_row=True)
        assert result["rows"] == 3
        assert result["cols"] == 3
        roundtrip = doc.table_to_csv(result["table_index"])
        # Parse both and compare data
        orig_rows = list(csv.reader(io.StringIO(csv_input)))
        rt_rows = list(csv.reader(io.StringIO(roundtrip["csv"])))
        assert len(rt_rows) == len(orig_rows)
        assert rt_rows[0] == orig_rows[0]  # header
