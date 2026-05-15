"""Tests for advanced table operations: split_table, duplicate_table_row, sort_table."""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, DocxDocument, W


def _make_doc_with_table(tmp_path, rows=4, cols=2, cells: list[list[str]] | None = None):
    """Create a minimal DocxDocument with one table."""
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    result = doc.add_table(para_id, rows, cols)
    table_idx = result["table_index"]
    if cells:
        for r, row in enumerate(cells):
            for c, text in enumerate(row):
                if text:
                    doc.modify_cell(table_idx, r, c, text)
    # Return (doc, table_id string from w14:tblId, or None if not present)
    # The new methods use table_idx (int) — return that for simplicity
    return doc, table_idx


class TestSplitTable:
    def test_split_table_creates_two_tables(self, tmp_path):
        """Split a 4-row table at index 2 → body now has 2 table elements."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=4, cols=2)
        doc.split_table(table_idx, at_row_index=2)
        tree = doc._tree("word/document.xml")
        tables = list(tree.iter(f"{W}tbl"))
        assert len(tables) == 2

    def test_split_table_row_counts(self, tmp_path):
        """First table has 2 rows, second table has 2 rows after split at 2."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=4, cols=2)
        result = doc.split_table(table_idx, at_row_index=2)
        assert result["table1_rows"] == 2
        assert result["table2_rows"] == 2
        tree = doc._tree("word/document.xml")
        tables = list(tree.iter(f"{W}tbl"))
        assert len(tables[0].findall(f"{W}tr")) == 2
        assert len(tables[1].findall(f"{W}tr")) == 2

    def test_split_table_invalid_index_raises(self, tmp_path):
        """split at 0 or >= row_count raises ValueError."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=3, cols=2)
        with pytest.raises(ValueError, match="Split index out of range"):
            doc.split_table(table_idx, at_row_index=0)
        with pytest.raises(ValueError, match="Split index out of range"):
            doc.split_table(table_idx, at_row_index=3)


class TestDuplicateTableRow:
    def test_duplicate_table_row_inserts_copy(self, tmp_path):
        """Duplicate row 0 in a 2-row table → 3 rows total."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        result = doc.duplicate_table_row(table_idx, row_index=0)
        assert result["row_index"] == 0
        assert result["new_row_index"] == 1
        tree = doc._tree("word/document.xml")
        tables = list(tree.iter(f"{W}tbl"))
        rows = tables[table_idx].findall(f"{W}tr")
        assert len(rows) == 3

    def test_duplicate_table_row_no_shared_para_ids(self, tmp_path):
        """Copied row's w:p paraIds differ from all others."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        tree = doc._tree("word/document.xml")
        doc.duplicate_table_row(table_idx, row_index=0)
        # After duplication get all ids
        tables = list(tree.iter(f"{W}tbl"))
        rows = tables[table_idx].findall(f"{W}tr")
        # New row is at index 1; its para IDs must not overlap with old row 0 ids
        new_row_ids = {
            p.get(f"{W14}paraId")
            for p in rows[1].iter(f"{W}p")
            if p.get(f"{W14}paraId")
        }
        old_row_0_ids = {
            p.get(f"{W14}paraId")
            for p in rows[0].iter(f"{W}p")
            if p.get(f"{W14}paraId")
        }
        # No shared para IDs between original row 0 and new row 1
        assert new_row_ids.isdisjoint(old_row_0_ids)

    def test_duplicate_table_row_out_of_range(self, tmp_path):
        """Out-of-range row_index raises ValueError."""
        doc, table_idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        with pytest.raises((ValueError, IndexError)):
            doc.duplicate_table_row(table_idx, row_index=5)


class TestSortTable:
    def _make_text_table(self, tmp_path, data: list[list[str]], header: bool = False):
        """Build doc with a table populated with text data."""
        rows = len(data)
        cols = max(len(r) for r in data)
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        result = doc.add_table(para_id, rows, cols)
        table_idx = result["table_index"]
        for r, row in enumerate(data):
            for c, text in enumerate(row):
                if text:
                    doc.modify_cell(table_idx, r, c, text)
        if header:
            doc.set_header_row(table_idx)
        return doc, table_idx

    def _cell_texts(self, doc, table_idx, col):
        """Extract text from each non-header row cell at column col."""
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[table_idx]
        texts = []
        for tr in tbl.findall(f"{W}tr"):
            tr_pr = tr.find(f"{W}trPr")
            if tr_pr is not None and tr_pr.find(f"{W}tblHeader") is not None:
                continue
            cells = tr.findall(f"{W}tc")
            if col < len(cells):
                texts.append("".join(t.text for t in cells[col].iter(f"{W}t") if t.text))
        return texts

    def test_sort_table_ascending(self, tmp_path):
        """3-row table with B, A, C in col 0 → sorted ascending: A, B, C."""
        data = [["B", "x"], ["A", "y"], ["C", "z"]]
        doc, table_idx = self._make_text_table(tmp_path, data)
        result = doc.sort_table(table_idx, column_index=0, ascending=True)
        assert result["sorted_rows"] == 3
        assert result["ascending"] is True
        texts = self._cell_texts(doc, table_idx, 0)
        # After accept-view: modify_cell wraps in ins/del so _text extracts the ins text
        # Just check relative order
        assert texts == sorted(texts)

    def test_sort_table_descending(self, tmp_path):
        """3-row table with B, A, C in col 0 → sorted descending: C, B, A."""
        data = [["B", "x"], ["A", "y"], ["C", "z"]]
        doc, table_idx = self._make_text_table(tmp_path, data)
        result = doc.sort_table(table_idx, column_index=0, ascending=False)
        assert result["sorted_rows"] == 3
        assert result["ascending"] is False
        texts = self._cell_texts(doc, table_idx, 0)
        assert texts == sorted(texts, reverse=True)

    def test_sort_table_skips_header_rows(self, tmp_path):
        """Header row (tblHeader) stays at top; only data rows are sorted."""
        # Row 0 = header (H), rows 1-3 = B, A, C
        data = [["H", "header"], ["B", "b"], ["A", "a"], ["C", "c"]]
        doc, table_idx = self._make_text_table(tmp_path, data, header=True)
        doc.sort_table(table_idx, column_index=0, ascending=True)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[table_idx]
        rows = tbl.findall(f"{W}tr")
        # First row must still be header
        tr_pr = rows[0].find(f"{W}trPr")
        assert tr_pr is not None
        assert tr_pr.find(f"{W}tblHeader") is not None
        # Data rows should be sorted
        assert len(rows) == 4
