"""Tables mixin: read and write table content."""

from __future__ import annotations

import copy
import csv
import io

from lxml import etree

from .base import W14, W, _now_iso, _preserve

_CM_TO_TWIPS = 567


class TablesMixin:
    """Table operations."""

    def get_tables(self) -> list[dict]:
        """Get all tables with their content."""
        doc = self._require("word/document.xml")
        tables = []
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            rows = []
            for tr in tbl.findall(f"{W}tr"):
                cells = []
                for tc in tr.findall(f"{W}tc"):
                    cells.append(self._text(tc))
                rows.append(cells)
            col_count = len(rows[0]) if rows else 0
            tables.append(
                {
                    "index": idx,
                    "row_count": len(rows),
                    "col_count": col_count,
                    "cells": rows,
                }
            )
        return tables

    def _get_table(self, table_idx: int) -> etree._Element:
        """Get table element by index, raising IndexError if not found."""
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")
        return tables[table_idx]

    def add_table(
        self,
        para_id: str,
        rows: int,
        cols: int,
        *,
        author: str = "Claude",
    ) -> dict:
        """Insert a new table after a paragraph with tracked insertion.

        Args:
            para_id: paraId of the paragraph to insert after.
            rows: Number of rows.
            cols: Number of columns.
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        now = _now_iso()
        cid = self._next_markup_id(doc)

        tbl = etree.Element(f"{W}tbl")
        # Table properties
        tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
        tbl_style = etree.SubElement(tbl_pr, f"{W}tblStyle")
        tbl_style.set(f"{W}val", "TableGrid")
        tbl_w = etree.SubElement(tbl_pr, f"{W}tblW")
        tbl_w.set(f"{W}w", "0")
        tbl_w.set(f"{W}type", "auto")
        # Track change on table
        ins = etree.SubElement(tbl_pr, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        # Grid columns
        grid = etree.SubElement(tbl, f"{W}tblGrid")
        for _ in range(cols):
            etree.SubElement(grid, f"{W}gridCol")

        # Rows and cells
        for _ in range(rows):
            tr = etree.SubElement(tbl, f"{W}tr")
            tr.set(f"{W14}paraId", self._new_para_id())
            tr.set(f"{W14}textId", "77777777")
            for _ in range(cols):
                tc = etree.SubElement(tr, f"{W}tc")
                p = etree.SubElement(tc, f"{W}p")
                p.set(f"{W14}paraId", self._new_para_id())
                p.set(f"{W14}textId", "77777777")

        para.addnext(tbl)
        self._mark("word/document.xml")

        # Calculate table index
        table_idx = list(doc.iter(f"{W}tbl")).index(tbl)
        return {"table_index": table_idx, "rows": rows, "cols": cols, "inserted": True}

    def modify_cell(
        self,
        table_idx: int,
        row: int,
        col: int,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Modify a table cell with tracked changes.

        Args:
            table_idx: Table index (0-based).
            row: Row index (0-based).
            col: Column index (0-based).
            text: New cell text.
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        rows = tbl.findall(f"{W}tr")
        if row < 0 or row >= len(rows):
            raise IndexError(f"Row {row} out of range (have {len(rows)})")
        cells = rows[row].findall(f"{W}tc")
        if col < 0 or col >= len(cells):
            raise IndexError(f"Column {col} out of range (have {len(cells)})")

        tc = cells[col]
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Find first paragraph in cell
        para = tc.find(f"{W}p")
        if para is None:
            para = etree.SubElement(tc, f"{W}p")
            para.set(f"{W14}paraId", self._new_para_id())
            para.set(f"{W14}textId", "77777777")

        # Delete existing runs
        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or not t_el.text:
                continue
            rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(rpr) if rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, t_el.text)
            parent.insert(pos, del_el)
            cid = self._next_markup_id(doc)

        # Insert new text
        ins = etree.SubElement(para, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)
        r = etree.SubElement(ins, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        self._mark("word/document.xml")
        return {"modified": True, "table_index": table_idx, "cell": [row, col], "text": text}

    def add_table_row(
        self,
        table_idx: int,
        row_idx: int | None = None,
        cells: list[str] | None = None,
        *,
        author: str = "Claude",
    ) -> dict:
        """Add a row to a table with tracked insertion.

        Args:
            table_idx: Table index (0-based).
            row_idx: Insert at this index. None = append at end.
            cells: Cell text content. None = empty cells.
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Determine column count from existing rows
        existing_rows = tbl.findall(f"{W}tr")
        col_count = len(existing_rows[0].findall(f"{W}tc")) if existing_rows else 1

        # Build new row
        tr = etree.Element(f"{W}tr")
        tr.set(f"{W14}paraId", self._new_para_id())
        tr.set(f"{W14}textId", "77777777")
        # Track change on row
        tr_pr = etree.SubElement(tr, f"{W}trPr")
        ins = etree.SubElement(tr_pr, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        for i in range(col_count):
            tc = etree.SubElement(tr, f"{W}tc")
            p = etree.SubElement(tc, f"{W}p")
            p.set(f"{W14}paraId", self._new_para_id())
            p.set(f"{W14}textId", "77777777")
            if cells and i < len(cells):
                r = etree.SubElement(p, f"{W}r")
                t = etree.SubElement(r, f"{W}t")
                _preserve(t, cells[i])

        # Insert or append
        if row_idx is not None and row_idx < len(existing_rows):
            existing_rows[row_idx].addprevious(tr)
            final_idx = row_idx
        else:
            tbl.append(tr)
            final_idx = len(existing_rows)

        self._mark("word/document.xml")
        new_row_count = len(tbl.findall(f"{W}tr"))
        return {
            "table_index": table_idx,
            "row_index": final_idx,
            "row_count": new_row_count,
            "inserted": True,
        }

    def delete_table_row(
        self,
        table_idx: int,
        row_idx: int,
        *,
        author: str = "Claude",
    ) -> dict:
        """Delete a table row with tracked changes.

        Args:
            table_idx: Table index (0-based).
            row_idx: Row index to delete (0-based).
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range (have {len(rows)})")

        tr = rows[row_idx]
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Mark row itself as deleted via trPr
        tr_pr = tr.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            tr.insert(0, tr_pr)
        del_el = etree.SubElement(tr_pr, f"{W}del")
        del_el.set(f"{W}id", str(cid))
        del_el.set(f"{W}author", author)
        del_el.set(f"{W}date", now)

        # Mark all runs in cells as deleted
        for tc in tr.findall(f"{W}tc"):
            for para in tc.findall(f"{W}p"):
                for run_el in list(para.findall(f"{W}r")):
                    t_el = run_el.find(f"{W}t")
                    if t_el is None or not t_el.text:
                        continue
                    cid = self._next_markup_id(doc)
                    rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(rpr) if rpr is not None else None
                    parent = run_el.getparent()
                    pos = list(parent).index(run_el)
                    parent.remove(run_el)
                    d = etree.Element(f"{W}del")
                    d.set(f"{W}id", str(cid))
                    d.set(f"{W}author", author)
                    d.set(f"{W}date", now)
                    dr = etree.SubElement(d, f"{W}r")
                    if rpr_bytes:
                        dr.append(etree.fromstring(rpr_bytes))
                    dt = etree.SubElement(dr, f"{W}delText")
                    _preserve(dt, t_el.text)
                    parent.insert(pos, d)

        self._mark("word/document.xml")
        return {"table_index": table_idx, "row_index": row_idx, "deleted": True}

    # ── New table improvement methods ────────────────────────────────────────

    def merge_cells(
        self,
        table_index: int,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
    ) -> dict:
        """Merge a rectangular range of cells.

        Horizontal merge (same row): set w:gridSpan on start cell, remove
        intermediate cells.
        Vertical merge (same col): set w:vMerge val="restart" on top cell,
        w:vMerge (no val) on continuation cells.
        Rectangular: apply gridSpan on each row, then vMerge across rows.

        Returns {"merged_rows": int, "merged_cols": int, "table_index": int}.
        """
        tbl = self._get_table(table_index)
        rows_els = tbl.findall(f"{W}tr")

        merged_rows = end_row - start_row + 1
        merged_cols = end_col - start_col + 1

        if merged_rows < 1 or merged_cols < 1:
            raise ValueError("end indices must be >= start indices")

        if end_row >= len(rows_els):
            raise ValueError(f"end_row {end_row} out of range")

        for r in range(start_row, end_row + 1):
            row_el = rows_els[r]
            cells = row_el.findall(f"{W}tc")
            if end_col >= len(cells):
                raise ValueError(f"end_col {end_col} out of range in row {r}")

            if merged_cols > 1:
                # Set gridSpan on the first cell in the range
                first_tc = cells[start_col]
                tc_pr = first_tc.find(f"{W}tcPr")
                if tc_pr is None:
                    tc_pr = etree.Element(f"{W}tcPr")
                    first_tc.insert(0, tc_pr)
                grid_span = tc_pr.find(f"{W}gridSpan")
                if grid_span is None:
                    grid_span = etree.SubElement(tc_pr, f"{W}gridSpan")
                grid_span.set(f"{W}val", str(merged_cols))

                # Remove intermediate cells
                for c in range(start_col + 1, end_col + 1):
                    # Re-fetch cells after removals since list changed
                    current_cells = row_el.findall(f"{W}tc")
                    # The cell at start_col+1 becomes the next one to remove
                    if len(current_cells) > start_col + 1:
                        row_el.remove(current_cells[start_col + 1])

            if merged_rows > 1:
                # Apply vMerge to the start_col cell in this row
                tc = row_el.findall(f"{W}tc")[start_col]
                tc_pr = tc.find(f"{W}tcPr")
                if tc_pr is None:
                    tc_pr = etree.Element(f"{W}tcPr")
                    tc.insert(0, tc_pr)
                vmerge = tc_pr.find(f"{W}vMerge")
                if vmerge is None:
                    vmerge = etree.SubElement(tc_pr, f"{W}vMerge")
                if r == start_row:
                    vmerge.set(f"{W}val", "restart")
                else:
                    # Continuation: no val attribute
                    if f"{W}val" in vmerge.attrib:
                        del vmerge.attrib[f"{W}val"]

        self._mark("word/document.xml")
        return {
            "table_index": table_index,
            "merged_rows": merged_rows,
            "merged_cols": merged_cols,
        }

    def set_header_row(self, table_index: int) -> dict:
        """Mark the first row as a header row that repeats across page breaks.

        Sets w:tblHeader on the first row's trPr.
        Returns {"table_index": int, "set": True}.
        """
        tbl = self._get_table(table_index)
        rows_els = tbl.findall(f"{W}tr")
        if not rows_els:
            raise ValueError("Table has no rows")

        first_row = rows_els[0]
        tr_pr = first_row.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            first_row.insert(0, tr_pr)

        if tr_pr.find(f"{W}tblHeader") is None:
            etree.SubElement(tr_pr, f"{W}tblHeader")

        self._mark("word/document.xml")
        return {"table_index": table_index, "set": True}

    def set_column_widths(self, table_index: int, widths_cm: list[float]) -> dict:
        """Set column widths in the table grid.

        Updates w:tblGrid/w:gridCol/@w:w (twips) and w:tcW on each cell.
        1 cm = 567 twips.
        Returns {"table_index": int, "column_count": int, "widths_cm": list[float]}.
        Raises ValueError if len(widths_cm) != column_count.
        """
        tbl = self._get_table(table_index)

        # Determine column count from tblGrid or first row
        grid = tbl.find(f"{W}tblGrid")
        if grid is not None:
            grid_cols = grid.findall(f"{W}gridCol")
            col_count = len(grid_cols)
        else:
            rows_els = tbl.findall(f"{W}tr")
            if rows_els:
                first_row_cells = rows_els[0].findall(f"{W}tc")
                col_count = 0
                for tc in first_row_cells:
                    tcPr = tc.find(f"{W}tcPr")
                    grid_span = tcPr.find(f"{W}gridSpan") if tcPr is not None else None
                    if grid_span is not None:
                        try:
                            col_count += int(grid_span.get(f"{W}val", "1"))
                        except ValueError:
                            col_count += 1
                    else:
                        col_count += 1
            else:
                col_count = 0

        if len(widths_cm) != col_count:
            raise ValueError(
                f"len(widths_cm)={len(widths_cm)} != column_count={col_count}"
            )

        twips = [int(w * _CM_TO_TWIPS) for w in widths_cm]

        # Update tblGrid
        if grid is not None:
            for i, gc in enumerate(grid_cols):
                gc.set(f"{W}w", str(twips[i]))
        else:
            # Create tblGrid after tblPr (or at start)
            grid = etree.Element(f"{W}tblGrid")
            tbl_pr = tbl.find(f"{W}tblPr")
            if tbl_pr is not None:
                tbl_pr.addnext(grid)
            else:
                tbl.insert(0, grid)
            for tw in twips:
                gc = etree.SubElement(grid, f"{W}gridCol")
                gc.set(f"{W}w", str(tw))

        # Update tcW on each cell
        for row_el in tbl.findall(f"{W}tr"):
            cells = row_el.findall(f"{W}tc")
            for i, tc in enumerate(cells):
                if i >= col_count:
                    break
                tc_pr = tc.find(f"{W}tcPr")
                if tc_pr is None:
                    tc_pr = etree.Element(f"{W}tcPr")
                    tc.insert(0, tc_pr)
                tc_w = tc_pr.find(f"{W}tcW")
                if tc_w is None:
                    tc_w = etree.SubElement(tc_pr, f"{W}tcW")
                tc_w.set(f"{W}type", "dxa")
                tc_w.set(f"{W}w", str(twips[i]))

        self._mark("word/document.xml")
        return {
            "table_index": table_index,
            "column_count": col_count,
            "widths_cm": widths_cm,
        }

    def csv_to_table(
        self,
        para_id: str,
        csv_text: str,
        header_row: bool = True,
    ) -> dict:
        """Insert a table from CSV text after a paragraph.

        Returns {"table_index": int, "rows": int, "cols": int}.
        """
        reader = csv.reader(io.StringIO(csv_text))
        data = list(reader)
        if not data:
            raise ValueError("csv_text is empty")

        num_rows = len(data)
        num_cols = max(len(r) for r in data)

        result = self.add_table(para_id, num_rows, num_cols)
        table_index = result["table_index"]

        for r, row in enumerate(data):
            for c, cell_text in enumerate(row):
                self.modify_cell(table_index, r, c, cell_text)

        if header_row and num_rows > 0:
            self.set_header_row(table_index)

        return {"table_index": table_index, "rows": num_rows, "cols": num_cols}

    def table_to_csv(self, table_index: int) -> dict:
        """Export table content as CSV string.

        Returns {"csv": str, "rows": int, "cols": int}.
        """
        tbl = self._get_table(table_index)
        rows_data = []
        col_count = 0
        for tr in tbl.findall(f"{W}tr"):
            row_cells = [self._text(tc) for tc in tr.findall(f"{W}tc")]
            rows_data.append(row_cells)
            col_count = max(col_count, len(row_cells))

        buf = io.StringIO()
        writer = csv.writer(buf)
        for row in rows_data:
            writer.writerow(row)

        return {"csv": buf.getvalue(), "rows": len(rows_data), "cols": col_count}

    def delete_table(self, table_idx: int) -> dict:
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")
        tbl = tables[table_idx]
        tbl.getparent().remove(tbl)
        self._mark("word/document.xml")
        return {"deleted": table_idx}

    def add_column_to_table(self, table_idx: int, header_text: str = "") -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        for i, tr in enumerate(rows):
            tc = etree.SubElement(tr, f"{W}tc")
            p = etree.SubElement(tc, f"{W}p")
            p.set(f"{W14}paraId", self._new_para_id())
            p.set(f"{W14}textId", "77777777")
            if i == 0 and header_text:
                r = etree.SubElement(p, f"{W}r")
                t = etree.SubElement(r, f"{W}t")
                t.text = header_text
        # Keep tblGrid in sync
        tbl_pr = tbl.find(f"{W}tblPr")
        tbl_grid = tbl.find(f"{W}tblGrid")
        if tbl_grid is None:
            tbl_grid = etree.Element(f"{W}tblGrid")
            if tbl_pr is not None:
                tbl_pr.addnext(tbl_grid)
            else:
                tbl.insert(0, tbl_grid)
        etree.SubElement(tbl_grid, f"{W}gridCol")
        col_count = len(tbl.findall(f"{W}tr")[0].findall(f"{W}tc")) if rows else 0
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "columns": col_count}

    def delete_column_from_table(self, table_idx: int, col_idx: int) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        # Validate all rows first to avoid partial mutation
        for tr in rows:
            cells = tr.findall(f"{W}tc")
            if col_idx < 0 or col_idx >= len(cells):
                raise IndexError(f"Column index {col_idx} out of range (have {len(cells)})")
        for tr in rows:
            cells = tr.findall(f"{W}tc")
            tr.remove(cells[col_idx])
        # Keep tblGrid in sync
        tbl_grid = tbl.find(f"{W}tblGrid")
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(f"{W}gridCol")
            if col_idx < len(grid_cols):
                tbl_grid.remove(grid_cols[col_idx])
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "col_idx": col_idx}

    def set_cell_width(self, table_idx: int, row_idx: int, col_idx: int, width_mm: float) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range")
        tc = cells[col_idx]
        tc_pr = tc.find(f"{W}tcPr")
        if tc_pr is None:
            tc_pr = etree.Element(f"{W}tcPr")
            tc.insert(0, tc_pr)
        tc_w = tc_pr.find(f"{W}tcW")
        if tc_w is None:
            tc_w = etree.SubElement(tc_pr, f"{W}tcW")
        dxa = round(width_mm * 1440 / 25.4)
        tc_w.set(f"{W}w", str(dxa))
        tc_w.set(f"{W}type", "dxa")
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "col_idx": col_idx, "width_dxa": dxa}

    def set_cell_vertical_alignment(
        self, table_idx: int, row_idx: int, col_idx: int, alignment: str
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range")
        tc = cells[col_idx]
        tc_pr = tc.find(f"{W}tcPr")
        if tc_pr is None:
            tc_pr = etree.Element(f"{W}tcPr")
            tc.insert(0, tc_pr)
        v_align = tc_pr.find(f"{W}vAlign")
        if v_align is None:
            v_align = etree.SubElement(tc_pr, f"{W}vAlign")
        v_align.set(f"{W}val", alignment)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "col_idx": col_idx, "alignment": alignment}

    def set_row_height(
        self, table_idx: int, row_idx: int, height_mm: float, rule: str = "exact"
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        tr = rows[row_idx]
        tr_pr = tr.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            tr.insert(0, tr_pr)
        tr_height = tr_pr.find(f"{W}trHeight")
        if tr_height is None:
            tr_height = etree.SubElement(tr_pr, f"{W}trHeight")
        dxa = round(height_mm * 1440 / 25.4)
        tr_height.set(f"{W}val", str(dxa))
        tr_height.set(f"{W}hRule", rule)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "height_dxa": dxa}

    def set_table_alignment(self, table_idx: int, alignment: str) -> dict:
        tbl = self._get_table(table_idx)
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        jc = tbl_pr.find(f"{W}jc")
        if jc is None:
            jc = etree.SubElement(tbl_pr, f"{W}jc")
        jc.set(f"{W}val", alignment)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "alignment": alignment}

    def set_table_borders(
        self,
        table_idx: int,
        border_style: str = "single",
        color: str = "000000",
        size: int = 4,
    ) -> dict:
        tbl = self._get_table(table_idx)
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        existing = tbl_pr.find(f"{W}tblBorders")
        if existing is not None:
            tbl_pr.remove(existing)
        borders = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
            el = etree.SubElement(borders, f"{W}{side}")
            el.set(f"{W}val", border_style)
            el.set(f"{W}sz", str(size))
            el.set(f"{W}space", "0")
            el.set(f"{W}color", color)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "border_style": border_style, "color": color, "size": size}

    def set_cell_shading(
        self,
        table_idx: int,
        row_idx: int,
        col_idx: int,
        fill_color: str,
        pattern: str = "clear",
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range (have {len(rows)})")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range (have {len(cells)})")
        tc = cells[col_idx]
        tc_pr = tc.find(f"{W}tcPr")
        if tc_pr is None:
            tc_pr = etree.Element(f"{W}tcPr")
            tc.insert(0, tc_pr)
        shd = tc_pr.find(f"{W}shd")
        if shd is None:
            shd = etree.SubElement(tc_pr, f"{W}shd")
        shd.set(f"{W}val", pattern)
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", fill_color)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "col_idx": col_idx, "fill_color": fill_color}

    def set_table_style(self, table_idx: int, style_name: str) -> dict:
        tbl = self._get_table(table_idx)
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        tbl_style = tbl_pr.find(f"{W}tblStyle")
        if tbl_style is None:
            tbl_style = etree.SubElement(tbl_pr, f"{W}tblStyle")
        tbl_style.set(f"{W}val", style_name)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "style_name": style_name}

    # ── Advanced table operations ────────────────────────────────────────────

    def split_table(self, table_idx: int, at_row_index: int) -> dict:
        """Split a table at at_row_index into two separate tables.

        Rows 0..at_row_index-1 stay in table 1; rows at_row_index..end go to
        the new table 2 inserted immediately after table 1 in the body.

        Args:
            table_idx: 0-based table index.
            at_row_index: 0-based row index; must be > 0 and < row_count.

        Returns:
            {"table1_rows": int, "table2_rows": int}
        """
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        row_count = len(rows)
        if at_row_index <= 0 or at_row_index >= row_count:
            raise ValueError(
                f"Split index out of range: at_row_index={at_row_index}, row_count={row_count}"
            )

        # Build table 2 as a deep copy of the original (preserves tblPr, tblGrid, etc.)
        tbl2 = copy.deepcopy(tbl)
        # Clear all w:tr children from tbl2, then repopulate with split rows
        for tr in tbl2.findall(f"{W}tr"):
            tbl2.remove(tr)
        for tr in rows[at_row_index:]:
            tbl2.append(copy.deepcopy(tr))

        # Remove split rows from original table 1
        for tr in rows[at_row_index:]:
            tbl.remove(tr)

        # Insert tbl2 immediately after tbl in its parent
        parent = tbl.getparent()
        tbl_pos = list(parent).index(tbl)
        parent.insert(tbl_pos + 1, tbl2)

        self._mark("word/document.xml")
        return {"table1_rows": at_row_index, "table2_rows": row_count - at_row_index}

    def duplicate_table_row(self, table_idx: int, row_index: int) -> dict:
        """Deep-copy a table row and insert the copy immediately after it.

        Clears w14:paraId / w14:textId attributes on copied paragraphs and
        assigns fresh unique IDs to avoid duplicates.

        Args:
            table_idx: 0-based table index.
            row_index: 0-based row to duplicate.

        Returns:
            {"row_index": row_index, "new_row_index": row_index + 1}
        """
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_index < 0 or row_index >= len(rows):
            raise ValueError(
                f"Row index {row_index} out of range (have {len(rows)})"
            )

        original = rows[row_index]
        new_row = copy.deepcopy(original)

        # Replace paraId / textId attributes on all w:p in the copy
        for p in new_row.iter(f"{W}p"):
            if p.get(f"{W14}paraId") is not None:
                p.set(f"{W14}paraId", self._new_para_id())
            if p.get(f"{W14}textId") is not None:
                p.set(f"{W14}textId", self._new_para_id())

        # Insert after the original row
        original.addnext(new_row)

        self._mark("word/document.xml")
        return {"row_index": row_index, "new_row_index": row_index + 1}

    def sort_table(self, table_idx: int, column_index: int, ascending: bool = True) -> dict:
        """Sort the non-header rows of a table by the text content of a column.

        Header rows (rows with w:tr/w:trPr/w:tblHeader) are kept at the top
        in their original order.

        Args:
            table_idx: 0-based table index.
            column_index: 0-based column to sort by.
            ascending: True for A→Z, False for Z→A.

        Returns:
            {"sorted_rows": int, "column_index": int, "ascending": bool}
        """
        tbl = self._get_table(table_idx)
        all_rows = tbl.findall(f"{W}tr")

        header_rows = []
        data_rows = []
        for tr in all_rows:
            tr_pr = tr.find(f"{W}trPr")
            if tr_pr is not None and tr_pr.find(f"{W}tblHeader") is not None:
                header_rows.append(tr)
            else:
                data_rows.append(tr)

        def _row_sort_key(tr: etree._Element) -> str:
            cells = tr.findall(f"{W}tc")
            if column_index < len(cells):
                return "".join(t.text for t in cells[column_index].iter(f"{W}t") if t.text)
            return ""

        data_rows.sort(key=_row_sort_key, reverse=not ascending)

        # Remove all rows from table, then re-append in order
        for tr in all_rows:
            tbl.remove(tr)
        for tr in header_rows + data_rows:
            tbl.append(tr)

        self._mark("word/document.xml")
        return {"sorted_rows": len(data_rows), "column_index": column_index, "ascending": ascending}
