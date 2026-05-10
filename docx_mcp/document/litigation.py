"""Litigation tools mixin: bates numbering, redaction, privilege/redaction logs."""

from __future__ import annotations

import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from .base import W, W14
from .pii import _make_redaction_drawing, _next_drawing_id


class LitigationMixin:
    """Litigation support: Bates numbering, text redaction, privilege/redaction logs."""

    # ── Bates Numbering ────────────────────────────────────────────────────────

    def bates_number(
        self,
        prefix: str,
        start: int = 1,
        digits: int = 6,
        position: str = "footer-right",
    ) -> dict:
        """Add a Bates stamp to the document body as a styled paragraph.

        Appends a right-aligned paragraph at the end of the document body
        containing the Bates stamp (prefix + zero-padded start number).
        Word preserves sequential page numbering from this anchor.

        Args:
            prefix: Bates prefix string (e.g. "ACME-").
            start: Starting Bates number.
            digits: Zero-padding width for the number.
            position: Hint for stamp position (currently "footer-right").

        Returns:
            dict with prefix, start, digits, sections_stamped.
        """
        stamp_text = f"{prefix}{str(start).zfill(digits)}"

        tree = self._require("word/document.xml")
        body = tree.find(f"{W}body")
        if body is None:
            raise RuntimeError("No body element in document.xml")

        # Create a right-aligned paragraph for the stamp
        p = etree.Element(f"{W}p")
        p.set(f"{W14}paraId", self._new_para_id())
        p.set(f"{W14}textId", "77777777")

        pPr = etree.SubElement(p, f"{W}pPr")
        jc = etree.SubElement(pPr, f"{W}jc")
        jc.set(f"{W}val", "right")

        r = etree.SubElement(p, f"{W}r")
        rPr = etree.SubElement(r, f"{W}rPr")
        b = etree.SubElement(rPr, f"{W}b")  # bold stamp
        t = etree.SubElement(r, f"{W}t")
        t.text = stamp_text

        # Insert before the last sectPr (or at end if no sectPr)
        children = list(body)
        sect_pr = body.find(f"{W}sectPr")
        if sect_pr is not None:
            idx = list(body).index(sect_pr)
            body.insert(idx, p)
        else:
            body.append(p)

        self._mark("word/document.xml")

        return {
            "prefix": prefix,
            "start": start,
            "digits": digits,
            "sections_stamped": 1,
            "stamp": stamp_text,
        }

    # ── Text Redaction ────────────────────────────────────────────────────────

    def redact_text(
        self,
        pattern: str | None = None,
        para_ids: list[str] | None = None,
        exact_text: str | None = None,
        reason: str = "",
    ) -> dict:
        """True redaction: remove text from XML and replace with black rectangle.

        Matches runs by exact_text equality or regex pattern on the run's text.
        The original w:t element is removed and replaced with a DrawingML
        solid-black inline rectangle — no text content remains in the OOXML.

        Args:
            pattern: Regex pattern to match run text.
            para_ids: Optional list of paragraph paraId values to limit scope.
            exact_text: Exact string to match against run text.
            reason: Reason for redaction (stored in log).

        Returns:
            dict with redacted_count and log list.
        """
        if not hasattr(self, "_redaction_log"):
            self._redaction_log: list[dict] = []

        if exact_text is None and pattern is None:
            raise ValueError("Provide either exact_text or pattern")

        tree = self._require("word/document.xml")
        used_ids: set[int] = set()
        redacted_count = 0
        log_entries: list[dict] = []

        for para in tree.iter(f"{W}p"):
            # Filter by para_ids if specified
            if para_ids is not None:
                pid = para.get(f"{W14}paraId")
                if pid not in para_ids:
                    continue

            # Collect runs to redact
            runs_to_redact: list[etree._Element] = []
            for run in list(para.findall(f"{W}r")):
                t_el = run.find(f"{W}t")
                if t_el is None or not t_el.text:
                    continue
                run_text = t_el.text

                matched = False
                if exact_text is not None and run_text == exact_text:
                    matched = True
                elif exact_text is not None and exact_text in run_text:
                    matched = True
                elif pattern is not None and re.search(pattern, run_text):
                    matched = True

                if matched:
                    runs_to_redact.append(run)

            for run in runs_to_redact:
                t_el = run.find(f"{W}t")
                original_text = t_el.text if t_el is not None else ""
                original_len = len(original_text) if original_text else 0

                # Find position of run in parent paragraph
                run_idx = list(para).index(run)

                # Build redaction drawing run
                drawing_run = _make_redaction_drawing(tree, used_ids)

                # Replace the run in-place
                para.remove(run)
                para.insert(run_idx, drawing_run)

                para_id = para.get(f"{W14}paraId", "")
                entry = {
                    "para_id": para_id,
                    "original_length": original_len,
                    "reason": reason,
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                }
                log_entries.append(entry)
                self._redaction_log.append(entry)
                redacted_count += 1

        if redacted_count > 0:
            self._mark("word/document.xml")

        return {"redacted_count": redacted_count, "log": log_entries}

    # ── Redaction Log ────────────────────────────────────────────────────────

    def generate_redaction_log(self, output_path: str = "") -> dict:
        """Write a DOCX table of all redactions made this session.

        Creates a new DOCX at output_path containing a table with columns:
        #, Para ID, Chars Removed, Reason, Reviewer, Date.

        Args:
            output_path: Destination path. If empty, writes to a temp file.

        Returns:
            dict with path and entry_count.
        """
        if not hasattr(self, "_redaction_log"):
            self._redaction_log = []

        if not output_path:
            tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            output_path = tmp.name
            tmp.close()

        log = self._redaction_log
        headers = ["#", "Para ID", "Chars Removed", "Reason", "Reviewer", "Date"]
        rows = len(log) + 1  # header row + data rows
        cols = len(headers)

        # Import here to avoid circular imports
        from docx_mcp.document import DocxDocument

        log_doc = DocxDocument.create(output_path)

        # Find the first paragraph para_id in the fresh doc
        doc_tree = log_doc._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")
        paras = body.findall(f"{W}p") if body is not None else []
        if not paras:
            raise RuntimeError("Fresh document has no paragraphs")
        para_id = paras[0].get(f"{W14}paraId")
        if para_id is None:
            raise RuntimeError("First paragraph has no paraId")

        result = log_doc.add_table(para_id, rows=rows, cols=cols)
        tbl_idx = result["table_index"]

        # Fill header row
        for col_idx, header in enumerate(headers):
            log_doc.modify_cell(tbl_idx, 0, col_idx, header)

        # Fill data rows
        for row_idx, entry in enumerate(log, start=1):
            log_doc.modify_cell(tbl_idx, row_idx, 0, str(row_idx))
            log_doc.modify_cell(tbl_idx, row_idx, 1, entry.get("para_id", ""))
            log_doc.modify_cell(tbl_idx, row_idx, 2, str(entry.get("original_length", "")))
            log_doc.modify_cell(tbl_idx, row_idx, 3, entry.get("reason", ""))
            log_doc.modify_cell(tbl_idx, row_idx, 4, "")
            log_doc.modify_cell(tbl_idx, row_idx, 5, entry.get("timestamp", ""))

        log_doc.save(output_path, backup=False)

        return {"path": output_path, "entry_count": len(log)}

    # ── Privilege Log ────────────────────────────────────────────────────────

    def generate_privilege_log(self, output_path: str = "") -> dict:
        """Generate a privilege log DOCX from document metadata.

        Creates a new DOCX with a privilege log table with columns:
        Bates Range, Author, Date, Subject, Privilege Basis.

        Args:
            output_path: Destination path. If empty, writes to a temp file.

        Returns:
            dict with path and entry_count.
        """
        if not output_path:
            tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            output_path = tmp.name
            tmp.close()

        props = self.get_properties()
        author = props.get("creator", "")
        created = props.get("created", "")
        title = props.get("title", "")

        headers = ["Bates Range", "Author", "Date", "Subject", "Privilege Basis"]
        rows = 2  # header + 1 data row
        cols = len(headers)

        from docx_mcp.document import DocxDocument

        log_doc = DocxDocument.create(output_path)

        doc_tree = log_doc._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")
        paras = body.findall(f"{W}p") if body is not None else []
        if not paras:
            raise RuntimeError("Fresh document has no paragraphs")
        para_id = paras[0].get(f"{W14}paraId")
        if para_id is None:
            raise RuntimeError("First paragraph has no paraId")

        result = log_doc.add_table(para_id, rows=rows, cols=cols)
        tbl_idx = result["table_index"]

        # Header row
        for col_idx, header in enumerate(headers):
            log_doc.modify_cell(tbl_idx, 0, col_idx, header)

        # Data row from document metadata
        log_doc.modify_cell(tbl_idx, 1, 0, "")          # Bates Range (TBD)
        log_doc.modify_cell(tbl_idx, 1, 1, author)
        log_doc.modify_cell(tbl_idx, 1, 2, created)
        log_doc.modify_cell(tbl_idx, 1, 3, title)
        log_doc.modify_cell(tbl_idx, 1, 4, "")          # Privilege Basis (TBD)

        log_doc.save(output_path, backup=False)

        return {"path": output_path, "entry_count": 1}
