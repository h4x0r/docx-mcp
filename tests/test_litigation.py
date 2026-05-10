"""Tests for LitigationMixin: bates_number, redact_text, generate_redaction_log, generate_privilege_log."""

from __future__ import annotations

import re
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14


def _make_doc(tmp_path: Path) -> DocxDocument:
    path = str(tmp_path / "test.docx")
    return DocxDocument.create(path)


class TestBatesNumbering:
    def test_bates_footer_contains_prefix(self, tmp_path):
        """bates_number result contains prefix in the return dict."""
        doc = _make_doc(tmp_path)
        result = doc.bates_number("ACME-", start=1, digits=6)
        assert result["prefix"] == "ACME-"
        assert result["sections_stamped"] >= 1

    def test_bates_start_override(self, tmp_path):
        """bates_number respects the start parameter."""
        doc = _make_doc(tmp_path)
        result = doc.bates_number("DOC-", start=100, digits=6)
        assert result["start"] == 100

    def test_bates_digits_padding(self, tmp_path):
        """bates_number respects the digits parameter."""
        doc = _make_doc(tmp_path)
        result = doc.bates_number("X-", start=1, digits=8)
        assert result["digits"] == 8


class TestRedactText:
    def test_redact_by_exact_text(self, tmp_path):
        """redact_text removes the run containing exact_text."""
        doc = _make_doc(tmp_path)
        # Add a paragraph with known text
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "secret phrase"
        body.insert(0, p)
        doc._mark("word/document.xml")

        result = doc.redact_text(exact_text="secret phrase", reason="Privileged")
        assert result["redacted_count"] >= 1

    def test_redacted_text_removed_from_xml(self, tmp_path):
        """After redact_text, the original text is no longer in the XML."""
        doc = _make_doc(tmp_path)
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "remove this text"
        body.insert(0, p)
        doc._mark("word/document.xml")

        doc.redact_text(exact_text="remove this text")
        tree2 = doc._tree("word/document.xml")
        all_text = "".join(el.text for el in tree2.iter(f"{W}t") if el.text)
        assert "remove this text" not in all_text

    def test_redaction_replaced_with_black_rect(self, tmp_path):
        """Redacted run is replaced with a drawing element (black rectangle)."""
        doc = _make_doc(tmp_path)
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "classified info"
        body.insert(0, p)
        doc._mark("word/document.xml")

        doc.redact_text(exact_text="classified info")
        tree2 = doc._tree("word/document.xml")
        drawings = list(tree2.iter(f"{W}drawing"))
        assert len(drawings) >= 1

    def test_redaction_log_generated(self, tmp_path):
        """redact_text populates the internal redaction log."""
        doc = _make_doc(tmp_path)
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "log this"
        body.insert(0, p)
        doc._mark("word/document.xml")

        doc.redact_text(exact_text="log this", reason="Test Reason")
        assert hasattr(doc, "_redaction_log")
        assert len(doc._redaction_log) >= 1
        assert doc._redaction_log[-1]["reason"] == "Test Reason"

    def test_redact_by_regex(self, tmp_path):
        """redact_text with pattern= redacts matching runs."""
        doc = _make_doc(tmp_path)
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "SSN: 123-45-6789"
        body.insert(0, p)
        doc._mark("word/document.xml")

        result = doc.redact_text(pattern=r"\d{3}-\d{2}-\d{4}")
        assert result["redacted_count"] >= 1


class TestPrivilegeLog:
    def test_privilege_log_is_valid_docx(self, tmp_path):
        """generate_privilege_log creates a valid .docx file."""
        doc = _make_doc(tmp_path)
        out = str(tmp_path / "priv_log.docx")
        result = doc.generate_privilege_log(out)
        assert result["entry_count"] >= 1
        assert Path(out).exists()

    def test_privilege_log_columns_present(self, tmp_path):
        """generate_privilege_log table has expected column headers."""
        doc = _make_doc(tmp_path)
        out = str(tmp_path / "priv_log2.docx")
        doc.generate_privilege_log(out)
        # Open the output as a DocxDocument and read table
        log_doc = DocxDocument(out)
        log_doc.open()
        tables = log_doc.get_tables()
        assert len(tables) >= 1
        header_row = tables[0]["cells"][0]
        joined = " ".join(header_row).lower()
        assert "author" in joined or "bates" in joined

    def test_redaction_log_docx_output(self, tmp_path):
        """generate_redaction_log creates a .docx with the redaction entries."""
        doc = _make_doc(tmp_path)
        # Make a redaction first
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        p = etree.Element(f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "redact me"
        body.insert(0, p)
        doc._mark("word/document.xml")
        doc.redact_text(exact_text="redact me", reason="Test")

        out = str(tmp_path / "redact_log.docx")
        result = doc.generate_redaction_log(out)
        assert Path(out).exists()
        assert result["entry_count"] >= 1
