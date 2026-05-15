"""Tests for raw XML part access — list_parts, read_part, write_part."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document.errors import DocxMcpError, ErrCode


def _j(result: str) -> dict | list:
    return json.loads(result)


# ── helpers ──────────────────────────────────────────────────────────────────


def _open(test_docx: Path) -> None:
    server.open_document(str(test_docx))


# ── TestListParts ─────────────────────────────────────────────────────────────


class TestListParts:
    def test_returns_list(self, test_docx: Path):
        _open(test_docx)
        result = _require_doc().list_parts()
        assert isinstance(result, list)

    def test_includes_document_xml(self, test_docx: Path):
        _open(test_docx)
        result = _require_doc().list_parts()
        assert "word/document.xml" in result

    def test_sorted(self, test_docx: Path):
        _open(test_docx)
        result = _require_doc().list_parts()
        assert result == sorted(result)


# ── TestReadPart ──────────────────────────────────────────────────────────────


class TestReadPart:
    def test_reads_document_xml(self, test_docx: Path):
        _open(test_docx)
        result = _require_doc().read_part("word/document.xml")
        assert "<w:document" in result["xml"]

    def test_returns_dict_with_part_and_xml(self, test_docx: Path):
        _open(test_docx)
        result = _require_doc().read_part("word/document.xml")
        assert result["part"] == "word/document.xml"
        assert "xml" in result

    def test_part_not_found_raises(self, test_docx: Path):
        _open(test_docx)
        with pytest.raises(DocxMcpError) as exc_info:
            _require_doc().read_part("word/nonexistent.xml")
        assert exc_info.value.code == ErrCode.PART_NOT_FOUND

    def test_error_hint_mentions_list_parts(self, test_docx: Path):
        _open(test_docx)
        with pytest.raises(DocxMcpError) as exc_info:
            _require_doc().read_part("word/nonexistent.xml")
        assert "list_parts" in exc_info.value.hint


# ── TestWritePart ─────────────────────────────────────────────────────────────


class TestWritePart:
    def test_write_then_read_roundtrip(self, test_docx: Path, tmp_path: Path):
        _open(test_docx)
        doc = _require_doc()
        # Build a minimal well-formed replacement for styles.xml
        minimal_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:style/>"
            "</w:styles>"
        )
        doc.write_part("word/styles.xml", minimal_xml)
        result = doc.read_part("word/styles.xml")
        assert "w:styles" in result["xml"]

    def test_invalid_xml_raises(self, test_docx: Path):
        _open(test_docx)
        with pytest.raises(DocxMcpError) as exc_info:
            _require_doc().write_part("word/document.xml", "<bad>not closed")
        assert exc_info.value.code == ErrCode.OOXML_INVALID

    def test_returns_bytes_written(self, test_docx: Path):
        _open(test_docx)
        minimal_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:style/>"
            "</w:styles>"
        )
        result = _require_doc().write_part("word/styles.xml", minimal_xml)
        assert result["bytes_written"] > 0

    def test_marks_part_dirty(self, test_docx: Path, tmp_path: Path):
        _open(test_docx)
        doc = _require_doc()
        minimal_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:style/>"
            "</w:styles>"
        )
        doc.write_part("word/styles.xml", minimal_xml)
        assert "word/styles.xml" in doc._modified

        # Save and reload to confirm the change persisted
        out = str(tmp_path / "saved.docx")
        doc.save(out)

        server.close_document()
        server.open_document(out)
        reloaded = server._doc.read_part("word/styles.xml")
        assert "w:style" in reloaded["xml"]


# ── TestRawPartsServer ────────────────────────────────────────────────────────


class TestRawPartsServer:
    def test_list_parts_server_tool(self, test_docx: Path):
        server.open_document(str(test_docx))
        result = _j(server.list_parts())
        assert isinstance(result, list)
        assert "word/document.xml" in result

    def test_read_part_server_tool(self, test_docx: Path):
        server.open_document(str(test_docx))
        result = _j(server.read_part("word/document.xml"))
        assert result["part"] == "word/document.xml"
        assert "<w:document" in result["xml"]

    def test_write_part_invalid_xml_error(self, test_docx: Path):
        server.open_document(str(test_docx))
        with pytest.raises(DocxMcpError) as exc_info:
            server.write_part("word/document.xml", "<bad>")
        assert exc_info.value.code == ErrCode.OOXML_INVALID


# ── private helper (not a fixture) ───────────────────────────────────────────


def _require_doc():
    """Access the server's current document directly."""
    assert server._doc is not None, "No document open"
    return server._doc
