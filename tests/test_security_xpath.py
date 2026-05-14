"""Security tests for XPath DoS protection."""
from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError, ErrCode


def _open_doc(tmp_path: Path) -> DocxDocument:
    from tests.conftest import _build_fixture
    path = tmp_path / "test.docx"
    _build_fixture(path)
    doc = DocxDocument(str(path))
    doc.open()
    return doc


def test_invalid_xpath_raises_docxmcperror(tmp_path: Path):
    """Syntactically invalid XPath raises DocxMcpError(XPATH_ERROR)."""
    doc = _open_doc(tmp_path)
    with pytest.raises(DocxMcpError) as exc_info:
        doc.xpath_query("///[invalid")
    assert exc_info.value.code == ErrCode.XPATH_ERROR
    doc.close()


def test_valid_xpath_returns_results(tmp_path: Path):
    """A valid XPath query returns a dict with expected keys."""
    doc = _open_doc(tmp_path)
    result = doc.xpath_query("//w:p")
    assert "count" in result
    assert "results" in result
    doc.close()


def test_part_not_found_raises_docxmcperror(tmp_path: Path):
    """Querying a non-existent part raises DocxMcpError(PART_NOT_FOUND)."""
    doc = _open_doc(tmp_path)
    with pytest.raises(DocxMcpError) as exc_info:
        doc.xpath_query("//w:p", part="word/nonexistent.xml")
    assert exc_info.value.code == ErrCode.PART_NOT_FOUND
    doc.close()
