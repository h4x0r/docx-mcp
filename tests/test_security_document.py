"""Security tests for document-layer hardening (V1–V4)."""
from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError, ErrCode


def _make_doc(tmp_path: Path) -> DocxDocument:
    """Return an open DocxDocument for a minimal valid DOCX."""
    from tests.conftest import _build_fixture
    path = tmp_path / "test.docx"
    _build_fixture(path)
    doc = DocxDocument(str(path))
    doc.open()
    return doc


# ── V4: BadZipFile ────────────────────────────────────────────────────────────

def test_corrupt_zip_raises_docxmcperror(tmp_path: Path):
    """A file with garbage bytes raises DocxMcpError(OOXML_INVALID), not BadZipFile."""
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"this is not a zip file at all")
    doc = DocxDocument(str(bad))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.OOXML_INVALID


def test_truncated_zip_raises_docxmcperror(tmp_path: Path):
    """A truncated ZIP raises DocxMcpError(OOXML_INVALID)."""
    bad = tmp_path / "truncated.docx"
    bad.write_bytes(b"PK\x03\x04")  # ZIP magic but incomplete
    doc = DocxDocument(str(bad))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.OOXML_INVALID


def test_empty_file_raises_docxmcperror(tmp_path: Path):
    """An empty file raises DocxMcpError(OOXML_INVALID)."""
    bad = tmp_path / "empty.docx"
    bad.write_bytes(b"")
    doc = DocxDocument(str(bad))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.OOXML_INVALID


# ── V1: ZipSlip ───────────────────────────────────────────────────────────────

def _make_zipslip_docx(tmp_path: Path, evil_entry: str) -> Path:
    """Build a fake DOCX (valid ZIP structure) with a traversal entry name."""
    path = tmp_path / "zipslip.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("_rels/.rels", "<Relationships/>")
        zf.writestr("word/document.xml", "<document/>")
        zf.writestr(evil_entry, "pwned")
    path.write_bytes(buf.getvalue())
    return path


def test_zipslip_dotdot_raises(tmp_path: Path):
    """ZIP entry with ../ path raises DocxMcpError(UNSAFE_PATH)."""
    path = _make_zipslip_docx(tmp_path, "../../evil.txt")
    doc = DocxDocument(str(path))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.UNSAFE_PATH


def test_zipslip_absolute_raises(tmp_path: Path):
    """ZIP entry with absolute path raises DocxMcpError(UNSAFE_PATH)."""
    path = _make_zipslip_docx(tmp_path, "/etc/cron.d/evil")
    doc = DocxDocument(str(path))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.UNSAFE_PATH


# ── V2: XXE ───────────────────────────────────────────────────────────────────

def _make_xxe_docx(tmp_path: Path) -> Path:
    """DOCX with an XXE payload in document.xml."""
    path = tmp_path / "xxe.docx"
    evil_xml = (
        '<?xml version="1.0"?>'
        '<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
        "<document>&xxe;</document>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("_rels/.rels", "<Relationships/>")
        zf.writestr("word/document.xml", evil_xml)
        zf.writestr("word/_rels/document.xml.rels", "<Relationships/>")
    path.write_bytes(buf.getvalue())
    return path


def test_xxe_entity_not_resolved(tmp_path: Path):
    """XXE payload in document.xml must not resolve external entities."""
    from lxml import etree
    path = _make_xxe_docx(tmp_path)
    doc = DocxDocument(str(path))
    doc.open()
    tree = doc._trees.get("word/document.xml")
    if tree is not None:
        content = etree.tostring(tree).decode()
        assert "root:" not in content, "XXE entity was resolved — /etc/passwd leaked"
    doc.close()


# ── V3: ZIP bomb ─────────────────────────────────────────────────────────────

def _make_too_many_entries_docx(tmp_path: Path, n: int) -> Path:
    """DOCX ZIP with n entries (entry count bomb)."""
    path = tmp_path / "bomb_entries.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        for i in range(n):
            zf.writestr(f"word/junk{i}.bin", b"x")
    path.write_bytes(buf.getvalue())
    return path


def test_too_many_zip_entries_raises(tmp_path: Path):
    """DOCX with >10 000 entries raises DocxMcpError(OOXML_INVALID)."""
    path = _make_too_many_entries_docx(tmp_path, 10_001)
    doc = DocxDocument(str(path))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.OOXML_INVALID


def _make_declared_size_bomb_docx(tmp_path: Path) -> Path:
    """DOCX ZIP that claims huge uncompressed size via a crafted local header."""
    path = tmp_path / "bomb_size.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("_rels/.rels", "<Relationships/>")
        zf.writestr("word/document.xml", "<document/>")
    raw = bytearray(buf.getvalue())
    # Patch uncompressed size at offset 22 in local file header to 600 MB
    struct.pack_into("<I", raw, 22, 629_145_600)
    path.write_bytes(bytes(raw))
    return path


def test_zip_bomb_declared_size_raises(tmp_path: Path):
    """DOCX claiming >500 MB uncompressed raises DocxMcpError(OOXML_INVALID)."""
    path = _make_declared_size_bomb_docx(tmp_path)
    doc = DocxDocument(str(path))
    with pytest.raises(DocxMcpError) as exc_info:
        doc.open()
    assert exc_info.value.code == ErrCode.OOXML_INVALID
