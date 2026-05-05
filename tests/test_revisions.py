"""
RED tests — Gap 4: Revision extraction as structured JSON.

get_tracked_changes() returns a JSON list of all pending tracked changes
(insertions and deletions) in the document body, in document order.

Each entry:
  {
    "type":      "insertion" | "deletion",
    "change_id": int,
    "author":    str,
    "date":      str (ISO 8601),
    "para_id":   str,
    "text":      str   (inserted or deleted text)
  }
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(s: str) -> list | dict:
    return json.loads(s)


def _minimal_docx_with_revisions(path: Path) -> None:
    """
    Build a DOCX with three paragraphs and existing tracked changes:

    Para AA000001: one insertion ("INSERTED_WORD") by Alice
    Para AA000002: one deletion ("DELETED_WORD") by Bob
    Para AA000003: plain text, no tracked changes
    """
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document\n'
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '  <w:body>\n'
        '    <w:p w14:paraId="AA000001" w14:textId="77777777">\n'
        '      <w:r><w:t xml:space="preserve">Before </w:t></w:r>\n'
        '      <w:ins w:id="1" w:author="Alice" w:date="2026-01-15T10:00:00Z">\n'
        '        <w:r><w:t>INSERTED_WORD</w:t></w:r>\n'
        '      </w:ins>\n'
        '      <w:r><w:t xml:space="preserve"> after.</w:t></w:r>\n'
        '    </w:p>\n'
        '    <w:p w14:paraId="AA000002" w14:textId="77777777">\n'
        '      <w:r><w:t xml:space="preserve">Text with </w:t></w:r>\n'
        '      <w:del w:id="2" w:author="Bob" w:date="2026-01-16T14:30:00Z">\n'
        '        <w:r><w:delText>DELETED_WORD</w:delText></w:r>\n'
        '      </w:del>\n'
        '      <w:r><w:t xml:space="preserve"> removed.</w:t></w:r>\n'
        '    </w:p>\n'
        '    <w:p w14:paraId="AA000003" w14:textId="77777777">\n'
        '      <w:r><w:t>No changes here.</w:t></w:r>\n'
        '    </w:p>\n'
        '  </w:body>\n'
        '</w:document>'
    )
    _write_zip(path, doc_xml)


def _minimal_docx_empty(path: Path) -> None:
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document\n'
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '  <w:body>\n'
        '    <w:p w14:paraId="AA000010" w14:textId="77777777">\n'
        '      <w:r><w:t>No tracked changes.</w:t></w:r>\n'
        '    </w:p>\n'
        '  </w:body>\n'
        '</w:document>'
    )
    _write_zip(path, doc_xml)


def _write_zip(path: Path, doc_xml: str) -> None:
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", doc_xml)


# ─────────────────────────────────────────────────────────────────────────────
# Tests
# ─────────────────────────────────────────────────────────────────────────────


class TestGetTrackedChanges:
    """get_tracked_changes() extracts all w:ins and w:del entries as JSON."""

    @pytest.fixture()
    def no_changes_doc(self, tmp_path: Path) -> Path:
        p = tmp_path / "empty.docx"
        _minimal_docx_empty(p)
        return p

    @pytest.fixture()
    def revision_doc(self, tmp_path: Path) -> Path:
        p = tmp_path / "revisions.docx"
        _minimal_docx_with_revisions(p)
        return p

    def test_no_tracked_changes_returns_empty_list(self, no_changes_doc: Path):
        server.open_document(str(no_changes_doc))
        result = _j(server.get_tracked_changes())
        assert result == []

    def test_insertion_entry_has_required_fields(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        insertions = [c for c in changes if c["type"] == "insertion"]
        assert len(insertions) == 1
        entry = insertions[0]
        assert entry["type"] == "insertion"
        assert entry["author"] == "Alice"
        assert entry["date"] == "2026-01-15T10:00:00Z"
        assert entry["text"] == "INSERTED_WORD"
        assert "change_id" in entry
        assert "para_id" in entry

    def test_deletion_entry_has_required_fields(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        deletions = [c for c in changes if c["type"] == "deletion"]
        assert len(deletions) == 1
        entry = deletions[0]
        assert entry["type"] == "deletion"
        assert entry["author"] == "Bob"
        assert entry["date"] == "2026-01-16T14:30:00Z"
        assert entry["text"] == "DELETED_WORD"
        assert "change_id" in entry
        assert "para_id" in entry

    def test_para_id_correctly_reported(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        insertions = [c for c in changes if c["type"] == "insertion"]
        deletions = [c for c in changes if c["type"] == "deletion"]
        assert insertions[0]["para_id"] == "AA000001"
        assert deletions[0]["para_id"] == "AA000002"

    def test_changes_returned_in_document_order(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        # Insertion is in para AA000001, deletion in AA000002 — insertion first
        assert changes[0]["type"] == "insertion"
        assert changes[1]["type"] == "deletion"

    def test_paragraph_with_no_changes_not_included(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        para_ids = [c["para_id"] for c in changes]
        assert "AA000003" not in para_ids

    def test_change_id_is_integer(self, revision_doc: Path):
        server.open_document(str(revision_doc))
        changes = _j(server.get_tracked_changes())
        for c in changes:
            assert isinstance(c["change_id"], int)

    def test_roundtrip_after_delete_text(self, revision_doc: Path):
        """Calling delete_text then get_tracked_changes includes the new deletion."""
        server.open_document(str(revision_doc))
        server.delete_text("AA000003", "No changes here")
        changes = _j(server.get_tracked_changes())
        para_ids = [c["para_id"] for c in changes]
        assert "AA000003" in para_ids
        assert len(changes) == 3  # 2 existing + 1 new
