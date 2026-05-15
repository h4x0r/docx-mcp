"""Tests for ReviewMergeMixin.merge_review_rounds (Phase 7.1)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document.errors import DocxMcpError, ErrCode

# ── Minimal DOCX builders ────────────────────────────────────────────────────

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

_TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

# Base doc: two paragraphs — "Alpha text" and "Beta text"
_BASE_DOC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:r><w:t>Alpha text</w:t></w:r>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Beta text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

# Reviewer doc: has a w:ins after "Alpha text" paragraph
_REVIEWER_INS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:r><w:t>Alpha text</w:t></w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t> inserted by alice</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Beta text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

# Second reviewer: identical ins (duplicate → should deduplicate)
_REVIEWER_INS_DUP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:r><w:t>Alpha text</w:t></w:r>
      <w:ins w:id="2" w:author="Bob" w:date="2026-01-02T00:00:00Z">
        <w:r><w:t> inserted by alice</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Beta text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

# Conflicting reviewer: same preceding text, different ins content
_REVIEWER_CONFLICT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:r><w:t>Alpha text</w:t></w:r>
      <w:ins w:id="3" w:author="Carol" w:date="2026-01-03T00:00:00Z">
        <w:r><w:t> conflicting insert</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Beta text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

# Reviewer with w:del
_REVIEWER_DEL_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:del w:id="4" w:author="Dave" w:date="2026-01-04T00:00:00Z">
        <w:r><w:delText>Alpha text</w:delText></w:r>
      </w:del>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Beta text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""


def _build_docx(path: Path, doc_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _TOP_RELS.strip())
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS.strip())
        zf.writestr("word/document.xml", doc_xml.strip())


# ── Tests ────────────────────────────────────────────────────────────────────


class TestMergeReviewRounds:
    @pytest.fixture(autouse=True)
    def _open_base(self, tmp_path: Path):
        base = tmp_path / "base.docx"
        _build_docx(base, _BASE_DOC_XML)
        server.open_document(str(base))
        self.tmp = tmp_path

    def test_single_reviewer_ins_merged(self):
        rev = self.tmp / "reviewer1.docx"
        _build_docx(rev, _REVIEWER_INS_XML)
        result = server._doc.merge_review_rounds([str(rev)])
        assert result["merged"] >= 1
        assert result["conflicts"] == []
        assert result["skipped_duplicates"] == 0

    def test_duplicate_changes_deduplicated(self):
        rev1 = self.tmp / "reviewer1.docx"
        rev2 = self.tmp / "reviewer2.docx"
        _build_docx(rev1, _REVIEWER_INS_XML)
        _build_docx(rev2, _REVIEWER_INS_DUP_XML)
        result = server._doc.merge_review_rounds([str(rev1), str(rev2)])
        assert result["merged"] == 1
        assert result["skipped_duplicates"] == 1
        assert result["conflicts"] == []

    def test_conflict_flagged(self):
        rev1 = self.tmp / "reviewer1.docx"
        rev2 = self.tmp / "reviewer_conflict.docx"
        _build_docx(rev1, _REVIEWER_INS_XML)
        _build_docx(rev2, _REVIEWER_CONFLICT_XML)
        result = server._doc.merge_review_rounds([str(rev1), str(rev2)])
        assert result["merged"] == 1
        assert len(result["conflicts"]) == 1
        conflict = result["conflicts"][0]
        assert "conflicting insert" in conflict["text"]
        assert conflict["author"] == "Carol"

    def test_missing_reviewer_raises(self):
        with pytest.raises(DocxMcpError) as exc_info:
            server._doc.merge_review_rounds([str(self.tmp / "nonexistent.docx")])
        assert exc_info.value.code == ErrCode.PART_NOT_FOUND

    def test_no_document_raises(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.merge_review_rounds([])

    def test_del_changes_merged(self):
        rev = self.tmp / "reviewer_del.docx"
        _build_docx(rev, _REVIEWER_DEL_XML)
        result = server._doc.merge_review_rounds([str(rev)])
        assert result["merged"] >= 1
        assert result["conflicts"] == []

    def test_server_tool_returns_json(self):
        rev = self.tmp / "reviewer1.docx"
        _build_docx(rev, _REVIEWER_INS_XML)
        import json

        raw = server.merge_review_rounds([str(rev)])
        data = json.loads(raw)
        assert "merged" in data
        assert "conflicts" in data
        assert "skipped_duplicates" in data
