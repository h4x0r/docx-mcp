"""
RED tests — Gap 1 completion: OOXML encoding artifact normalization.

Covers:
  - Invisible characters that must be REMOVED (not substituted with space):
      soft hyphen U+00AD, zero-width space U+200B, word-joiner U+2060
  - Non-standard space characters that collapse to regular space:
      thin space U+2009, figure space U+2007, narrow no-break space U+202F
  - Case-insensitive matching via ignore_case=True parameter
  - Original document characters preserved in tracked-change output regardless
    of normalization applied during matching
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import W, W14


def _j(s: str) -> dict:
    return json.loads(s)


def _para_xml(para_id: str) -> etree._Element:
    doc = server._doc._trees["word/document.xml"]
    for p in doc.iter(f"{W}p"):
        if p.get(f"{W14}paraId") == para_id:
            return p
    raise KeyError(para_id)


def _minimal_docx(path: Path, body_xml: str) -> None:
    """Write a minimal .docx whose document body is *body_xml*."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document\n'
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '  <w:body>\n'
        + body_xml
        + '\n  </w:body>\n</w:document>'
    )
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
# 1. Soft hyphen (U+00AD) — invisible, must be removed not substituted
# ─────────────────────────────────────────────────────────────────────────────

class TestSoftHyphen:
    """
    Soft hyphens are used by Word as line-break hints inside words.
    They render as nothing; the word appears unbroken.
    Query without soft hyphen must match document text with soft hyphen.
    """

    PARA_ID = "00000030"

    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "softhyphen.docx"
        # "consid\u00aderation" — soft hyphen mid-word
        body = (
            f'    <w:p w14:paraId="{self.PARA_ID}" w14:textId="77777777">\n'
            '      <w:r><w:t xml:space="preserve">consid\u00aderation of the matter</w:t></w:r>\n'
            '    </w:p>'
        )
        _minimal_docx(path, body)
        server.open_document(str(path))

    def test_query_without_soft_hyphen_matches_doc(self):
        """'consideration' matches document 'consid\u00aderation'."""
        result = _j(server.delete_text(self.PARA_ID, "consideration"))
        assert result["type"] == "deletion"

    def test_deleted_output_preserves_original_soft_hyphen(self):
        """w:delText contains the original chars including U+00AD."""
        server.delete_text(self.PARA_ID, "consideration")
        p = _para_xml(self.PARA_ID)
        del_texts = "".join(dt.text or "" for dt in p.iter(f"{W}delText"))
        assert "\u00ad" in del_texts


# ─────────────────────────────────────────────────────────────────────────────
# 2. Zero-width space (U+200B) — invisible, must be removed
# ─────────────────────────────────────────────────────────────────────────────

class TestZeroWidthSpace:
    """
    ZWS is inserted by Word in some CJK or URL contexts.
    It is invisible; queries must not need to include it.
    """

    PARA_ID = "00000031"

    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "zws.docx"
        # "force\u200bmajeure" — ZWS between two words (no visible space)
        body = (
            f'    <w:p w14:paraId="{self.PARA_ID}" w14:textId="77777777">\n'
            '      <w:r><w:t xml:space="preserve">The force\u200bmajeure clause applies.</w:t></w:r>\n'
            '    </w:p>'
        )
        _minimal_docx(path, body)
        server.open_document(str(path))

    def test_query_without_zws_matches_doc_with_zws(self):
        """'forcemajeure' matches 'force\u200bmajeure' — ZWS removed, not spaced."""
        result = _j(server.delete_text(self.PARA_ID, "forcemajeure"))
        assert result["type"] == "deletion"

    def test_deleted_output_preserves_original_zws(self):
        server.delete_text(self.PARA_ID, "forcemajeure")
        p = _para_xml(self.PARA_ID)
        del_texts = "".join(dt.text or "" for dt in p.iter(f"{W}delText"))
        assert "\u200b" in del_texts


# ─────────────────────────────────────────────────────────────────────────────
# 3. Word-joiner (U+2060) — invisible, must be removed
# ─────────────────────────────────────────────────────────────────────────────

class TestWordJoiner:
    PARA_ID = "00000032"

    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "wj.docx"
        # Word uses U+2060 to prevent line breaks at certain positions
        body = (
            f'    <w:p w14:paraId="{self.PARA_ID}" w14:textId="77777777">\n'
            '      <w:r><w:t xml:space="preserve">indemnification\u2060obligations</w:t></w:r>\n'
            '    </w:p>'
        )
        _minimal_docx(path, body)
        server.open_document(str(path))

    def test_query_without_word_joiner_matches(self):
        """'indemnificationobligations' matches doc with word-joiner between."""
        result = _j(server.delete_text(self.PARA_ID, "indemnificationobligations"))
        assert result["type"] == "deletion"


# ─────────────────────────────────────────────────────────────────────────────
# 4. Non-standard spaces — normalize to regular space then collapse
# ─────────────────────────────────────────────────────────────────────────────

class TestNonStandardSpaces:
    """
    Thin space (U+2009), figure space (U+2007), narrow no-break space (U+202F)
    should all match as regular spaces in queries.
    """

    PARA_ID = "00000033"

    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "spaces.docx"
        # "30\u2009days" — thin space between number and unit
        body = (
            f'    <w:p w14:paraId="{self.PARA_ID}" w14:textId="77777777">\n'
            '      <w:r><w:t xml:space="preserve">within 30\u2009days from the date</w:t></w:r>\n'
            '    </w:p>'
        )
        _minimal_docx(path, body)
        server.open_document(str(path))

    def test_regular_space_query_matches_thin_space_doc(self):
        """Query '30 days' (regular space) matches document '30\u2009days' (thin space)."""
        result = _j(server.delete_text(self.PARA_ID, "30 days"))
        assert result["type"] == "deletion"


# ─────────────────────────────────────────────────────────────────────────────
# 5. Case-insensitive matching
# ─────────────────────────────────────────────────────────────────────────────

class TestCaseInsensitive:
    """
    ignore_case=True enables case-folded matching.
    Matched text in output always uses the document's original casing.
    Default (ignore_case=False) is case-sensitive.
    """

    PARA_ID = "00000034"

    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "case.docx"
        body = (
            f'    <w:p w14:paraId="{self.PARA_ID}" w14:textId="77777777">\n'
            '      <w:r><w:t xml:space="preserve">This Agreement shall remain in force.</w:t></w:r>\n'
            '    </w:p>'
        )
        _minimal_docx(path, body)
        server.open_document(str(path))

    def test_case_sensitive_by_default(self):
        """Default: 'agreement' does not match 'Agreement' — raises ValueError."""
        with pytest.raises(ValueError):
            server.delete_text(self.PARA_ID, "agreement")

    def test_ignore_case_matches_title_case(self):
        """ignore_case=True: 'agreement' matches 'Agreement'."""
        result = _j(server.delete_text(self.PARA_ID, "agreement", ignore_case=True))
        assert result["type"] == "deletion"

    def test_ignore_case_preserves_original_casing_in_output(self):
        """The w:delText contains 'Agreement' (original), not 'agreement' (query)."""
        server.delete_text(self.PARA_ID, "agreement", ignore_case=True)
        p = _para_xml(self.PARA_ID)
        del_texts = "".join(dt.text or "" for dt in p.iter(f"{W}delText"))
        assert "Agreement" in del_texts
        assert "agreement" not in del_texts

    def test_ignore_case_on_context_before(self):
        """context_before is also matched case-insensitively when ignore_case=True."""
        # "agreement" appears once; context_before="THIS " (uppercase) should still match
        result = _j(server.delete_text(
            self.PARA_ID, "Agreement",
            context_before="THIS ",
            ignore_case=True,
        ))
        assert result["type"] == "deletion"

    def test_ignore_case_replace_text(self):
        """replace_text with ignore_case=True matches case-insensitively."""
        result = _j(server.replace_text(
            self.PARA_ID,
            find="agreement",
            replace="Contract",
            ignore_case=True,
        ))
        assert result["type"] == "replacement"
