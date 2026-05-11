"""RED tests for section extended tools: get_sections, set_section_columns, delete_section_break."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import W, W14


def _j(result: str) -> dict | list:
    return json.loads(result)


# ── helpers ──────────────────────────────────────────────────────────────────

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

_BASE_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

_BASE_TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_BASE_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>"""

_BASE_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr/>
  </w:style>
</w:styles>"""

_BASE_SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"""


def _make_docx(tmp_path: Path, document_xml: str) -> Path:
    path = tmp_path / "test.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _BASE_CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _BASE_TOP_RELS.strip())
        zf.writestr("word/document.xml", document_xml.strip())
        zf.writestr("word/_rels/document.xml.rels", _BASE_DOC_RELS.strip())
        zf.writestr("word/styles.xml", _BASE_STYLES.strip())
        zf.writestr("word/settings.xml", _BASE_SETTINGS.strip())
    return path


def _single_section_xml() -> str:
    """Minimal document with only the final body sectPr (letter size)."""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def _two_section_xml() -> str:
    """Document with one intermediate sectPr + one final sectPr."""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:pPr>
        <w:sectPr>
          <w:type w:val="nextPage"/>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="720" w:bottom="720"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>Section 1</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002" w14:textId="77777777">
      <w:r><w:t>Section 2</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1440" w:bottom="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def _para_with_sectpr_xml() -> str:
    """Document where para 00000001 has a sectPr in pPr (section break)."""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:pPr>
        <w:sectPr>
          <w:type w:val="continuous"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>Para with section break</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002" w14:textId="77777777">
      <w:r><w:t>Normal para</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>"""


# ═══════════════════════════════════════════════════════════════════════════
#  get_sections
# ═══════════════════════════════════════════════════════════════════════════


class TestGetSections:
    @pytest.fixture(autouse=True)
    def _open_single(self, tmp_path: Path):
        path = _make_docx(tmp_path, _single_section_xml())
        server.open_document(str(path))

    def test_get_sections_returns_final_section(self):
        result = _j(server.get_sections())
        assert isinstance(result, list)
        assert len(result) >= 1
        sec = result[-1]  # final section
        assert sec["page_width"] == 12240
        assert sec["page_height"] == 15840
        assert sec["orientation"] == "portrait"
        assert sec["margin_top"] == 1440
        assert sec["margin_bottom"] == 1440

    def test_get_sections_two_sections(self, tmp_path: Path):
        server.close_document()
        path = _make_docx(tmp_path, _two_section_xml())
        server.open_document(str(path))
        result = _j(server.get_sections())
        assert len(result) == 2
        # First section (intermediate)
        s0 = result[0]
        assert s0["index"] == 0
        assert s0["break_type"] == "nextPage"
        assert s0["page_width"] == 12240
        assert s0["margin_top"] == 720
        # Final section
        s1 = result[1]
        assert s1["index"] == 1
        assert s1["break_type"] == ""
        assert s1["page_width"] == 15840
        assert s1["orientation"] == "landscape"


# ═══════════════════════════════════════════════════════════════════════════
#  set_section_columns
# ═══════════════════════════════════════════════════════════════════════════


class TestSetSectionColumns:
    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = _make_docx(tmp_path, _single_section_xml())
        server.open_document(str(path))

    def test_set_section_columns_single(self):
        result = _j(server.set_section_columns(0, 1))
        assert result["section_index"] == 0
        assert result["num_columns"] == 1
        # Verify XML
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        cols = sect_pr.find(f"{W}cols")
        assert cols is not None
        assert cols.get(f"{W}num") == "1"

    def test_set_section_columns_multi(self):
        result = _j(server.set_section_columns(0, 3, equal_width=True))
        assert result["section_index"] == 0
        assert result["num_columns"] == 3
        assert result["equal_width"] is True
        # Verify XML
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        cols = sect_pr.find(f"{W}cols")
        assert cols is not None
        assert cols.get(f"{W}num") == "3"
        assert cols.get(f"{W}equalWidth") == "1"


# ═══════════════════════════════════════════════════════════════════════════
#  delete_section_break
# ═══════════════════════════════════════════════════════════════════════════


class TestDeleteSectionBreak:
    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = _make_docx(tmp_path, _para_with_sectpr_xml())
        server.open_document(str(path))

    def test_delete_section_break_removes_sectPr(self):
        result = _j(server.delete_section_break("00000001"))
        assert result["deleted"] is True
        assert result["para_id"] == "00000001"
        # Verify sectPr is gone from pPr
        doc = server._doc._trees["word/document.xml"]
        para = server._doc._find_para(doc, "00000001")
        assert para is not None
        ppr = para.find(f"{W}pPr")
        # pPr should be gone (no other children) or sectPr removed
        if ppr is not None:
            assert ppr.find(f"{W}sectPr") is None

    def test_delete_section_break_raises_when_none(self):
        with pytest.raises(ValueError, match="No section break"):
            server.delete_section_break("00000002")
