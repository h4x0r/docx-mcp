"""Test fixtures — builds a minimal valid DOCX from raw XML + zipfile."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp import server

# ── Automatic server state cleanup ──────────────────────────────────────────


@pytest.fixture(autouse=True)
def _reset_server():
    """Ensure no document leaks between tests."""
    yield
    if server._doc is not None:
        server._doc.close()
        server._doc = None


# ── DOCX fixture ────────────────────────────────────────────────────────────


@pytest.fixture()
def test_docx(tmp_path: Path) -> Path:
    """Create a minimal valid DOCX with headings, footnote, bookmark, watermark, and split runs."""
    path = tmp_path / "test.docx"
    _build_fixture(path)
    return path


def _build_fixture(path: Path) -> None:
    files = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _TOP_RELS,
        "word/document.xml": _DOCUMENT_XML,
        "word/_rels/document.xml.rels": _DOC_RELS,
        "word/footnotes.xml": _FOOTNOTES_XML,
        "word/header1.xml": _HEADER_XML,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.strip())


# ── XML templates ───────────────────────────────────────────────────────────

_CONTENT_TYPES = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/header1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>
"""

_TOP_RELS = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
"""

_DOC_RELS = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Target="footnotes.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
</Relationships>
"""

# 6 paragraphs:
#   00000001  H1  "Introduction"
#   00000002  body with footnoteRef #1
#   00000003  H2  "Background"
#   00000004  body with bookmark, searchable text
#   00000005  body with bold formatting
#   00000006  body with split runs (two <w:r> elements)
_DOCUMENT_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Introduction</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002" w14:textId="77777777">
      <w:r><w:t>This is the first paragraph with important content.</w:t></w:r>
      <w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteReference w:id="1"/>
      </w:r>
    </w:p>
    <w:p w14:paraId="00000003" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Background</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000004" w14:textId="77777777">
      <w:bookmarkStart w:id="0" w:name="section_bg"/>
      <w:r><w:t>The contract term is 30 days from the effective date.</w:t></w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
    <w:p w14:paraId="00000005" w14:textId="77777777">
      <w:r><w:rPr><w:b/></w:rPr><w:t>Final paragraph with bold review content.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000006" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">First </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> last</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

_FOOTNOTES_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:footnote w:type="separator" w:id="-1">
    <w:p w14:paraId="00000F01" w14:textId="77777777">
      <w:r><w:separator/></w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="00000F02" w14:textId="77777777">
      <w:r><w:continuationSeparator/></w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p w14:paraId="00000F03" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
      <w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteRef/>
      </w:r>
      <w:r><w:t xml:space="preserve"> </w:t></w:r>
      <w:r><w:t>See appendix A for supporting evidence.</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
"""

# Header with a DRAFT VML watermark
_HEADER_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:p w14:paraId="00000E01" w14:textId="77777777">
    <w:pPr><w:pStyle w:val="Header"/></w:pPr>
    <w:r>
      <w:pict>
        <v:shape id="_x0000_s2049" type="#_x0000_t136"
          style="position:absolute;width:527pt;height:131pt;rotation:315;z-index:-251658752"
          fillcolor="silver" stroked="f">
          <v:textpath style="font-family:Calibri;font-size:1pt" string="DRAFT"/>
        </v:shape>
      </w:pict>
    </w:r>
  </w:p>
</w:hdr>
"""
