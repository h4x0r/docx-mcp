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


@pytest.fixture()
def mike_corpus_docx(tmp_path: Path) -> Path:
    """DOCX fixture covering mike-specific editing scenarios."""
    path = tmp_path / "mike_corpus.docx"
    _build_mike_corpus(path)
    return path


def _build_mike_corpus(path: Path) -> None:
    files = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _TOP_RELS,
        "word/document.xml": _MIKE_CORPUS_DOCUMENT_XML,
        "word/_rels/document.xml.rels": _MIKE_CORPUS_DOC_RELS,
        "word/footnotes.xml": _FOOTNOTES_XML,
        "word/endnotes.xml": _ENDNOTES_XML,
        "word/styles.xml": _STYLES_XML,
        "word/settings.xml": _SETTINGS_XML,
        "word/header1.xml": _HEADER_XML,
        "docProps/core.xml": _CORE_XML,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.strip())


# Smallest valid 1x1 PNG (67 bytes)
_TINY_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
    b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
    b"\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00"
    b"\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
)


def _build_fixture(path: Path) -> None:
    files = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _TOP_RELS,
        "word/document.xml": _DOCUMENT_XML,
        "word/_rels/document.xml.rels": _DOC_RELS,
        "word/footnotes.xml": _FOOTNOTES_XML,
        "word/endnotes.xml": _ENDNOTES_XML,
        "word/header1.xml": _HEADER_XML,
        "word/styles.xml": _STYLES_XML,
        "word/settings.xml": _SETTINGS_XML,
        "docProps/core.xml": _CORE_XML,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.strip())
        zf.writestr("word/media/image1.png", _TINY_PNG)


# ── XML templates ───────────────────────────────────────────────────────────

_CONTENT_TYPES = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/header1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"""

_TOP_RELS = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
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
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    Target="endnotes.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>
"""

# 6 body paragraphs + 9 table cell paragraphs + 1 image paragraph = 16
#   00000001  H1  "Introduction"
#   00000002  body with footnoteRef #1
#   00000003  H2  "Background"
#   00000004  body with bookmark, searchable text
#   00000005  body with bold formatting + endnote ref #1
#   00000006  body with split runs (two <w:r> elements)
#   00000007  body with embedded image
#   Table: 3 rows x 2 cols (paraIds A01-A09)
_DOCUMENT_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
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
      <w:r>
        <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
        <w:endnoteReference w:id="1"/>
      </w:r>
    </w:p>
    <w:p w14:paraId="00000006" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">First </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> last</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000007" w14:textId="77777777">
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Picture 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId6"/>
                  </pic:blipFill>
                  <pic:spPr/>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr w14:paraId="0000A001" w14:textId="77777777">
        <w:tc><w:p w14:paraId="0000A002" w14:textId="77777777"><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="0000A003" w14:textId="77777777"><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr w14:paraId="0000A004" w14:textId="77777777">
        <w:tc><w:p w14:paraId="0000A005" w14:textId="77777777"><w:r><w:t>Row 1 A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="0000A006" w14:textId="77777777"><w:r><w:t>Row 1 B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr w14:paraId="0000A007" w14:textId="77777777">
        <w:tc><w:p w14:paraId="0000A008" w14:textId="77777777"><w:r><w:t>Row 2 A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="0000A009" w14:textId="77777777"><w:r><w:t>Row 2 B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
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

_ENDNOTES_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:endnote w:type="separator" w:id="-1">
    <w:p w14:paraId="00000E11" w14:textId="77777777">
      <w:r><w:separator/></w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="00000E12" w14:textId="77777777">
      <w:r><w:continuationSeparator/></w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:id="1">
    <w:p w14:paraId="00000E13" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
      <w:r><w:t>Endnote reference material.</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>
"""

_STYLES_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
  </w:style>
</w:styles>
"""

_SETTINGS_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:settings>
"""

_CORE_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <dc:subject>Test Subject</dc:subject>
  <dc:description>Test Description</dc:description>
  <cp:lastModifiedBy>Test Editor</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-06-15T12:00:00Z</dcterms:modified>
</cp:coreProperties>
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
  <w:p w14:paraId="00000E02" w14:textId="77777777">
    <w:r><w:t>Document Header Text</w:t></w:r>
  </w:p>
</w:hdr>
"""

_MIKE_CORPUS_DOC_RELS = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Target="footnotes.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    Target="endnotes.xml"/>
</Relationships>
"""

# Paragraphs:
#   00000101  body with w:hyperlink wrapping runs
#   00000102  body with narrow no-break space U+202F
#   00000103  body with smart quotes
#   00000104  "twelve months" para 1 (ambiguous phrase - first)
#   00000105  "twelve months" para 2 (ambiguous phrase - second)
#   00000106  full-paragraph replacement candidate
#   00000107  para with existing w:del + w:ins (pre-existing tracked changes)
#   00000108-0A  three paragraphs for batch-edit test
_MIKE_CORPUS_DOCUMENT_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000101" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">Payment terms require </w:t></w:r>
      <w:hyperlink w:anchor="clause1">
        <w:r><w:t>thirty calendar days</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t xml:space="preserve"> for settlement.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000102" w14:textId="77777777">
      <w:r><w:t>The purchase price is $5 000 per unit.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000103" w14:textId="77777777">
      <w:r><w:t>The “Effective Date” means the date of execution.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000104" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">Section 1: The term shall be twelve months from commencement.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000105" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">Section 2: The renewal period is also twelve months unless terminated.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000106" w14:textId="77777777">
      <w:r><w:t>This entire clause shall be deleted upon execution.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000107" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">The fee is </w:t></w:r>
      <w:del w:id="50" w:author="Reviewer" w:date="2025-01-01T00:00:00Z">
        <w:r><w:delText>one hundred</w:delText></w:r>
      </w:del>
      <w:ins w:id="51" w:author="Reviewer" w:date="2025-01-01T00:00:00Z">
        <w:r><w:t>two hundred</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> dollars per month.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000108" w14:textId="77777777">
      <w:r><w:t>Clause A governs the initial deposit amount.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000109" w14:textId="77777777">
      <w:r><w:t>Clause B governs the ongoing maintenance fee.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="0000010A" w14:textId="77777777">
      <w:r><w:t>Clause C governs the termination penalty.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
