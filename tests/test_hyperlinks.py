"""Tests for HyperlinksMixin — Hyperlink CRUD."""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError, ErrCode

# ── Namespace constants used in assertions ───────────────────────────────────

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# ── Fixture helpers ──────────────────────────────────────────────────────────

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

_TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>"""

_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr>
  </w:style>
</w:styles>"""

_CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>Test</dc:creator>
</cp:coreProperties>"""

# Document with plain paras (no hyperlinks initially)
_DOCUMENT_XML_PLAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="AA000001" w14:textId="77777777">
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
    <w:p w14:paraId="AA000002" w14:textId="77777777">
      <w:r><w:t>Second paragraph</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""

# Document with a pre-existing external hyperlink
_DOCUMENT_XML_WITH_LINKS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="BB000001" w14:textId="77777777">
      <w:hyperlink r:id="rId2">
        <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Click here</w:t></w:r>
      </w:hyperlink>
    </w:p>
    <w:p w14:paraId="BB000002" w14:textId="77777777">
      <w:hyperlink w:anchor="MyBookmark">
        <w:r><w:t>Go to section</w:t></w:r>
      </w:hyperlink>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""

_DOC_RELS_WITH_LINK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com"
    TargetMode="External"/>
</Relationships>"""


def _build_docx(path: Path, *, document_xml: str = _DOCUMENT_XML_PLAIN, rels: str = _DOC_RELS) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.strip())
        zf.writestr("_rels/.rels", _TOP_RELS.strip())
        zf.writestr("word/document.xml", document_xml.strip())
        zf.writestr("word/_rels/document.xml.rels", rels.strip())
        zf.writestr("word/styles.xml", _STYLES_XML.strip())
        zf.writestr("docProps/core.xml", _CORE_XML.strip())
    return path


def _open(path: Path) -> DocxDocument:
    doc = DocxDocument(str(path))
    doc.open()
    return doc


# ── TestListHyperlinks ────────────────────────────────────────────────────────

class TestListHyperlinks:
    def test_empty_returns_empty(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        result = doc.list_hyperlinks()
        assert result == []

    def test_lists_external_hyperlink(self, tmp_path: Path):
        path = _build_docx(
            tmp_path / "doc.docx",
            document_xml=_DOCUMENT_XML_WITH_LINKS,
            rels=_DOC_RELS_WITH_LINK,
        )
        doc = _open(path)
        links = doc.list_hyperlinks()
        external = [lnk for lnk in links if lnk["type"] == "external"]
        assert len(external) == 1
        assert external[0]["url_or_anchor"] == "https://example.com"
        assert external[0]["text"] == "Click here"
        assert external[0]["id"] == "rId2"

    def test_lists_internal_hyperlink(self, tmp_path: Path):
        path = _build_docx(
            tmp_path / "doc.docx",
            document_xml=_DOCUMENT_XML_WITH_LINKS,
            rels=_DOC_RELS_WITH_LINK,
        )
        doc = _open(path)
        links = doc.list_hyperlinks()
        internal = [lnk for lnk in links if lnk["type"] == "internal"]
        assert len(internal) == 1
        assert internal[0]["url_or_anchor"] == "MyBookmark"
        assert internal[0]["text"] == "Go to section"
        assert internal[0]["id"] is None


# ── TestAddHyperlink ──────────────────────────────────────────────────────────

class TestAddHyperlink:
    def test_add_external_hyperlink(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        result = doc.add_hyperlink("AA000001", "Visit us", "https://example.org")
        assert result["para_id"] == "AA000001"
        assert result["url"] == "https://example.org"
        assert result["r_id"].startswith("rId")

    def test_hyperlink_element_in_xml(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        result = doc.add_hyperlink("AA000001", "Link", "https://test.com")
        rid = result["r_id"]
        tree = doc._require("word/document.xml")
        hyperlinks = tree.findall(f".//{{{W}}}hyperlink")
        assert len(hyperlinks) == 1
        assert hyperlinks[0].get(f"{{{R}}}id") == rid

    def test_relationship_in_rels_file(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        result = doc.add_hyperlink("AA000001", "Link", "https://test.com")
        rid = result["r_id"]
        rels = doc._require("word/_rels/document.xml.rels")
        rel = rels.find(f'{{{RELS}}}Relationship[@Id="{rid}"]')
        assert rel is not None
        assert rel.get("Target") == "https://test.com"
        assert rel.get("TargetMode") == "External"
        assert rel.get("Type") == HYPERLINK_TYPE

    def test_para_not_found_raises(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.add_hyperlink("DEADBEEF", "Bad", "https://x.com")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND

    def test_r_id_uniqueness(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        r1 = doc.add_hyperlink("AA000001", "First", "https://one.com")
        r2 = doc.add_hyperlink("AA000002", "Second", "https://two.com")
        assert r1["r_id"] != r2["r_id"]


# ── TestAddInternalLink ───────────────────────────────────────────────────────

class TestAddInternalLink:
    def test_add_internal_link(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        result = doc.add_internal_link("AA000001", "See below", "SectionAnchor")
        assert result["anchor"] == "SectionAnchor"
        assert result["para_id"] == "AA000001"

    def test_anchor_in_xml(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        doc.add_internal_link("AA000001", "See below", "SectionAnchor")
        tree = doc._require("word/document.xml")
        hyperlinks = tree.findall(f".//{{{W}}}hyperlink")
        assert len(hyperlinks) == 1
        assert hyperlinks[0].get(f"{{{W}}}anchor") == "SectionAnchor"

    def test_no_relationship_added(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        rels_before = doc._require("word/_rels/document.xml.rels")
        count_before = len(rels_before.findall(f"{{{RELS}}}Relationship"))
        doc.add_internal_link("AA000001", "See below", "SectionAnchor")
        rels_after = doc._require("word/_rels/document.xml.rels")
        count_after = len(rels_after.findall(f"{{{RELS}}}Relationship"))
        assert count_after == count_before


# ── TestRemoveHyperlink ───────────────────────────────────────────────────────

class TestRemoveHyperlink:
    def test_remove_preserves_text(self, tmp_path: Path):
        path = _build_docx(
            tmp_path / "doc.docx",
            document_xml=_DOCUMENT_XML_WITH_LINKS,
            rels=_DOC_RELS_WITH_LINK,
        )
        doc = _open(path)
        doc.remove_hyperlink("BB000001", "https://example.com")
        tree = doc._require("word/document.xml")
        # No hyperlink wrapper remaining in BB000001
        hyperlinks = tree.findall(f".//{{{W}}}hyperlink")
        external = [h for h in hyperlinks if h.get(f"{{{R}}}id") is not None]
        assert len(external) == 0
        # Text is still there — find the para and check w:t descendants
        para = doc._find_para(tree, "BB000001")
        texts = [t.text for t in para.iter(f"{{{W}}}t") if t.text]
        assert "Click here" in texts

    def test_not_found_raises(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.remove_hyperlink("AA000001", "https://nobody.com")
        assert exc_info.value.code == ErrCode.BOOKMARK_NOT_FOUND


# ── TestUpdateHyperlink ───────────────────────────────────────────────────────

class TestUpdateHyperlink:
    def test_update_url(self, tmp_path: Path):
        path = _build_docx(
            tmp_path / "doc.docx",
            document_xml=_DOCUMENT_XML_WITH_LINKS,
            rels=_DOC_RELS_WITH_LINK,
        )
        doc = _open(path)
        result = doc.update_hyperlink("rId2", "https://updated.com")
        assert result["r_id"] == "rId2"
        assert result["new_url"] == "https://updated.com"
        rels = doc._require("word/_rels/document.xml.rels")
        rel = rels.find(f'{{{RELS}}}Relationship[@Id="rId2"]')
        assert rel.get("Target") == "https://updated.com"

    def test_not_found_raises(self, tmp_path: Path):
        path = _build_docx(tmp_path / "doc.docx")
        doc = _open(path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.update_hyperlink("rId99", "https://x.com")
        assert exc_info.value.code == ErrCode.BOOKMARK_NOT_FOUND
