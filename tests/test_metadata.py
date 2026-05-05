"""
RED tests — Gap 5: Metadata sanitization.

sanitize_metadata(level=1, redact_authors_as="") strips identifying information
from a DOCX while preserving tracked-change markup and document content.

Level 1: Remove rsid attributes (revision save IDs) from all elements.
Level 2: + Replace author names in w:ins / w:del with *redact_authors_as*
           (defaults to "Anonymous" if empty string supplied).
Level 3: + Clear creator / lastModifiedBy from docProps/core.xml
         + Clear Company / Application from docProps/app.xml
         + Remove attachedTemplate reference from word/settings.xml
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


# ─────────────────────────────────────────────────────────────────────────────
# Fixture helpers
# ─────────────────────────────────────────────────────────────────────────────

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/docProps/core.xml"'
    ' ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>'
)

_TOP_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="word/document.xml"/>'
    '<Relationship Id="rId2"'
    ' Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"'
    ' Target="docProps/core.xml"/>'
    '<Relationship Id="rId3"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"'
    ' Target="docProps/app.xml"/>'
    '</Relationships>'
)

_CORE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties'
    ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
    ' xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:creator>Alice Smith</dc:creator>'
    '<cp:lastModifiedBy>Bob Jones</cp:lastModifiedBy>'
    '<dc:title>Confidential Agreement</dc:title>'
    '<cp:revision>42</cp:revision>'
    '</cp:coreProperties>'
)

_APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>Microsoft Office Word</Application>'
    '<Company>Acme Legal LLP</Company>'
    '</Properties>'
)

_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:attachedTemplate r:id="rId1"/>'
    '<w:rsids><w:rsidDel w:val="00AB1234"/><w:rsidR w:val="00CD5678"/></w:rsids>'
    '</w:settings>'
)

_WORD_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"'
    ' Target="file:///C:/Users/alice/AppData/Roaming/Microsoft/Templates/Acme.dotx"'
    ' TargetMode="External"/>'
    '</Relationships>'
)


def _build_full_docx(path: Path) -> None:
    """
    Build a DOCX with:
    - document.xml with rsid attributes, tracked changes (w:ins/w:del)
    - docProps/core.xml with creator/lastModifiedBy
    - docProps/app.xml with Company
    - word/settings.xml with attachedTemplate + rsids
    """
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document\n'
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '  <w:body>\n'
        '    <w:p w14:paraId="BB000001" w14:textId="77777777"'
        '      w:rsidR="00AB1234" w:rsidRDefault="00AB1234">\n'
        '      <w:r w:rsidRPr="00CD5678"><w:t xml:space="preserve">Unchanged text. </w:t></w:r>\n'
        '      <w:ins w:id="1" w:author="Alice Smith" w:date="2026-01-01T00:00:00Z">\n'
        '        <w:r><w:t>inserted</w:t></w:r>\n'
        '      </w:ins>\n'
        '      <w:del w:id="2" w:author="Alice Smith" w:date="2026-01-01T00:00:00Z">\n'
        '        <w:r><w:delText>deleted</w:delText></w:r>\n'
        '      </w:del>\n'
        '    </w:p>\n'
        '  </w:body>\n'
        '</w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _TOP_RELS)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("docProps/core.xml", _CORE_XML)
        zf.writestr("docProps/app.xml", _APP_XML)
        zf.writestr("word/settings.xml", _SETTINGS_XML)
        zf.writestr("word/_rels/document.xml.rels", _WORD_RELS)


def _get_doc_xml(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def _get_zip_text(path: Path, member: str) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(member).decode()


# ─────────────────────────────────────────────────────────────────────────────
# Tests
# ─────────────────────────────────────────────────────────────────────────────


class TestMetadataSanitize:

    @pytest.fixture()
    def full_docx(self, tmp_path: Path) -> Path:
        p = tmp_path / "full.docx"
        _build_full_docx(p)
        return p

    # ── Level 1: rsid stripping ───────────────────────────────────────────────

    def test_level1_removes_rsid_attributes(self, full_docx: Path, tmp_path: Path):
        """After level=1 sanitize, no w:rsidR / w:rsidRPr attributes remain."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=1, output_path=str(out)))
        root = _get_doc_xml(out)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rsid_attrs = [
            f"{{{W_NS}}}{name}"
            for name in ("rsidR", "rsidRPr", "rsidDel", "rsidRDefault", "rsidSect")
        ]
        for el in root.iter():
            for attr in rsid_attrs:
                assert attr not in el.attrib, f"{attr} still present on <{el.tag}>"

    def test_level1_preserves_tracked_changes(self, full_docx: Path, tmp_path: Path):
        """Level=1 does not remove w:ins or w:del elements."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=1, output_path=str(out)))
        root = _get_doc_xml(out)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ins = list(root.iter(f"{{{W_NS}}}ins"))
        dels = list(root.iter(f"{{{W_NS}}}del"))
        assert len(ins) == 1
        assert len(dels) == 1

    def test_level1_preserves_document_text(self, full_docx: Path, tmp_path: Path):
        """Level=1 does not alter visible text content."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=1, output_path=str(out)))
        root = _get_doc_xml(out)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        texts = "".join(t.text or "" for t in root.iter(f"{{{W_NS}}}t"))
        assert "Unchanged text." in texts

    # ── Level 2: author anonymization ─────────────────────────────────────────

    def test_level2_replaces_author_in_ins(self, full_docx: Path, tmp_path: Path):
        """Level=2 replaces w:author on w:ins with the redact_authors_as value."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=2, output_path=str(out), redact_authors_as="REDACTED"))
        root = _get_doc_xml(out)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for ins in root.iter(f"{{{W_NS}}}ins"):
            assert ins.get(f"{{{W_NS}}}author") == "REDACTED"
        for del_ in root.iter(f"{{{W_NS}}}del"):
            assert del_.get(f"{{{W_NS}}}author") == "REDACTED"

    def test_level2_defaults_to_anonymous(self, full_docx: Path, tmp_path: Path):
        """Level=2 with no redact_authors_as defaults to 'Anonymous'."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=2, output_path=str(out)))
        root = _get_doc_xml(out)
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for ins in root.iter(f"{{{W_NS}}}ins"):
            assert ins.get(f"{{{W_NS}}}author") == "Anonymous"

    # ── Level 3: document properties ─────────────────────────────────────────

    def test_level3_clears_core_xml_creator(self, full_docx: Path, tmp_path: Path):
        """Level=3 empties dc:creator and cp:lastModifiedBy in core.xml."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=3, output_path=str(out)))
        core = _get_zip_text(out, "docProps/core.xml")
        assert "Alice Smith" not in core
        assert "Bob Jones" not in core

    def test_level3_clears_app_xml_company(self, full_docx: Path, tmp_path: Path):
        """Level=3 empties Company in app.xml."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=3, output_path=str(out)))
        app = _get_zip_text(out, "docProps/app.xml")
        assert "Acme Legal LLP" not in app

    def test_level3_removes_attached_template(self, full_docx: Path, tmp_path: Path):
        """Level=3 removes the attachedTemplate element from settings.xml."""
        server.open_document(str(full_docx))
        out = tmp_path / "out.docx"
        _j(server.sanitize_metadata(level=3, output_path=str(out)))
        settings = _get_zip_text(out, "word/settings.xml")
        assert "attachedTemplate" not in settings

    def test_level3_output_path_required(self, full_docx: Path):
        """sanitize_metadata requires output_path (does not overwrite input)."""
        server.open_document(str(full_docx))
        with pytest.raises((ValueError, TypeError)):
            server.sanitize_metadata(level=1, output_path="")
