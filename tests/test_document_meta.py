"""RED tests — document meta tools: get_document_outline, set_document_language, set_track_changes."""  # noqa: E501

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(s: str) -> list | dict:
    return json.loads(s)


# ── Fixture helpers ──────────────────────────────────────────────────────────


def _build_outline_docx(path: Path) -> None:
    """DOCX with H1, H2, H3 headings and w14:paraId attributes."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<w:document\n"
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        "  <w:body>\n"
        '    <w:p w14:paraId="00000001" w14:textId="77777777">\n'
        '      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>\n'
        "      <w:r><w:t>Chapter One</w:t></w:r>\n"
        "    </w:p>\n"
        '    <w:p w14:paraId="00000002" w14:textId="77777777">\n'
        "      <w:r><w:t>Body text here.</w:t></w:r>\n"
        "    </w:p>\n"
        '    <w:p w14:paraId="00000003" w14:textId="77777777">\n'
        '      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>\n'
        "      <w:r><w:t>Section One</w:t></w:r>\n"
        "    </w:p>\n"
        '    <w:p w14:paraId="00000004" w14:textId="77777777">\n'
        '      <w:pPr><w:pStyle w:val="Heading3"/></w:pPr>\n'
        "      <w:r><w:t>Subsection</w:t></w:r>\n"
        "    </w:p>\n"
        '    <w:p w14:paraId="00000005" w14:textId="77777777">\n'
        '      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>\n'
        "      <w:r><w:t>Chapter Two</w:t></w:r>\n"
        "    </w:p>\n"
        "  </w:body>\n"
        "</w:document>\n"
    )
    _write_minimal_docx(path, doc_xml)


def _build_lang_docx(path: Path) -> None:
    """DOCX with a default paragraph style in styles.xml (no existing w:lang)."""
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">\n'
        '    <w:name w:val="Normal"/>\n'
        "    <w:rPr/>\n"
        "  </w:style>\n"
        '  <w:style w:type="paragraph" w:styleId="Heading1">\n'
        '    <w:name w:val="heading 1"/>\n'
        '    <w:basedOn w:val="Normal"/>\n'
        "  </w:style>\n"
        "</w:styles>\n"
    )
    _write_minimal_docx(path, _minimal_doc_xml(), styles_xml=styles_xml)


def _build_lang_no_rpr_docx(path: Path) -> None:
    """DOCX with default paragraph style but NO existing w:rPr."""
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">\n'
        '    <w:name w:val="Normal"/>\n'
        "  </w:style>\n"
        "</w:styles>\n"
    )
    _write_minimal_docx(path, _minimal_doc_xml(), styles_xml=styles_xml)


def _minimal_doc_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<w:document\n"
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        "  <w:body>\n"
        '    <w:p w14:paraId="00000001" w14:textId="77777777">\n'
        "      <w:r><w:t>Hello</w:t></w:r>\n"
        "    </w:p>\n"
        "  </w:body>\n"
        "</w:document>\n"
    )


def _write_minimal_docx(
    path: Path,
    doc_xml: str,
    styles_xml: str | None = None,
    settings_xml: str | None = None,
) -> None:
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'  # noqa: E501
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/word/document.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'  # noqa: E501
    )
    if styles_xml:
        content_types += (
            '  <Override PartName="/word/styles.xml"\n'
            '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'  # noqa: E501
        )
    if settings_xml:
        content_types += (
            '  <Override PartName="/word/settings.xml"\n'
            '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'  # noqa: E501
        )
    content_types += "</Types>\n"

    top_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1"\n'
        '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"\n'
        '    Target="word/document.xml"/>\n'
        "</Relationships>\n"
    )

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    )
    if styles_xml:
        doc_rels += (
            '  <Relationship Id="rId1"\n'
            '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"\n'
            '    Target="styles.xml"/>\n'
        )
    if settings_xml:
        doc_rels += (
            '  <Relationship Id="rId2"\n'
            '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"\n'
            '    Target="settings.xml"/>\n'
        )
    doc_rels += "</Relationships>\n"

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", doc_xml)
        if styles_xml:
            zf.writestr("word/styles.xml", styles_xml)
        if settings_xml:
            zf.writestr("word/settings.xml", settings_xml)


# ═══════════════════════════════════════════════════════════════════════════
#  get_document_outline
# ═══════════════════════════════════════════════════════════════════════════


class TestGetDocumentOutline:
    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        path = tmp_path / "outline.docx"
        _build_outline_docx(path)
        server.open_document(str(path))

    def test_returns_headings(self):
        outline = _j(server.get_document_outline())
        assert isinstance(outline, list)
        assert len(outline) == 4  # H1, H2, H3, H1

    def test_heading_fields(self):
        outline = _j(server.get_document_outline())
        h1 = outline[0]
        assert h1["level"] == 1
        assert h1["text"] == "Chapter One"
        assert h1["para_id"] == "00000001"

    def test_heading_levels_correct(self):
        outline = _j(server.get_document_outline())
        levels = [h["level"] for h in outline]
        assert levels == [1, 2, 3, 1]

    def test_max_level_filters(self):
        outline = _j(server.get_document_outline(max_level=1))
        assert len(outline) == 2
        assert all(h["level"] == 1 for h in outline)
        texts = [h["text"] for h in outline]
        assert "Chapter One" in texts
        assert "Chapter Two" in texts

    def test_max_level_2_excludes_h3(self):
        outline = _j(server.get_document_outline(max_level=2))
        assert len(outline) == 3
        assert all(h["level"] <= 2 for h in outline)

    def test_para_id_present(self):
        outline = _j(server.get_document_outline())
        for h in outline:
            assert "para_id" in h
            assert isinstance(h["para_id"], str)

    def test_body_text_excluded(self):
        outline = _j(server.get_document_outline())
        texts = [h["text"] for h in outline]
        assert "Body text here." not in texts

    def test_default_max_level_6(self):
        """Default max_level=6 returns all headings."""
        outline = _j(server.get_document_outline())
        assert len(outline) == 4


# ═══════════════════════════════════════════════════════════════════════════
#  set_document_language
# ═══════════════════════════════════════════════════════════════════════════


class TestSetDocumentLanguage:
    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path):
        self.tmp_path = tmp_path
        path = tmp_path / "lang.docx"
        _build_lang_docx(path)
        server.open_document(str(path))

    def test_returns_language(self):
        result = _j(server.set_document_language("en-US"))
        assert result["language"] == "en-US"

    def test_sets_lang_val_in_styles(self):
        server.set_document_language("fr-FR")
        # Verify the in-memory tree was updated
        doc = server._doc
        styles = doc._tree("word/styles.xml")
        assert styles is not None
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        # Find default paragraph style
        default_style = None
        for s in styles.findall(f"{W}style"):
            if s.get(f"{W}type") == "paragraph" and s.get(f"{W}default") == "1":
                default_style = s
                break
        assert default_style is not None, "No default paragraph style found"
        rpr = default_style.find(f"{W}rPr")
        assert rpr is not None, "No rPr in default paragraph style"
        lang = rpr.find(f"{W}lang")
        assert lang is not None, "No w:lang in rPr"
        assert lang.get(f"{W}val") == "fr-FR"

    def test_marks_styles_dirty(self):
        server.set_document_language("de-DE")
        assert "word/styles.xml" in server._doc._modified

    def test_no_rpr_creates_it(self, tmp_path: Path):
        """Works even when rPr doesn't exist on default style."""
        path = tmp_path / "lang_norpr.docx"
        _build_lang_no_rpr_docx(path)
        server.open_document(str(path))
        result = _j(server.set_document_language("ja-JP"))
        assert result["language"] == "ja-JP"
        doc = server._doc
        styles = doc._tree("word/styles.xml")
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        for s in styles.findall(f"{W}style"):
            if s.get(f"{W}type") == "paragraph" and s.get(f"{W}default") == "1":
                rpr = s.find(f"{W}rPr")
                assert rpr is not None
                lang = rpr.find(f"{W}lang")
                assert lang is not None
                assert lang.get(f"{W}val") == "ja-JP"
                break


# ═══════════════════════════════════════════════════════════════════════════
#  set_track_changes
# ═══════════════════════════════════════════════════════════════════════════


class TestSetTrackChanges:
    @pytest.fixture(autouse=True)
    def _open(self, tmp_path: Path, test_docx: Path):
        server.open_document(str(test_docx))

    def test_enable_returns_true(self):
        result = _j(server.set_track_changes(enabled=True))
        assert result["track_changes"] is True

    def test_enable_sets_trackchanges_element(self):
        server.set_track_changes(enabled=True)
        doc = server._doc
        settings = doc._tree("word/settings.xml")
        assert settings is not None
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        tc = settings.find(f"{W}trackChanges")
        assert tc is not None, "w:trackChanges not found in settings.xml"

    def test_disable_removes_trackchanges_element(self):
        # Enable first
        server.set_track_changes(enabled=True)
        # Then disable
        server.set_track_changes(enabled=False)
        doc = server._doc
        settings = doc._tree("word/settings.xml")
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        tc = settings.find(f"{W}trackChanges")
        assert tc is None, "w:trackChanges should be absent after disabling"

    def test_disable_returns_false(self):
        server.set_track_changes(enabled=True)
        result = _j(server.set_track_changes(enabled=False))
        assert result["track_changes"] is False

    def test_author_in_response(self):
        result = _j(server.set_track_changes(enabled=True, author="Alice"))
        assert result["author"] == "Alice"

    def test_marks_settings_dirty(self):
        server.set_track_changes(enabled=True)
        assert "word/settings.xml" in server._doc._modified

    def test_disable_when_already_disabled_is_idempotent(self):
        """Disabling when not enabled should not raise."""
        result = _j(server.set_track_changes(enabled=False))
        assert result["track_changes"] is False

    def test_enable_twice_is_idempotent(self):
        """Enabling twice should not duplicate w:trackChanges."""
        server.set_track_changes(enabled=True)
        server.set_track_changes(enabled=True)
        doc = server._doc
        settings = doc._tree("word/settings.xml")
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        elements = settings.findall(f"{W}trackChanges")
        assert len(elements) == 1, "Should not have duplicate w:trackChanges"


def _build_no_settings_docx(path: Path) -> None:
    """DOCX that deliberately omits word/settings.xml."""
    _write_minimal_docx(path, _minimal_doc_xml())


class TestSetTrackChangesCreatesSettings:
    """set_track_changes must register settings.xml when it doesn't exist yet."""

    def test_set_track_changes_creates_settings_if_absent(self, tmp_path: Path):
        path = tmp_path / "no_settings.docx"
        _build_no_settings_docx(path)
        server.open_document(str(path))

        # Confirm settings.xml is absent before the call
        doc = server._doc
        assert doc._tree("word/settings.xml") is None

        server.set_track_changes(enabled=True)

        W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        settings = doc._tree("word/settings.xml")
        assert settings is not None, "settings.xml tree must be created"
        tc = settings.find(f"{W_NS}trackChanges")
        assert tc is not None, "w:trackChanges must be present after enabling"

    def test_content_type_registered_when_settings_created(self, tmp_path: Path):
        path = tmp_path / "no_settings_ct.docx"
        _build_no_settings_docx(path)
        server.open_document(str(path))

        server.set_track_changes(enabled=True)

        doc = server._doc
        CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
        ct = doc._tree("[Content_Types].xml")
        assert ct is not None
        parts = {e.get("PartName") for e in ct.findall(f"{CT_NS}Override")}
        assert "/word/settings.xml" in parts, "content-type Override must be registered"

    def test_relationship_registered_when_settings_created(self, tmp_path: Path):
        path = tmp_path / "no_settings_rels.docx"
        _build_no_settings_docx(path)
        server.open_document(str(path))

        server.set_track_changes(enabled=True)

        doc = server._doc
        RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rels = doc._tree("word/_rels/document.xml.rels")
        assert rels is not None
        targets = {r.get("Target") for r in rels.findall(f"{RELS_NS}Relationship")}
        assert "settings.xml" in targets, "relationship to settings.xml must be registered"
