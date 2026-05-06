"""
RED tests — Gap 6: PII scrubbing via Presidio.

scrub_pii(output_path, *, dry_run=False, entities=[], confidence_threshold=0.35,
          redaction_marker="█", also_sanitize_metadata=True) detects and
permanently redacts PII from the open document.

Requires: pip install presidio-analyzer presidio-anonymizer
          python -m spacy download en_core_web_lg

Behavior:
- dry_run=True  → returns detected entities as JSON, writes NO file.
- dry_run=False → redacts in-place in output DOCX, returns path + entity list.
- Redacted runs have w:highlight val="black" so they appear as black bars.
- All occurrences of a detected entity text are redacted (deduplication pass).
- output_path is required when dry_run=False; raises ValueError if empty.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server


def _j(s: str) -> dict:
    return json.loads(s)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _write_docx(path: Path, paragraphs: list[str]) -> None:
    """Minimal DOCX with one w:p per string."""
    paras_xml = ""
    for i, text in enumerate(paragraphs):
        paras_xml += (
            f'    <w:p w14:paraId="DD{i:06d}" w14:textId="77777777">\n'
            f'      <w:r><w:t xml:space="preserve">{text}</w:t></w:r>\n'
            f'    </w:p>\n'
        )
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document\n'
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '  <w:body>\n'
        + paras_xml
        + '  </w:body>\n</w:document>'
    )
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType='
        '"application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type='
        '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", doc_xml)


def _doc_text(path: Path) -> str:
    """Concatenate all w:t text from output DOCX."""
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    return "".join(t.text or "" for t in root.iter(_w("t")))


def _get_doc_root(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


# ─────────────────────────────────────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────────────────────────────────────


@pytest.fixture()
def email_docx(tmp_path: Path) -> Path:
    p = tmp_path / "email.docx"
    _write_docx(p, [
        "Please contact Alice at alice@acme-legal.com for further details.",
    ])
    return p


@pytest.fixture()
def person_docx(tmp_path: Path) -> Path:
    p = tmp_path / "person.docx"
    _write_docx(p, [
        "The defendant, Robert Johnson, appeared before the court.",
    ])
    return p


@pytest.fixture()
def phone_docx(tmp_path: Path) -> Path:
    p = tmp_path / "phone.docx"
    _write_docx(p, [
        "Call our office at (415) 555-0192 to schedule an appointment.",
    ])
    return p


@pytest.fixture()
def mixed_pii_docx(tmp_path: Path) -> Path:
    """Document with email + person name + surrounding non-PII context."""
    p = tmp_path / "mixed.docx"
    _write_docx(p, [
        "Plaintiff Sarah Connor filed suit against Cyberdyne Systems.",
        "Her email is sarah.connor@resistance.org and she is represented by counsel.",
        "The foregoing agreement is binding on all parties.",
    ])
    return p


@pytest.fixture()
def duplicate_pii_docx(tmp_path: Path) -> Path:
    """Same email address appears in two separate paragraphs."""
    p = tmp_path / "dup.docx"
    _write_docx(p, [
        "Send the brief to bob@lawfirm.io by Friday.",
        "All filings must also be copied to bob@lawfirm.io as per standing order.",
    ])
    return p


# ─────────────────────────────────────────────────────────────────────────────
# Detection tests (dry_run=True)
# ─────────────────────────────────────────────────────────────────────────────


class TestPiiDetection:

    def test_detects_email_address(self, email_docx: Path):
        """Presidio identifies EMAIL_ADDRESS in dry_run mode."""
        server.open_document(str(email_docx))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        types = [e["type"] for e in result["entities"]]
        assert "EMAIL_ADDRESS" in types

    def test_detected_email_text_is_correct(self, email_docx: Path):
        """The entity text matches the actual email in the document."""
        server.open_document(str(email_docx))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        email_entities = [e for e in result["entities"] if e["type"] == "EMAIL_ADDRESS"]
        assert any("alice@acme-legal.com" in e["text"] for e in email_entities)

    def test_detects_person_name(self, person_docx: Path):
        """Presidio NER identifies PERSON entities (requires spaCy model)."""
        server.open_document(str(person_docx))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        types = [e["type"] for e in result["entities"]]
        assert "PERSON" in types

    def test_detects_phone_number(self, phone_docx: Path):
        """Presidio identifies PHONE_NUMBER."""
        server.open_document(str(phone_docx))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        types = [e["type"] for e in result["entities"]]
        assert "PHONE_NUMBER" in types

    def test_dry_run_writes_no_file(self, email_docx: Path, tmp_path: Path):
        """dry_run=True must not create any output file."""
        server.open_document(str(email_docx))
        out = tmp_path / "should_not_exist.docx"
        result = _j(server.scrub_pii(output_path=str(out), dry_run=True))
        assert not out.exists()
        assert result["path"] is None

    def test_dry_run_entities_include_para_index(self, email_docx: Path):
        """Each entity dict includes a para_index field for location."""
        server.open_document(str(email_docx))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        for entity in result["entities"]:
            assert "para_index" in entity

    def test_no_pii_returns_empty_entities(self, tmp_path: Path):
        """Document with no PII returns empty entity list in dry_run."""
        p = tmp_path / "clean.docx"
        _write_docx(p, [
            "The foregoing terms and conditions are agreed upon by all parties.",
            "This agreement shall be governed by the laws of the jurisdiction.",
        ])
        server.open_document(str(p))
        result = _j(server.scrub_pii(output_path="", dry_run=True))
        assert result["entities"] == []


# ─────────────────────────────────────────────────────────────────────────────
# Redaction tests (dry_run=False)
# ─────────────────────────────────────────────────────────────────────────────


class TestPiiRedaction:

    def test_email_absent_from_output(self, email_docx: Path, tmp_path: Path):
        """After scrub, the email address is not present in the output document."""
        server.open_document(str(email_docx))
        out = tmp_path / "out.docx"
        server.scrub_pii(output_path=str(out))
        assert "alice@acme-legal.com" not in _doc_text(out)

    def test_non_pii_text_preserved(self, mixed_pii_docx: Path, tmp_path: Path):
        """Surrounding non-PII text is unchanged in the output."""
        server.open_document(str(mixed_pii_docx))
        out = tmp_path / "out.docx"
        server.scrub_pii(output_path=str(out))
        text = _doc_text(out)
        assert "The foregoing agreement is binding on all parties." in text

    def test_redacted_run_has_black_highlight(self, email_docx: Path, tmp_path: Path):
        """Runs containing redacted PII have w:highlight val='black' in their rPr."""
        server.open_document(str(email_docx))
        out = tmp_path / "out.docx"
        server.scrub_pii(output_path=str(out))
        root = _get_doc_root(out)
        highlights = [
            el.get(_w("val"))
            for el in root.iter(_w("highlight"))
        ]
        assert "black" in highlights

    def test_deduplication_redacts_all_occurrences(self, duplicate_pii_docx: Path, tmp_path: Path):
        """The same email appearing in two paragraphs is redacted in both."""
        server.open_document(str(duplicate_pii_docx))
        out = tmp_path / "out.docx"
        server.scrub_pii(output_path=str(out))
        text = _doc_text(out)
        assert "bob@lawfirm.io" not in text

    def test_output_path_required_when_not_dry_run(self, email_docx: Path):
        """scrub_pii with dry_run=False and empty output_path raises ValueError."""
        server.open_document(str(email_docx))
        with pytest.raises((ValueError, TypeError)):
            server.scrub_pii(output_path="", dry_run=False)

    def test_result_contains_output_path(self, email_docx: Path, tmp_path: Path):
        """JSON result includes path to the scrubbed output file."""
        server.open_document(str(email_docx))
        out = tmp_path / "out.docx"
        result = _j(server.scrub_pii(output_path=str(out)))
        assert "path" in result
        assert Path(result["path"]).exists()

    def test_result_contains_entity_list(self, email_docx: Path, tmp_path: Path):
        """JSON result includes the list of entities that were redacted."""
        server.open_document(str(email_docx))
        out = tmp_path / "out.docx"
        result = _j(server.scrub_pii(output_path=str(out)))
        assert "entities" in result
        assert len(result["entities"]) >= 1

    def test_entity_filter_limits_redaction(self, mixed_pii_docx: Path, tmp_path: Path):
        """When entities=['EMAIL_ADDRESS'], only emails are redacted (not names)."""
        server.open_document(str(mixed_pii_docx))
        out = tmp_path / "out.docx"
        server.scrub_pii(output_path=str(out), entities=["EMAIL_ADDRESS"])
        text = _doc_text(out)
        assert "sarah.connor@resistance.org" not in text
        # Person names should NOT be redacted when filtered to email only
        assert "Sarah Connor" in text
