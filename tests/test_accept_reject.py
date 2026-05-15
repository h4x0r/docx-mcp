"""RED tests — Task #33: accept_change, reject_change, accept_all_changes, reject_all_changes."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(s: str) -> list | dict:
    return json.loads(s)


# ── helpers ──────────────────────────────────────────────────────────────────

_CT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'  # noqa: E501
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'  # noqa: E501
    "</Types>"
)

_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="word/document.xml"/>'
    "</Relationships>"
)


def _write_docx(path: Path, doc_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("word/document.xml", doc_xml)


def _make_doc(tmp_path: Path, doc_xml: str) -> Path:
    p = tmp_path / "doc.docx"
    _write_docx(p, doc_xml)
    return p


_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)

_DOC_ONE_INS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {_NS}>
  <w:body>
    <w:p w14:paraId="BB000001">
      <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
      <w:ins w:id="10" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>world</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> end.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

_DOC_ONE_DEL = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {_NS}>
  <w:body>
    <w:p w14:paraId="BB000002">
      <w:r><w:t xml:space="preserve">Keep </w:t></w:r>
      <w:del w:id="20" w:author="Bob" w:date="2026-01-02T00:00:00Z">
        <w:r><w:delText>gone</w:delText></w:r>
      </w:del>
      <w:r><w:t xml:space="preserve"> rest.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

_DOC_BOTH = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {_NS}>
  <w:body>
    <w:p w14:paraId="CC000001">
      <w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>inserted</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="CC000002">
      <w:del w:id="2" w:author="Bob" w:date="2026-01-02T00:00:00Z">
        <w:r><w:delText>deleted</w:delText></w:r>
      </w:del>
    </w:p>
  </w:body>
</w:document>"""


# ── tests ────────────────────────────────────────────────────────────────────


class TestAcceptChange:
    def test_accept_ins_moves_runs_to_parent(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        server.accept_change(10)
        changes = _j(server.get_tracked_changes())
        assert not any(c["change_id"] == 10 for c in changes)

    def test_accept_ins_returns_correct_dict(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        result = _j(server.accept_change(10))
        assert result == {"change_id": 10, "action": "accepted", "type": "insertion"}

    def test_accept_del_removes_element(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_DEL)
        server.open_document(str(p))
        server.accept_change(20)
        changes = _j(server.get_tracked_changes())
        assert not any(c["change_id"] == 20 for c in changes)

    def test_accept_del_returns_correct_dict(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_DEL)
        server.open_document(str(p))
        result = _j(server.accept_change(20))
        assert result == {"change_id": 20, "action": "accepted", "type": "deletion"}

    def test_accept_unknown_id_raises_value_error(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        with pytest.raises(ValueError):
            server.accept_change(9999)

    def test_accept_ins_text_preserved_in_parent(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        server.accept_change(10)
        content = _j(server.get_paragraph("BB000001"))
        assert "world" in content["text"]


class TestRejectChange:
    def test_reject_ins_removes_element(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        server.reject_change(10)
        changes = _j(server.get_tracked_changes())
        assert not any(c["change_id"] == 10 for c in changes)

    def test_reject_ins_returns_correct_dict(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_INS)
        server.open_document(str(p))
        result = _j(server.reject_change(10))
        assert result == {"change_id": 10, "action": "rejected", "type": "insertion"}

    def test_reject_del_restores_text(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_DEL)
        server.open_document(str(p))
        server.reject_change(20)
        content = _j(server.get_paragraph("BB000002"))
        assert "gone" in content["text"]

    def test_reject_del_returns_correct_dict(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_DEL)
        server.open_document(str(p))
        result = _j(server.reject_change(20))
        assert result == {"change_id": 20, "action": "rejected", "type": "deletion"}

    def test_reject_unknown_id_raises_exception(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_ONE_DEL)
        server.open_document(str(p))
        with pytest.raises(ValueError):
            server.reject_change(9999)


class TestAcceptAllChanges:
    def test_accept_all_returns_count(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_BOTH)
        server.open_document(str(p))
        result = _j(server.accept_all_changes())
        assert result == {"accepted": 2}

    def test_accept_all_removes_all_tracked_changes(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_BOTH)
        server.open_document(str(p))
        server.accept_all_changes()
        changes = _j(server.get_tracked_changes())
        assert changes == []


class TestRejectAllChanges:
    def test_reject_all_returns_count(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_BOTH)
        server.open_document(str(p))
        result = _j(server.reject_all_changes())
        assert result == {"rejected": 2}

    def test_reject_all_removes_all_tracked_changes(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_BOTH)
        server.open_document(str(p))
        server.reject_all_changes()
        changes = _j(server.get_tracked_changes())
        assert changes == []

    def test_reject_all_restores_deleted_text(self, tmp_path: Path):
        p = _make_doc(tmp_path, _DOC_BOTH)
        server.open_document(str(p))
        server.reject_all_changes()
        content = _j(server.get_paragraph("CC000002"))
        assert "deleted" in content["text"]
