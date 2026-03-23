"""Tests for Phase 4 content creation tools."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  add_list
# ═══════════════════════════════════════════════════════════════════════════


class TestAddList:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_bullet_list(self):
        result = _j(server.add_list(["00000004", "00000005"], style="bullet"))
        assert result["paragraphs_updated"] == 2
        assert result["list_id"] >= 1

    def test_numbered_list(self):
        result = _j(server.add_list(["00000004"], style="numbered"))
        assert result["paragraphs_updated"] == 1

    def test_bad_para_id(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_list(["DEADBEEF"])

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_list(["00000004"])


# ═══════════════════════════════════════════════════════════════════════════
#  insert_image
# ═══════════════════════════════════════════════════════════════════════════


class TestInsertImage:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_insert_image(self, tmp_path: Path):
        # Create a tiny PNG
        img = tmp_path / "test.png"
        img.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
            b"\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02"
            b"\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx"
            b"\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N"
            b"\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        result = _j(server.insert_image("00000004", str(img)))
        assert "rId" in result
        assert result["filename"].endswith(".png")
        # Verify image count increased
        images = _j(server.get_images())
        assert len(images) == 2  # original + new

    def test_insert_image_new_content_type(self, tmp_path: Path):
        """Insert a JPEG — content type not in fixture, triggers CT addition."""
        img = tmp_path / "test.jpg"
        img.write_bytes(b"\xff\xd8\xff\xe0")  # minimal JPEG header
        result = _j(server.insert_image("00000004", str(img)))
        assert result["filename"].endswith(".jpg")

    def test_insert_image_bad_para(self, tmp_path: Path):
        img = tmp_path / "test.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n")
        with pytest.raises(ValueError, match="not found"):
            server.insert_image("DEADBEEF", str(img))

    def test_insert_image_no_document(self, tmp_path: Path):
        server.close_document()
        img = tmp_path / "test.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n")
        with pytest.raises(RuntimeError, match="No document"):
            server.insert_image("00000004", str(img))


# ═══════════════════════════════════════════════════════════════════════════
#  edit_header_footer
# ═══════════════════════════════════════════════════════════════════════════


class TestEditHeaderFooter:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_edit_header(self):
        result = _j(server.edit_header_footer("header", "Document Header Text", "New Header"))
        assert result["changes"] >= 1
        assert result["location"] == "header"

    def test_edit_header_substring(self):
        """Edit a substring — triggers text-before and text-after branches."""
        result = _j(server.edit_header_footer("header", "Header", "Footer"))
        assert result["changes"] >= 1

    def test_edit_header_with_rpr(self):
        """Edit header run that has rPr — preserves formatting in tracked changes."""
        from lxml import etree

        from docx_mcp.document import W

        tree = server._doc._trees["word/header1.xml"]
        # Find the text run and add rPr
        for r in tree.iter(f"{W}r"):
            t = r.find(f"{W}t")
            if t is not None and t.text and "Document" in t.text:
                rpr = etree.SubElement(r, f"{W}rPr")
                etree.SubElement(rpr, f"{W}b")
                r.remove(rpr)
                r.insert(0, rpr)
                break
        result = _j(server.edit_header_footer("header", "Document", "Updated"))
        assert result["changes"] >= 1

    def test_text_not_found(self):
        with pytest.raises(ValueError, match="not found"):
            server.edit_header_footer("header", "NONEXISTENT", "replacement")

    def test_bad_location(self):
        with pytest.raises(ValueError, match="No footer"):
            server.edit_header_footer("footer", "text", "new")

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.edit_header_footer("header", "old", "new")


# ═══════════════════════════════════════════════════════════════════════════
#  add_endnote / validate_endnotes
# ═══════════════════════════════════════════════════════════════════════════


class TestAddEndnote:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_add_endnote(self):
        result = _j(server.add_endnote("00000004", "New endnote text"))
        assert result["endnote_id"] >= 1
        assert result["text"] == "New endnote text"
        # Verify endnote shows up in list
        endnotes = _j(server.get_endnotes())
        assert any("New endnote text" in e["text"] for e in endnotes)

    def test_add_endnote_bootstrap(self, tmp_path: Path):
        """Add endnote to document without endnotes.xml — bootstraps the file."""
        path = tmp_path / "noen.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        result = _j(server.add_endnote("00000001", "Bootstrapped endnote"))
        assert result["endnote_id"] == 1
        endnotes = _j(server.get_endnotes())
        assert len(endnotes) == 1

    def test_add_endnote_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_endnote("DEADBEEF", "text")

    def test_add_endnote_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_endnote("00000004", "text")


class TestValidateEndnotes:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_valid(self):
        result = _j(server.validate_endnotes())
        assert result["valid"] is True

    def test_after_add(self):
        server.add_endnote("00000004", "New endnote")
        result = _j(server.validate_endnotes())
        assert result["valid"] is True

    def test_no_endnotes_xml(self, tmp_path: Path):
        """Document without endnotes.xml — valid with 0 total."""
        path = tmp_path / "noen.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        result = _j(server.validate_endnotes())
        assert result["valid"] is True
        assert result["total"] == 0

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.validate_endnotes()
