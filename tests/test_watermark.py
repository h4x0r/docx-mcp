"""Tests for WatermarkMixin: insert_watermark and remove_watermark."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument

V_NS = "urn:schemas-microsoft-com:vml"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _j(result: str) -> dict | list:
    return json.loads(result)


def _header_tree(doc: DocxDocument) -> etree._Element | None:
    for key, tree in doc._trees.items():
        if key.startswith("word/header"):
            return tree
    return None


def _find_vshapes(doc: DocxDocument) -> list[etree._Element]:
    shapes = []
    for key, tree in doc._trees.items():
        if key.startswith("word/header"):
            shapes.extend(tree.iter(f"{{{V_NS}}}shape"))
    return shapes


def _find_textpaths(doc: DocxDocument) -> list[etree._Element]:
    paths = []
    for key, tree in doc._trees.items():
        if key.startswith("word/header"):
            paths.extend(tree.iter(f"{{{V_NS}}}textpath"))
    return paths


class TestInsertWatermark:
    @pytest.fixture(autouse=True)
    def _open_fresh(self, tmp_path: Path):
        path = tmp_path / "wm_test.docx"
        doc = DocxDocument.create(str(path))
        doc.close()
        server.open_document(str(path))

    def test_insert_creates_vshape_in_header(self):
        server.insert_watermark("CONFIDENTIAL")
        doc = server._doc
        shapes = _find_vshapes(doc)
        assert len(shapes) >= 1

    def test_insert_vshape_has_correct_text(self):
        server.insert_watermark("DRAFT")
        doc = server._doc
        paths = _find_textpaths(doc)
        assert any(tp.get("string") == "DRAFT" for tp in paths)

    def test_insert_diagonal_true_no_rotation_zero(self):
        server.insert_watermark("DRAFT", diagonal=True)
        doc = server._doc
        shapes = _find_vshapes(doc)
        assert shapes, "Expected at least one v:shape"
        style = shapes[0].get("style", "")
        assert "mso-rotation:0" not in style

    def test_insert_diagonal_false_sets_rotation_zero(self):
        server.insert_watermark("DRAFT", diagonal=False)
        doc = server._doc
        shapes = _find_vshapes(doc)
        assert shapes, "Expected at least one v:shape"
        style = shapes[0].get("style", "")
        assert "mso-rotation:0" in style

    def test_insert_returns_correct_dict(self):
        result = _j(server.insert_watermark("SECRET"))
        assert result["header"] == "default"
        assert result["text"] == "SECRET"
        assert result["diagonal"] is True

    def test_insert_diagonal_false_returns_correct_dict(self):
        result = _j(server.insert_watermark("SECRET", diagonal=False))
        assert result["diagonal"] is False


class TestRemoveWatermark:
    @pytest.fixture(autouse=True)
    def _open_fixture(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_remove_after_insert_clears_watermark(self, tmp_path: Path):
        path = tmp_path / "fresh.docx"
        doc = DocxDocument.create(str(path))
        doc.close()
        server._doc.close()
        server._doc = DocxDocument.create(str(path))
        server.insert_watermark("DRAFT")
        result = _j(server.remove_watermark())
        assert result["removed"] >= 1
        paths = _find_textpaths(server._doc)
        assert len(paths) == 0

    def test_remove_returns_count_one_from_fixture(self):
        result = _j(server.remove_watermark())
        assert result["removed"] == 1

    def test_remove_idempotent_returns_zero(self):
        server.remove_watermark()
        result = _j(server.remove_watermark())
        assert result["removed"] == 0

    def test_remove_on_fresh_doc_returns_zero(self, tmp_path: Path):
        path = tmp_path / "fresh2.docx"
        doc = DocxDocument.create(str(path))
        doc.close()
        server._doc.close()
        server._doc = DocxDocument.create(str(path))
        result = _j(server.remove_watermark())
        assert result["removed"] == 0
