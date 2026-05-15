"""Tests for Comment CRUD methods: update_comment, delete_comment,
resolve_comment, list_comment_threads."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


class TestCommentCRUD:
    """Tests for CommentsMixin CRUD extensions."""

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _open(self, path: Path) -> None:
        server._doc = None
        server.open_document(str(path))

    def _add_comment(self, para_id: str, text: str) -> int:
        result = json.loads(server.add_comment(para_id, text))
        return result["comment_id"]

    def _add_reply(self, parent_id: int, text: str) -> int:
        result = json.loads(server.reply_to_comment(parent_id, text))
        return result["comment_id"]

    # ── update_comment ────────────────────────────────────────────────────────

    def test_update_comment_text(self, test_docx: Path) -> None:
        self._open(test_docx)
        cid = self._add_comment("00000002", "original text")

        result = json.loads(server.update_comment(cid, "updated text"))

        assert result["comment_id"] == cid
        assert result["text"] == "updated text"

        # Verify via get_comments
        comments = json.loads(server.get_comments())
        match = next(c for c in comments if c["id"] == cid)
        assert match["text"] == "updated text"

    def test_update_comment_not_found(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.update_comment(9999, "text")

    # ── delete_comment ────────────────────────────────────────────────────────

    def test_delete_comment_removes_from_xml(self, test_docx: Path) -> None:
        self._open(test_docx)
        cid = self._add_comment("00000002", "to be deleted")

        result = json.loads(server.delete_comment(cid))

        assert result["deleted"] == cid

        comments = json.loads(server.get_comments())
        ids = [c["id"] for c in comments]
        assert cid not in ids

    def test_delete_comment_removes_range_markers(self, test_docx: Path) -> None:

        self._open(test_docx)
        cid = self._add_comment("00000002", "range marker test")
        server.delete_comment(cid)

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        doc_tree = doc._tree("word/document.xml")

        cid_str = str(cid)
        for tag in (f"{W}commentRangeStart", f"{W}commentRangeEnd", f"{W}commentReference"):
            for el in doc_tree.iter(tag):
                assert el.get(f"{W}id") != cid_str, f"{tag} with id={cid_str} still present"

    def test_delete_comment_not_found(self, test_docx: Path) -> None:
        self._open(test_docx)
        # First add a comment so comments.xml exists
        self._add_comment("00000002", "seed")
        with pytest.raises(ValueError, match="not found"):
            server._doc.delete_comment(9999)

    # ── resolve_comment ───────────────────────────────────────────────────────

    def test_resolve_comment_sets_done(self, test_docx: Path) -> None:
        """Resolving a comment that has a reply (which creates commentsExtended.xml)."""

        self._open(test_docx)
        cid = self._add_comment("00000002", "parent comment")
        # reply_to_comment creates commentsExtended.xml
        self._add_reply(cid, "a reply")

        result = json.loads(server.resolve_comment(cid))

        assert result["resolved"] == cid
        assert result["found_extended"] is True

        # Verify done="1" is set in commentsExtended.xml
        W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"

        doc = server._doc
        ext = doc._tree("word/commentsExtended.xml")
        assert ext is not None

        cm_tree = doc._tree("word/comments.xml")
        # Get the paraId of parent comment's first paragraph
        parent_el = next(
            c
            for c in cm_tree.findall(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment"
            )
            if c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id") == str(cid)
        )
        parent_para = parent_el.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
        )
        parent_para_id = parent_para.get(f"{W14}paraId")

        found = False
        for ce in ext.iter(f"{W15}commentEx"):
            if ce.get(f"{W15}paraId") == parent_para_id:
                assert ce.get(f"{W15}done") == "1"
                found = True
        assert found, "No commentEx found for parent paraId"

    def test_resolve_comment_no_extended(self, test_docx: Path) -> None:
        """Resolving a comment when commentsExtended.xml does not exist."""
        self._open(test_docx)
        cid = self._add_comment("00000002", "standalone comment")
        # No reply → no commentsExtended.xml

        result = json.loads(server.resolve_comment(cid))

        assert result["resolved"] == cid
        assert result["found_extended"] is False

    # ── list_comment_threads ──────────────────────────────────────────────────

    def test_list_comment_threads_empty(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.list_comment_threads())
        assert result == []

    def test_list_comment_threads_with_reply(self, test_docx: Path) -> None:
        self._open(test_docx)
        root_id = self._add_comment("00000002", "root comment")
        reply_id = self._add_reply(root_id, "reply text")

        threads = json.loads(server.list_comment_threads())

        assert len(threads) == 1
        thread = threads[0]
        assert thread["root"]["id"] == root_id
        assert thread["root"]["text"] == "root comment"
        assert len(thread["replies"]) == 1
        assert thread["replies"][0]["id"] == reply_id
        assert thread["replies"][0]["text"] == "reply text"

    def test_list_comment_threads_multiple_roots(self, test_docx: Path) -> None:
        """Two independent comments form two separate root threads."""
        self._open(test_docx)
        id1 = self._add_comment("00000002", "first root")
        id2 = self._add_comment("00000003", "second root")

        threads = json.loads(server.list_comment_threads())

        assert len(threads) == 2
        root_ids = {t["root"]["id"] for t in threads}
        assert root_ids == {id1, id2}
        for t in threads:
            assert t["replies"] == []
