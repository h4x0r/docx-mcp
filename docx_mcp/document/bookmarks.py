"""Bookmark CRUD mixin."""
from __future__ import annotations

import contextlib

from lxml import etree

from .base import W, W14
from .errors import DocxMcpError, ErrCode


class BookmarksMixin:
    """CRUD operations for Word bookmarks (w:bookmarkStart / w:bookmarkEnd pairs)."""

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def list_bookmarks(self) -> list[dict]:
        """Return all user bookmarks: [{"id": int, "name": str, "para_id": str|None}].

        Only returns bookmarks with matching start+end pairs.
        Skips Word-internal bookmarks (names starting with '_').
        """
        doc = self._require("word/document.xml")
        # Build lookup: id -> bookmarkEnd element
        end_ids: set[str] = {
            el.get(f"{W}id")
            for el in doc.iter(f"{W}bookmarkEnd")
            if el.get(f"{W}id") is not None
        }

        result: list[dict] = []
        for start in doc.iter(f"{W}bookmarkStart"):
            name = start.get(f"{W}name", "")
            if name.startswith("_"):
                continue
            bm_id_str = start.get(f"{W}id")
            if bm_id_str is None or bm_id_str not in end_ids:
                continue  # no matching end
            bm_id = 0
            with contextlib.suppress(ValueError):
                bm_id = int(bm_id_str)

            # Find parent paragraph for para_id
            para_id: str | None = None
            parent = start.getparent()
            if parent is not None:
                para_id = parent.get(f"{W14}paraId")

            result.append({"id": bm_id, "name": name, "para_id": para_id})

        return result

    def add_bookmark(self, para_id: str, name: str) -> dict:
        """Wrap paragraph in bookmark start/end.

        Places bookmarkStart as first child of the paragraph (before any runs),
        and bookmarkEnd as last child.

        Returns: {"id": int, "name": str, "para_id": str}
        Raises:
          DocxMcpError(PARA_NOT_FOUND) if para_id not found
          DocxMcpError(OOXML_INVALID) if name already exists or starts with '_'
        """
        if name.startswith("_"):
            raise DocxMcpError(
                ErrCode.OOXML_INVALID,
                f"Bookmark name '{name}' is reserved (starts with '_').",
                hint="Choose a name that does not start with an underscore.",
            )

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        # Check for duplicate name
        for start in doc.iter(f"{W}bookmarkStart"):
            if start.get(f"{W}name") == name:
                raise DocxMcpError(
                    ErrCode.OOXML_INVALID,
                    f"Bookmark '{name}' already exists.",
                    hint="Use a unique bookmark name.",
                )

        # Locate paragraph
        para = self._find_para(body, para_id)
        if para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                f"Paragraph with paraId '{para_id}' not found.",
            )

        # Next unique ID
        bm_id = self._next_bookmark_id(doc)

        # Create elements
        bm_start = etree.Element(f"{W}bookmarkStart")
        bm_start.set(f"{W}id", str(bm_id))
        bm_start.set(f"{W}name", name)

        bm_end = etree.Element(f"{W}bookmarkEnd")
        bm_end.set(f"{W}id", str(bm_id))

        # Insert at position 0 (before any runs), append at end
        para.insert(0, bm_start)
        para.append(bm_end)

        self._mark("word/document.xml")
        return {"id": bm_id, "name": name, "para_id": para_id}

    def remove_bookmark(self, name: str) -> dict:
        """Remove bookmarkStart + bookmarkEnd pair by name.

        Returns: {"removed": name}
        Raises: DocxMcpError(BOOKMARK_NOT_FOUND) if name not found
        """
        doc = self._require("word/document.xml")

        # Find bookmarkStart
        bm_start = None
        for el in doc.iter(f"{W}bookmarkStart"):
            if el.get(f"{W}name") == name:
                bm_start = el
                break

        if bm_start is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND,
                f"Bookmark '{name}' not found.",
            )

        bm_id_str = bm_start.get(f"{W}id")

        # Remove bookmarkStart
        parent = bm_start.getparent()
        if parent is not None:
            parent.remove(bm_start)

        # Remove matching bookmarkEnd
        if bm_id_str is not None:
            for el in doc.iter(f"{W}bookmarkEnd"):
                if el.get(f"{W}id") == bm_id_str:
                    end_parent = el.getparent()
                    if end_parent is not None:
                        end_parent.remove(el)
                    break

        self._mark("word/document.xml")
        return {"removed": name}

    def get_bookmarked_text(self, name: str) -> dict:
        """Return text in the paragraph containing the named bookmark.

        Returns: {"name": str, "text": str}
        Raises: DocxMcpError(BOOKMARK_NOT_FOUND) if name not found
        """
        doc = self._require("word/document.xml")

        bm_start = None
        for el in doc.iter(f"{W}bookmarkStart"):
            if el.get(f"{W}name") == name:
                bm_start = el
                break

        if bm_start is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND,
                f"Bookmark '{name}' not found.",
            )

        # Collect w:t text from the parent paragraph
        parent = bm_start.getparent()
        text = self._text(parent) if parent is not None else ""
        return {"name": name, "text": text}

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _next_bookmark_id(doc: etree._Element) -> int:
        """Return next unique bookmark ID (max of existing + 1)."""
        max_id = -1
        for tag in (f"{W}bookmarkStart", f"{W}bookmarkEnd"):
            for el in doc.iter(tag):
                id_str = el.get(f"{W}id")
                if id_str is not None:
                    with contextlib.suppress(ValueError):
                        max_id = max(max_id, int(id_str))
        return max_id + 1
