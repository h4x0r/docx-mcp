"""Hyperlink CRUD mixin."""
from __future__ import annotations

import contextlib

from lxml import etree

from .base import RELS, W14, R, W
from .errors import DocxMcpError, ErrCode

_HYPERLINK_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
_RELS_PATH = "word/_rels/document.xml.rels"


class HyperlinksMixin:
    """Hyperlink CRUD operations."""

    # ── List ─────────────────────────────────────────────────────────────────

    def list_hyperlinks(self) -> list[dict]:
        """Return all hyperlinks in the document.

        Returns a list of dicts:
          {"id": str|None, "url_or_anchor": str, "text": str,
           "para_id": str|None, "type": "external"|"internal"}
        """
        doc = self._tree("word/document.xml")
        if doc is None:
            return []
        rels = self._tree(_RELS_PATH)

        results: list[dict] = []
        for hl in doc.iter(f"{W}hyperlink"):
            r_id = hl.get(f"{R}id")
            anchor = hl.get(f"{W}anchor")
            text = "".join(t.text for t in hl.iter(f"{W}t") if t.text)

            # Parent para_id
            para = hl.getparent()
            while para is not None and para.tag != f"{W}p":
                para = para.getparent()
            para_id = para.get(f"{W14}paraId") if para is not None else None

            if r_id is not None:
                # External — look up URL from rels
                url = ""
                if rels is not None:
                    rel = rels.find(f'{{{RELS[1:-1]}}}Relationship[@Id="{r_id}"]')
                    if rel is not None:
                        url = rel.get("Target", "")
                results.append(
                    {
                        "id": r_id,
                        "url_or_anchor": url,
                        "text": text,
                        "para_id": para_id,
                        "type": "external",
                    }
                )
            elif anchor is not None:
                results.append(
                    {
                        "id": None,
                        "url_or_anchor": anchor,
                        "text": text,
                        "para_id": para_id,
                        "type": "internal",
                    }
                )

        return results

    # ── Add external hyperlink ────────────────────────────────────────────────

    def add_hyperlink(self, para_id: str, text: str, url: str) -> dict:
        """Append an external hyperlink run at the end of the paragraph.

        Creates an r:id relationship in word/_rels/document.xml.rels.
        Returns: {"r_id": str, "para_id": str, "url": str}
        Raises: DocxMcpError(PARA_NOT_FOUND) if para_id not found.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, f"Paragraph not found: {para_id}")

        # Generate next rId
        rels = self._require(_RELS_PATH)
        rid = _next_rid(rels)

        # Add relationship
        rel = etree.SubElement(rels, f"{{{RELS[1:-1]}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", _HYPERLINK_TYPE)
        rel.set("Target", url)
        rel.set("TargetMode", "External")
        self._mark(_RELS_PATH)

        # Build hyperlink element
        hl = _make_hyperlink_external(rid, text)
        para.append(hl)
        self._mark("word/document.xml")

        return {"r_id": rid, "para_id": para_id, "url": url}

    # ── Add internal hyperlink ────────────────────────────────────────────────

    def add_internal_link(self, para_id: str, text: str, bookmark: str) -> dict:
        """Append an internal hyperlink (w:anchor) at the end of the paragraph.

        Does NOT add a relationship — internal links use w:anchor directly.
        Returns: {"anchor": str, "para_id": str}
        Raises: DocxMcpError(PARA_NOT_FOUND) if para_id not found.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, f"Paragraph not found: {para_id}")

        hl = _make_hyperlink_internal(bookmark, text)
        para.append(hl)
        self._mark("word/document.xml")

        return {"anchor": bookmark, "para_id": para_id}

    # ── Remove hyperlink ──────────────────────────────────────────────────────

    def remove_hyperlink(self, para_id: str, url_or_anchor: str) -> dict:
        """Unwrap hyperlink — keeps text runs, removes w:hyperlink wrapper.

        Returns: {"removed": url_or_anchor}
        Raises: DocxMcpError(BOOKMARK_NOT_FOUND) if not found.
        """
        doc = self._require("word/document.xml")
        rels = self._tree(_RELS_PATH)

        para = self._find_para(doc, para_id)
        if para is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND, f"Hyperlink not found: {url_or_anchor}"
            )

        target_hl = _find_hyperlink(para, url_or_anchor, rels)
        if target_hl is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND, f"Hyperlink not found: {url_or_anchor}"
            )

        # Move child runs to parent before the hyperlink element, then remove it
        parent = target_hl.getparent()
        idx = list(parent).index(target_hl)
        for child in list(target_hl):
            parent.insert(idx, child)
            idx += 1
        parent.remove(target_hl)
        self._mark("word/document.xml")

        return {"removed": url_or_anchor}

    # ── Update hyperlink ──────────────────────────────────────────────────────

    def update_hyperlink(self, r_id: str, new_url: str) -> dict:
        """Update the Target URL for an existing hyperlink relationship.

        Returns: {"r_id": str, "new_url": str}
        Raises: DocxMcpError(BOOKMARK_NOT_FOUND) if r_id not found.
        """
        rels = self._require(_RELS_PATH)
        rel = rels.find(f'{{{RELS[1:-1]}}}Relationship[@Id="{r_id}"]')
        if rel is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND, f"Relationship not found: {r_id}"
            )
        rel.set("Target", new_url)
        self._mark(_RELS_PATH)
        return {"r_id": r_id, "new_url": new_url}


# ── Module-level helpers ──────────────────────────────────────────────────────

def _next_rid(rels: etree._Element) -> str:
    """Return the next available rIdN by scanning existing Relationship/@Id values."""
    rels_ns = RELS[1:-1]  # strip braces
    max_n = 0
    for rel in rels.findall(f"{{{rels_ns}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            with contextlib.suppress(ValueError):
                max_n = max(max_n, int(rid[3:]))
    return f"rId{max_n + 1}"


def _make_hyperlink_external(rid: str, text: str) -> etree._Element:
    """Build a <w:hyperlink r:id="..."> element with a styled run."""
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    hl = etree.Element(f"{W}hyperlink", nsmap={"r": ns_r, "w": ns_w})
    hl.set(f"{R}id", rid)
    run = etree.SubElement(hl, f"{W}r")
    rpr = etree.SubElement(run, f"{W}rPr")
    rstyle = etree.SubElement(rpr, f"{W}rStyle")
    rstyle.set(f"{W}val", "Hyperlink")
    t = etree.SubElement(run, f"{W}t")
    t.text = text
    return hl


def _make_hyperlink_internal(anchor: str, text: str) -> etree._Element:
    """Build a <w:hyperlink w:anchor="..."> element with a run."""
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    hl = etree.Element(f"{W}hyperlink", nsmap={"w": ns_w})
    hl.set(f"{W}anchor", anchor)
    run = etree.SubElement(hl, f"{W}r")
    t = etree.SubElement(run, f"{W}t")
    t.text = text
    return hl


def _find_hyperlink(
    para: etree._Element,
    url_or_anchor: str,
    rels: etree._Element | None,
) -> etree._Element | None:
    """Find a w:hyperlink within para matching url_or_anchor."""
    rels_ns = RELS[1:-1]
    for hl in para.iter(f"{W}hyperlink"):
        r_id = hl.get(f"{R}id")
        if r_id is not None and rels is not None:
            rel = rels.find(f'{{{rels_ns}}}Relationship[@Id="{r_id}"]')
            if rel is not None and rel.get("Target") == url_or_anchor:
                return hl
        anchor = hl.get(f"{W}anchor")
        if anchor == url_or_anchor:
            return hl
    return None
