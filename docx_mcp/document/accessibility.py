"""Accessibility mixin: alt text and accessibility check."""
from __future__ import annotations

from .base import WP, W


class AccessibilityMixin:
    """Provides image alt-text management and accessibility checking."""

    def set_alt_text(self, image_index: int, alt_text: str, title: str = "") -> dict:
        """Set the alt text (and optionally title) on an image by 0-based index.

        Args:
            image_index: 0-based index across all wp:docPr elements in document.xml.
            alt_text: Accessibility description to set on the image.
            title: Optional title attribute; removed from XML if empty.

        Returns:
            {"image_index": int, "alt_text": str}

        Raises:
            IndexError: If image_index is out of range.
        """
        doc = self._require("word/document.xml")
        doc_prs = list(doc.iter(f"{WP}docPr"))
        if image_index < 0 or image_index >= len(doc_prs):
            raise IndexError(f"Image index {image_index} out of range")
        el = doc_prs[image_index]
        el.set("descr", alt_text)
        if title:
            el.set("title", title)
        else:
            if "title" in el.attrib:
                del el.attrib["title"]
        self._mark("word/document.xml")
        return {"image_index": image_index, "alt_text": alt_text}

    def get_alt_text(self, image_index: int) -> dict:
        """Get the alt text and title for an image by 0-based index.

        Args:
            image_index: 0-based index across all wp:docPr elements in document.xml.

        Returns:
            {"image_index": int, "alt_text": str, "title": str}
            alt_text is "" if descr attribute is absent.
            title is "" if title attribute is absent.

        Raises:
            IndexError: If image_index is out of range.
        """
        doc = self._require("word/document.xml")
        doc_prs = list(doc.iter(f"{WP}docPr"))
        if image_index < 0 or image_index >= len(doc_prs):
            raise IndexError(f"Image index {image_index} out of range")
        el = doc_prs[image_index]
        return {
            "image_index": image_index,
            "alt_text": el.get("descr", ""),
            "title": el.get("title", ""),
        }

    def check_accessibility(self) -> dict:
        """Scan the document for accessibility issues.

        Checks:
        - Images missing alt text (wp:docPr with absent or empty descr)
        - Tables missing a header row (w:tbl where first w:tr lacks w:tblHeader)

        Returns:
            {"issue_count": int, "issues": list[dict]}

            Each issue dict has:
            - "type": "missing_alt_text" or "table_no_header"
            - "image_index" or "table_index": int
            - "description": str
        """
        doc = self._require("word/document.xml")
        issues: list[dict] = []

        # Check images for missing alt text
        for idx, el in enumerate(doc.iter(f"{WP}docPr")):
            descr = el.get("descr", "")
            if not descr:
                issues.append({
                    "type": "missing_alt_text",
                    "image_index": idx,
                    "description": f"Image {idx} has no alt text",
                })

        # Check tables for missing header row
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            rows = tbl.findall(f"{W}tr")
            if not rows:
                continue
            first_row = rows[0]
            trpr = first_row.find(f"{W}trPr")
            has_header = trpr is not None and trpr.find(f"{W}tblHeader") is not None
            if not has_header:
                issues.append({
                    "type": "table_no_header",
                    "table_index": idx,
                    "description": f"Table {idx} has no header row",
                })

        return {"issue_count": len(issues), "issues": issues}
