"""Revision extraction mixin: read tracked changes as structured JSON."""

from __future__ import annotations

from .base import W, W14


class RevisionsMixin:
    """Extract pending tracked changes (w:ins / w:del) as structured data."""

    def get_tracked_changes(self) -> list[dict]:
        """Return all pending tracked changes in document order.

        Each entry:
          {
            "type":      "insertion" | "deletion",
            "change_id": int,
            "author":    str,
            "date":      str,
            "para_id":   str | None,
            "text":      str
          }
        """
        doc = self._require("word/document.xml")
        results: list[dict] = []

        for para in doc.iter(f"{W}p"):
            para_id = para.get(f"{W14}paraId")

            for child in para:
                tag = child.tag
                if tag == f"{W}ins":
                    text = "".join(
                        t.text or ""
                        for t in child.iter(f"{W}t")
                    )
                    results.append({
                        "type": "insertion",
                        "change_id": _int_id(child.get(f"{W}id")),
                        "author": child.get(f"{W}author", ""),
                        "date": child.get(f"{W}date", ""),
                        "para_id": para_id,
                        "text": text,
                    })
                elif tag == f"{W}del":
                    text = "".join(
                        dt.text or ""
                        for dt in child.iter(f"{W}delText")
                    )
                    results.append({
                        "type": "deletion",
                        "change_id": _int_id(child.get(f"{W}id")),
                        "author": child.get(f"{W}author", ""),
                        "date": child.get(f"{W}date", ""),
                        "para_id": para_id,
                        "text": text,
                    })

        return results


def _int_id(val: str | None) -> int:
    try:
        return int(val or 0)
    except (ValueError, TypeError):
        return 0
