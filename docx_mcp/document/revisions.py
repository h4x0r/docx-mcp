"""Revision extraction mixin: read tracked changes as structured JSON."""

from __future__ import annotations

from lxml import etree

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

    def accept_change(self, change_id: int) -> dict:
        doc = self._require("word/document.xml")
        el = _find_change(doc, change_id)
        if el is None:
            raise ValueError(f"No tracked change with id={change_id}")
        if el.tag == f"{W}ins":
            _unwrap(el)
            return {"change_id": change_id, "action": "accepted", "type": "insertion"}
        else:
            el.getparent().remove(el)
            return {"change_id": change_id, "action": "accepted", "type": "deletion"}

    def reject_change(self, change_id: int) -> dict:
        doc = self._require("word/document.xml")
        el = _find_change(doc, change_id)
        if el is None:
            raise ValueError(f"No tracked change with id={change_id}")
        if el.tag == f"{W}ins":
            el.getparent().remove(el)
            return {"change_id": change_id, "action": "rejected", "type": "insertion"}
        else:
            _unwrap_del(el)
            return {"change_id": change_id, "action": "rejected", "type": "deletion"}

    def accept_all_changes(self) -> dict:
        doc = self._require("word/document.xml")
        elements = list(doc.iter(f"{W}ins")) + list(doc.iter(f"{W}del"))
        count = 0
        for el in elements:
            if el.tag == f"{W}ins":
                _unwrap(el)
            else:
                el.getparent().remove(el)
            count += 1
        return {"accepted": count}

    def reject_all_changes(self) -> dict:
        doc = self._require("word/document.xml")
        elements = list(doc.iter(f"{W}ins")) + list(doc.iter(f"{W}del"))
        count = 0
        for el in elements:
            if el.tag == f"{W}ins":
                el.getparent().remove(el)
            else:
                _unwrap_del(el)
            count += 1
        return {"rejected": count}


def _find_change(doc, change_id: int):
    id_str = str(change_id)
    for tag in (f"{W}ins", f"{W}del"):
        for el in doc.iter(tag):
            if el.get(f"{W}id") == id_str:
                return el
    return None


def _unwrap(el) -> None:
    parent = el.getparent()
    idx = list(parent).index(el)
    children = list(el)
    parent.remove(el)
    for i, child in enumerate(children):
        parent.insert(idx + i, child)


def _unwrap_del(el) -> None:
    for r in el.findall(f"{W}r"):
        for dt in r.findall(f"{W}delText"):
            t = etree.Element(f"{W}t")
            t.text = dt.text
            if dt.text and (dt.text.startswith(" ") or dt.text.endswith(" ")):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.remove(dt)
            r.append(t)
    _unwrap(el)


def _int_id(val: str | None) -> int:
    try:
        return int(val or 0)
    except (ValueError, TypeError):
        return 0
