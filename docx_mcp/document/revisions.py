"""Revision extraction mixin: read tracked changes as structured JSON."""

from __future__ import annotations

import contextlib

from lxml import etree

from .base import CT, RELS, W, W14


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

    def set_track_changes(self, enabled: bool, author: str = "") -> dict:
        """Enable or disable revision tracking in word/settings.xml.

        Args:
            enabled: True to enable tracking, False to disable.
            author: Author name for reference (returned in response only).

        Returns:
            {"track_changes": bool, "author": str}
        """
        # Get or create settings tree
        settings = self._tree("word/settings.xml")
        if settings is None:
            settings = etree.Element(
                f"{W}settings",
                nsmap={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
            )
            self._trees["word/settings.xml"] = settings

            # Write to disk so workdir is consistent
            fp = self.workdir / "word" / "settings.xml"
            fp.parent.mkdir(parents=True, exist_ok=True)
            etree.ElementTree(settings).write(
                str(fp), xml_declaration=True, encoding="UTF-8"
            )

            # Register content-type Override
            ct = self._tree("[Content_Types].xml")
            if ct is not None:
                existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
                if "/word/settings.xml" not in existing:
                    ov = etree.SubElement(ct, f"{CT}Override")
                    ov.set("PartName", "/word/settings.xml")
                    ov.set(
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.settings+xml",
                    )
                    self._mark("[Content_Types].xml")

            # Register relationship entry
            rels = self._tree("word/_rels/document.xml.rels")
            if rels is not None:
                existing_targets = {
                    r.get("Target") for r in rels.findall(f"{RELS}Relationship")
                }
                if "settings.xml" not in existing_targets:
                    max_rid = 0
                    for r in rels.findall(f"{RELS}Relationship"):
                        rid = r.get("Id", "")
                        if rid.startswith("rId"):
                            with contextlib.suppress(ValueError):
                                max_rid = max(max_rid, int(rid[3:]))
                    rel = etree.SubElement(rels, f"{RELS}Relationship")
                    rel.set("Id", f"rId{max_rid + 1}")
                    rel.set(
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument"
                        "/2006/relationships/settings",
                    )
                    rel.set("Target", "settings.xml")
                    self._mark("word/_rels/document.xml.rels")

        tc_tag = f"{W}trackChanges"
        existing = settings.find(tc_tag)

        if enabled:
            if existing is None:
                etree.SubElement(settings, tc_tag)
        else:
            if existing is not None:
                settings.remove(existing)

        self._mark("word/settings.xml")
        return {"track_changes": enabled, "author": author}

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

    def flatten_document(self) -> dict:
        """Accept all tracked changes and strip all revision markup.

        Accepts every w:ins / w:del element (via accept_all_changes) and then
        removes all w:rPrChange and w:pPrChange elements so no revision markup
        remains in the document.

        Returns:
            {"changes_accepted": int, "formatting_changes_removed": int}
        """
        accepted_result = self.accept_all_changes()
        changes_accepted: int = accepted_result["accepted"]

        doc = self._require("word/document.xml")
        fmt_count = 0
        for tag in (f"{W}rPrChange", f"{W}pPrChange"):
            for el in list(doc.iter(tag)):
                el.getparent().remove(el)
                fmt_count += 1

        if fmt_count:
            self._mark("word/document.xml")

        return {
            "changes_accepted": changes_accepted,
            "formatting_changes_removed": fmt_count,
        }

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
            if dt.text and any(c.isspace() for c in dt.text):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.remove(dt)
            r.append(t)
    _unwrap(el)


def _int_id(val: str | None) -> int:
    try:
        return int(val or 0)
    except (ValueError, TypeError):
        return 0
