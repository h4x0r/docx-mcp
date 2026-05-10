"""Footnotes mixin: get, add, validate."""

from __future__ import annotations

from lxml import etree

from .base import W14, W, _preserve


class FootnotesMixin:
    """Footnote operations."""

    def get_footnotes(self) -> list[dict]:
        fn_tree = self._tree("word/footnotes.xml")
        if fn_tree is None:
            return []
        result = []
        for fn in self._real_footnotes(fn_tree):
            result.append(
                {
                    "id": int(fn.get(f"{W}id", "0")),
                    "text": self._text(fn),
                }
            )
        return result

    def add_footnote(self, para_id: str, text: str) -> dict:
        """Add a footnote to a paragraph. Returns the new footnote ID."""
        doc = self._require("word/document.xml")
        fn_tree = self._require("word/footnotes.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # Next ID
        existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        next_id = max(existing | {0}) + 1

        # Build footnote in footnotes.xml
        fn_el = etree.SubElement(fn_tree, f"{W}footnote")
        fn_el.set(f"{W}id", str(next_id))

        fn_para = etree.SubElement(fn_el, f"{W}p")
        fn_para.set(f"{W14}paraId", self._new_para_id())
        fn_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(fn_para, f"{W}pPr")
        ps = etree.SubElement(ppr, f"{W}pStyle")
        ps.set(f"{W}val", "FootnoteText")

        # Footnote ref mark
        ref_run = etree.SubElement(fn_para, f"{W}r")
        ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
        ref_style = etree.SubElement(ref_rpr, f"{W}rStyle")
        ref_style.set(f"{W}val", "FootnoteReference")
        etree.SubElement(ref_run, f"{W}footnoteRef")

        # Space
        sp_run = etree.SubElement(fn_para, f"{W}r")
        sp_t = etree.SubElement(sp_run, f"{W}t")
        _preserve(sp_t, " ")

        # Text
        txt_run = etree.SubElement(fn_para, f"{W}r")
        txt_t = etree.SubElement(txt_run, f"{W}t")
        _preserve(txt_t, text)

        self._mark("word/footnotes.xml")

        # Add reference in document paragraph
        r = etree.SubElement(para, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "FootnoteReference")
        fref = etree.SubElement(r, f"{W}footnoteReference")
        fref.set(f"{W}id", str(next_id))
        self._mark("word/document.xml")

        return {"footnote_id": next_id, "para_id": para_id}

    def update_footnote(self, footnote_id: int, text: str) -> dict:
        """Update the text of an existing footnote.

        Replaces the content of the first non-reference text run in the footnote.
        Built-in footnotes (id < 1) are rejected.
        """
        if footnote_id < 1:
            raise ValueError(f"Footnote id {footnote_id} not found")
        fn_tree = self._require("word/footnotes.xml")
        # Find the target footnote element
        target = None
        for fn in fn_tree.findall(f"{W}footnote"):
            if fn.get(f"{W}id") == str(footnote_id):
                target = fn
                break
        if target is None:
            raise ValueError(f"Footnote id {footnote_id} not found")
        # Find all w:r elements inside the footnote paragraphs
        # Skip the reference run (has w:rStyle[@w:val="FootnoteReference"])
        # Update (or create) the first text run after the reference run
        for para in target.findall(f"{W}p"):
            text_run = None
            for run in para.findall(f"{W}r"):
                rpr = run.find(f"{W}rPr")
                is_ref = False
                if rpr is not None:
                    rs = rpr.find(f"{W}rStyle")
                    if rs is not None and rs.get(f"{W}val") == "FootnoteReference":
                        is_ref = True
                if is_ref:
                    continue
                # First non-ref run — look for w:t
                t_el = run.find(f"{W}t")
                if t_el is not None:
                    text_run = run
                    break
            if text_run is not None:
                t_el = text_run.find(f"{W}t")
                _preserve(t_el, text)
                self._mark("word/footnotes.xml")
                return {"footnote_id": footnote_id, "text": text}
            # No text run found — add one to this para
            new_run = etree.SubElement(para, f"{W}r")
            new_t = etree.SubElement(new_run, f"{W}t")
            _preserve(new_t, text)
            self._mark("word/footnotes.xml")
            return {"footnote_id": footnote_id, "text": text}
        # Footnote has no paragraphs — add one
        para = etree.SubElement(target, f"{W}p")
        new_run = etree.SubElement(para, f"{W}r")
        new_t = etree.SubElement(new_run, f"{W}t")
        _preserve(new_t, text)
        self._mark("word/footnotes.xml")
        return {"footnote_id": footnote_id, "text": text}

    def delete_footnote(self, footnote_id: int) -> dict:
        """Delete a footnote definition and its in-body reference run.

        Removes w:footnote from word/footnotes.xml and removes the w:r
        containing w:footnoteReference[@w:id="{footnote_id}"] from document.xml.
        """
        fn_tree = self._require("word/footnotes.xml")
        target = None
        for fn in fn_tree.findall(f"{W}footnote"):
            if fn.get(f"{W}id") == str(footnote_id):
                target = fn
                break
        if target is None:
            raise ValueError(f"Footnote id {footnote_id} not found")
        fn_tree.remove(target)
        self._mark("word/footnotes.xml")
        # Remove in-body reference run
        doc = self._tree("word/document.xml")
        if doc is not None:
            for ref_el in doc.iter(f"{W}footnoteReference"):
                if ref_el.get(f"{W}id") == str(footnote_id):
                    ref_run = ref_el.getparent()
                    if ref_run is not None:
                        para = ref_run.getparent()
                        if para is not None:
                            para.remove(ref_run)
                    self._mark("word/document.xml")
                    break
        return {"deleted": footnote_id}

    def validate_footnotes(self) -> dict:
        """Cross-reference footnote IDs between document.xml and footnotes.xml."""
        doc = self._tree("word/document.xml")
        fn_tree = self._tree("word/footnotes.xml")
        if doc is None:
            return {"error": "No document open"}
        if fn_tree is None:
            return {"valid": True, "references": 0, "definitions": 0}

        ref_ids = set()
        for ref in doc.iter(f"{W}footnoteReference"):
            fid = ref.get(f"{W}id")
            if fid:
                ref_ids.add(int(fid))

        def_ids = {int(f.get(f"{W}id", "0")) for f in self._real_footnotes(fn_tree)}

        missing = sorted(ref_ids - def_ids)
        orphans = sorted(def_ids - ref_ids)
        return {
            "valid": not missing and not orphans,
            "references": len(ref_ids),
            "definitions": len(def_ids),
            "missing_definitions": missing,
            "orphan_definitions": orphans,
        }
