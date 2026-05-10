"""TocMixin — Table of Contents, List of Figures, List of Tables generation."""
from __future__ import annotations

import re
from lxml import etree

from .base import W, W14, XML_SPACE, _preserve
from .errors import DocxMcpError, ErrCode

_TOC_FIELD = ' TOC \\o "1-3" \\h \\z \\u '
_LOF_FIELD = ' TOC \\h \\z \\c "Figure" '
_LOT_FIELD = ' TOC \\h \\z \\c "Table" '


class TocMixin:
    """Methods for generating and updating Table of Contents fields."""

    # ── Internal helpers ─────────────────────────────────────────────────────

    def _collect_headings(self, doc_tree, max_level: int) -> list[tuple[str, int]]:
        """Return [(text, level), ...] for all Heading N paragraphs at depth ≤ max_level."""
        results = []
        for p in doc_tree.iter(f"{W}p"):
            pPr = p.find(f"{W}pPr")
            if pPr is None:
                continue
            pStyle = pPr.find(f"{W}pStyle")
            if pStyle is None:
                continue
            style_val = pStyle.get(f"{W}val", "")
            for lvl in range(1, max_level + 1):
                if style_val.lower() in (f"heading {lvl}", f"heading{lvl}"):
                    text = "".join(t.text for t in p.iter(f"{W}t") if t.text)
                    if text:
                        results.append((text, lvl))
                    break
        return results

    def _make_field_para(self, instr_text: str) -> etree._Element:
        """Create a w:p containing a complex TOC field."""
        p = etree.Element(f"{W}p")
        # begin
        r = etree.SubElement(p, f"{W}r")
        fc = etree.SubElement(r, f"{W}fldChar")
        fc.set(f"{W}fldCharType", "begin")
        fc.set(f"{W}dirty", "true")
        # instrText
        r2 = etree.SubElement(p, f"{W}r")
        it = etree.SubElement(r2, f"{W}instrText")
        it.set(XML_SPACE, "preserve")
        it.text = instr_text
        # separate
        r3 = etree.SubElement(p, f"{W}r")
        fc2 = etree.SubElement(r3, f"{W}fldChar")
        fc2.set(f"{W}fldCharType", "separate")
        # end
        r4 = etree.SubElement(p, f"{W}r")
        fc3 = etree.SubElement(r4, f"{W}fldChar")
        fc3.set(f"{W}fldCharType", "end")
        return p

    def _make_toc_entry(self, text: str, level: int) -> etree._Element:
        """Create a paragraph styled 'TOC N' with text + tab + page placeholder."""
        p = etree.Element(f"{W}p")
        pPr = etree.SubElement(p, f"{W}pPr")
        pStyle = etree.SubElement(pPr, f"{W}pStyle")
        pStyle.set(f"{W}val", f"TOC{level}")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = text
        _preserve(t, text)
        # tab
        r2 = etree.SubElement(p, f"{W}r")
        etree.SubElement(r2, f"{W}tab")
        # page number placeholder
        r3 = etree.SubElement(p, f"{W}r")
        t2 = etree.SubElement(r3, f"{W}t")
        t2.text = "1"
        return p

    def _make_title_para(self, title: str) -> etree._Element:
        """Create a plain bold paragraph for the ToC title (TOCHeading style)."""
        p = etree.Element(f"{W}p")
        pPr = etree.SubElement(p, f"{W}pPr")
        pStyle = etree.SubElement(pPr, f"{W}pStyle")
        pStyle.set(f"{W}val", "TOCHeading")
        r = etree.SubElement(p, f"{W}r")
        rPr = etree.SubElement(r, f"{W}rPr")
        etree.SubElement(rPr, f"{W}b")
        t = etree.SubElement(r, f"{W}t")
        t.text = title
        _preserve(t, title)
        return p

    def _find_toc_field_para(self, body) -> tuple[etree._Element | None, int]:
        """Find the first TOC field paragraph (not LoF/LoT) in body."""
        children = list(body)
        for idx, child in enumerate(children):
            if child.tag != f"{W}p":
                continue
            for it in child.iter(f"{W}instrText"):
                if it.text and " TOC " in it.text and "\\c" not in it.text:
                    return child, idx
        return None, -1

    def _find_caption_field_para(self, body, caption_type: str) -> tuple[etree._Element | None, int]:
        """Find existing LoF/LoT field paragraph by caption type (Figure|Table)."""
        children = list(body)
        for idx, child in enumerate(children):
            if child.tag != f"{W}p":
                continue
            for it in child.iter(f"{W}instrText"):
                if it.text and f'\\c "{caption_type}"' in it.text:
                    return child, idx
        return None, -1

    def _remove_toc_entries_after(self, body, field_idx: int) -> None:
        """Remove consecutive TOC N-styled paragraphs (TOC1, TOC2, …) after field_idx."""
        children = list(body)
        to_remove = []
        for i in range(field_idx + 1, len(children)):
            child = children[i]
            if child.tag != f"{W}p":
                break
            pPr = child.find(f"{W}pPr")
            if pPr is None:
                break
            pStyle = pPr.find(f"{W}pStyle")
            if pStyle is None:
                break
            if re.match(r"^TOC\d+$", pStyle.get(f"{W}val", "")):
                to_remove.append(child)
            else:
                break
        for child in to_remove:
            body.remove(child)

    def _insert_toc_block(
        self,
        instr_text: str,
        entries: list[tuple[str, int]],
        title: str | None,
        insert_pos: int,
    ) -> int:
        """Insert title (optional), field para, and entry paras at insert_pos in body."""
        doc_tree = self._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")

        elements: list[etree._Element] = []
        if title:
            elements.append(self._make_title_para(title))
        elements.append(self._make_field_para(instr_text))
        for text, level in entries:
            elements.append(self._make_toc_entry(text, level))

        for offset, el in enumerate(elements):
            body.insert(insert_pos + offset, el)

        self._mark("word/document.xml")
        return insert_pos

    # ── Public API ───────────────────────────────────────────────────────────

    def generate_toc(
        self,
        max_level: int = 3,
        title: str = "Table of Contents",
    ) -> dict:
        """Insert a ToC field + cached entries from current headings at start of document.

        Returns {"inserted_at": 0, "entry_count": int, "title": str}.
        If a ToC field already exists, replaces cached entries and updates field.
        """
        doc_tree = self._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")

        headings = self._collect_headings(doc_tree, max_level)

        # Check if ToC already exists — if so, delegate to update logic
        existing_para, existing_idx = self._find_toc_field_para(body)
        if existing_para is not None:
            # Remove old cached entries and re-insert fresh ones
            self._remove_toc_entries_after(body, existing_idx)
            for offset, (text, level) in enumerate(headings):
                body.insert(existing_idx + 1 + offset, self._make_toc_entry(text, level))
            # Mark field dirty
            for fc in existing_para.iter(f"{W}fldChar"):
                if fc.get(f"{W}fldCharType") == "begin":
                    fc.set(f"{W}dirty", "true")
            self._mark("word/document.xml")
            return {"inserted_at": existing_idx, "entry_count": len(headings), "title": title}

        # Build field instruction using the max_level range
        if max_level == 3:
            instr = _TOC_FIELD
        else:
            instr = f' TOC \\o "1-{max_level}" \\h \\z \\u '

        inserted_at = self._insert_toc_block(instr, headings, title, 0)
        return {"inserted_at": inserted_at, "entry_count": len(headings), "title": title}

    def update_toc(self) -> dict:
        """Regenerate cached ToC entries from current headings.

        Finds existing TOC field paragraph, removes old cached entries,
        inserts fresh ones. Also marks TOC field as dirty.
        Returns {"entry_count": int, "updated": True}.
        """
        doc_tree = self._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")

        field_para, field_idx = self._find_toc_field_para(body)
        if field_para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                "No TOC field found in document. Call generate_toc first.",
                hint="Use generate_toc to insert a TOC before calling update_toc.",
            )

        # Determine max_level from field instrText
        max_level = 3
        for it in field_para.iter(f"{W}instrText"):
            if it.text and " TOC " in it.text:
                m = re.search(r'\\o\s+"1-(\d+)"', it.text)
                if m:
                    max_level = int(m.group(1))
                break

        headings = self._collect_headings(doc_tree, max_level)

        # Remove old cached entries
        self._remove_toc_entries_after(body, field_idx)

        # Insert fresh entries
        for offset, (text, level) in enumerate(headings):
            body.insert(field_idx + 1 + offset, self._make_toc_entry(text, level))

        # Mark field dirty
        for fc in field_para.iter(f"{W}fldChar"):
            if fc.get(f"{W}fldCharType") == "begin":
                fc.set(f"{W}dirty", "true")

        self._mark("word/document.xml")
        return {"entry_count": len(headings), "updated": True}

    def _collect_captions(self, doc_tree, prefix: str) -> list[tuple[str, int]]:
        """Return [(text, 1), ...] for caption paragraphs whose text starts with prefix."""
        results = []
        for p in doc_tree.iter(f"{W}p"):
            pPr = p.find(f"{W}pPr")
            if pPr is None:
                continue
            pStyle = pPr.find(f"{W}pStyle")
            if pStyle is None:
                continue
            if pStyle.get(f"{W}val", "").lower() != "caption":
                continue
            text = "".join(t.text for t in p.iter(f"{W}t") if t.text)
            if text.startswith(prefix):
                results.append((text, 1))
        return results

    def generate_tof(self, para_id: str, title: str = "List of Figures") -> dict:
        """Insert a Table of Figures field block AFTER the paragraph with para_id.

        Returns {"para_id": str, "title": str, "entry_count": int}.
        Raises ValueError if para_id not found.
        """
        doc = self._tree("word/document.xml")
        body = doc.find(f"{W}body")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph with paraId '{para_id}' not found")
        body_children = list(body)
        idx = body_children.index(para)
        entries = self._collect_captions(doc, "Figure")
        self._insert_toc_block(_LOF_FIELD, entries, title, idx + 1)
        return {"para_id": para_id, "title": title, "entry_count": len(entries)}

    def generate_tot(self, para_id: str, title: str = "List of Tables") -> dict:
        """Insert a Table of Tables field block AFTER the paragraph with para_id.

        Returns {"para_id": str, "title": str, "entry_count": int}.
        Raises ValueError if para_id not found.
        """
        doc = self._tree("word/document.xml")
        body = doc.find(f"{W}body")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph with paraId '{para_id}' not found")
        body_children = list(body)
        idx = body_children.index(para)
        entries = self._collect_captions(doc, "Table")
        self._insert_toc_block(_LOT_FIELD, entries, title, idx + 1)
        return {"para_id": para_id, "title": title, "entry_count": len(entries)}

    def generate_list_of_figures(self) -> dict:
        """Insert a List of Figures (TOC \\c "Figure") field.

        Returns {"inserted_at": int, "entry_count": int}.
        Entry count is 0 since blank docs have no SEQ Figure captions.
        """
        doc_tree = self._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")

        # Count existing SEQ Figure fields as a proxy for entry count
        entry_count = 0
        for it in doc_tree.iter(f"{W}instrText"):
            if it.text and "SEQ" in it.text and "Figure" in it.text:
                entry_count += 1

        # Insert after any existing ToC block, or at position 0
        _, toc_idx = self._find_toc_field_para(body)
        if toc_idx >= 0:
            # Find end of ToC block
            children = list(body)
            insert_pos = toc_idx + 1
            while insert_pos < len(children):
                child = children[insert_pos]
                if child.tag != f"{W}p":
                    break
                pPr = child.find(f"{W}pPr")
                if pPr is None:
                    break
                pStyle = pPr.find(f"{W}pStyle")
                if pStyle is not None and re.match(r"^TOC\d+$", pStyle.get(f"{W}val", "")):
                    insert_pos += 1
                else:
                    break
        else:
            insert_pos = 0

        inserted_at = self._insert_toc_block(_LOF_FIELD, [], None, insert_pos)
        return {"inserted_at": inserted_at, "entry_count": entry_count}

    def generate_list_of_tables(self) -> dict:
        """Insert a List of Tables (TOC \\c "Table") field.

        Returns {"inserted_at": int, "entry_count": int}.
        """
        doc_tree = self._tree("word/document.xml")
        body = doc_tree.find(f"{W}body")

        # Count existing SEQ Table fields as proxy
        entry_count = 0
        for it in doc_tree.iter(f"{W}instrText"):
            if it.text and "SEQ" in it.text and "Table" in it.text:
                entry_count += 1

        # Insert after any existing ToC/LoF block, or at position 0
        _, toc_idx = self._find_toc_field_para(body)
        _, lof_idx = self._find_caption_field_para(body, "Figure")
        anchor_idx = max(toc_idx, lof_idx)
        if anchor_idx >= 0:
            children = list(body)
            insert_pos = anchor_idx + 1
            while insert_pos < len(children):
                child = children[insert_pos]
                if child.tag != f"{W}p":
                    break
                pPr = child.find(f"{W}pPr")
                if pPr is None:
                    break
                pStyle = pPr.find(f"{W}pStyle")
                if pStyle is not None and re.match(r"^TOC\d+$", pStyle.get(f"{W}val", "")):
                    insert_pos += 1
                else:
                    break
        else:
            insert_pos = 0

        inserted_at = self._insert_toc_block(_LOT_FIELD, [], None, insert_pos)
        return {"inserted_at": inserted_at, "entry_count": entry_count}
