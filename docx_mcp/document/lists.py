"""Lists mixin: create bullet and numbered lists."""

from __future__ import annotations

from lxml import etree

from .base import NSMAP, W


class ListsMixin:
    """List operations."""

    def add_list(
        self,
        para_ids: list[str],
        *,
        style: str = "bullet",
    ) -> dict:
        """Apply list formatting to paragraphs.

        Creates numbering definitions in numbering.xml (bootstrapped if missing)
        and sets w:numPr on each target paragraph.

        Args:
            para_ids: List of paraIds to format as list items.
            style: "bullet" or "numbered".
        """
        doc = self._require("word/document.xml")

        # Bootstrap numbering.xml if missing
        num_tree = self._tree("word/numbering.xml")
        if num_tree is None:
            num_tree = etree.Element(f"{W}numbering", nsmap=NSMAP)
            self._trees["word/numbering.xml"] = num_tree
            self._mark("word/numbering.xml")

        # Determine next abstract num ID and num ID
        existing_abstract = num_tree.findall(f"{W}abstractNum")
        abs_id = (
            max(
                (int(a.get(f"{W}abstractNumId", "0")) for a in existing_abstract),
                default=-1,
            )
            + 1
        )
        existing_nums = num_tree.findall(f"{W}num")
        num_id = max((int(n.get(f"{W}numId", "0")) for n in existing_nums), default=0) + 1

        # Create abstract numbering definition
        abstract = etree.SubElement(num_tree, f"{W}abstractNum")
        abstract.set(f"{W}abstractNumId", str(abs_id))
        lvl = etree.SubElement(abstract, f"{W}lvl")
        lvl.set(f"{W}ilvl", "0")
        fmt = etree.SubElement(lvl, f"{W}numFmt")
        if style == "numbered":
            fmt.set(f"{W}val", "decimal")
            lvl_text = etree.SubElement(lvl, f"{W}lvlText")
            lvl_text.set(f"{W}val", "%1.")
        else:
            fmt.set(f"{W}val", "bullet")
            lvl_text = etree.SubElement(lvl, f"{W}lvlText")
            lvl_text.set(f"{W}val", "\u2022")

        # Create num entry referencing abstract
        num_el = etree.SubElement(num_tree, f"{W}num")
        num_el.set(f"{W}numId", str(num_id))
        ref = etree.SubElement(num_el, f"{W}abstractNumId")
        ref.set(f"{W}val", str(abs_id))

        self._mark("word/numbering.xml")

        # Apply numPr to each paragraph
        count = 0
        for pid in para_ids:
            para = self._find_para(doc, pid)
            if para is None:
                raise ValueError(f"Paragraph '{pid}' not found")

            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.SubElement(para, f"{W}pPr")
                para.remove(ppr)
                para.insert(0, ppr)

            num_pr = etree.SubElement(ppr, f"{W}numPr")
            ilvl = etree.SubElement(num_pr, f"{W}ilvl")
            ilvl.set(f"{W}val", "0")
            nid = etree.SubElement(num_pr, f"{W}numId")
            nid.set(f"{W}val", str(num_id))
            count += 1

        self._mark("word/document.xml")
        return {"list_id": num_id, "paragraphs_updated": count}

    def create_multilevel_list(
        self,
        name: str,
        levels: list[dict],
    ) -> dict:
        """Create a multilevel abstractNum + num entry in numbering.xml.

        Each level dict has keys:
          - num_fmt: str — "decimal", "lowerLetter", "lowerRoman", "bullet", etc.
          - lvl_text: str — format string like "%1." or "%1.%2." or "•"
          - indent: int — left indent in twips (default 720)
          - hanging: int — hanging indent in twips (default 360)
          - style: str | None — linked style (e.g. "Heading 1"), optional

        Returns {"abstract_num_id": int, "num_id": int, "name": str, "level_count": int}.
        """
        # Bootstrap numbering.xml if missing
        num_tree = self._tree("word/numbering.xml")
        if num_tree is None:
            num_tree = etree.Element(f"{W}numbering", nsmap=NSMAP)
            self._trees["word/numbering.xml"] = num_tree
            self._mark("word/numbering.xml")

        # Determine next abstract num ID and num ID
        existing_abstract = num_tree.findall(f"{W}abstractNum")
        abs_id = (
            max(
                (int(a.get(f"{W}abstractNumId", "0")) for a in existing_abstract),
                default=-1,
            )
            + 1
        )
        existing_nums = num_tree.findall(f"{W}num")
        num_id = max((int(n.get(f"{W}numId", "0")) for n in existing_nums), default=0) + 1

        # Create abstractNum element
        abstract = etree.SubElement(num_tree, f"{W}abstractNum")
        abstract.set(f"{W}abstractNumId", str(abs_id))

        # multiLevelType
        ml_type = etree.SubElement(abstract, f"{W}multiLevelType")
        ml_type.set(f"{W}val", "multilevel")

        # name
        name_el = etree.SubElement(abstract, f"{W}name")
        name_el.set(f"{W}val", name)

        # Create w:lvl elements for each level dict
        for ilvl, lvl_def in enumerate(levels):
            lvl = etree.SubElement(abstract, f"{W}lvl")
            lvl.set(f"{W}ilvl", str(ilvl))

            start_el = etree.SubElement(lvl, f"{W}start")
            start_el.set(f"{W}val", "1")

            fmt_el = etree.SubElement(lvl, f"{W}numFmt")
            fmt_el.set(f"{W}val", lvl_def.get("num_fmt", "decimal"))

            lvl_text_el = etree.SubElement(lvl, f"{W}lvlText")
            lvl_text_el.set(f"{W}val", lvl_def.get("lvl_text", ""))

            jc_el = etree.SubElement(lvl, f"{W}lvlJc")
            jc_el.set(f"{W}val", "left")

            pPr = etree.SubElement(lvl, f"{W}pPr")
            ind = etree.SubElement(pPr, f"{W}ind")
            ind.set(f"{W}left", str(lvl_def.get("indent", 720)))
            ind.set(f"{W}hanging", str(lvl_def.get("hanging", 360)))

            # Bind to style if specified
            style_val = lvl_def.get("style")
            if style_val:
                pStyle_el = etree.SubElement(pPr, f"{W}pStyle")
                pStyle_el.set(f"{W}val", style_val)

        # Create num entry referencing abstractNum
        num_el = etree.SubElement(num_tree, f"{W}num")
        num_el.set(f"{W}numId", str(num_id))
        ref = etree.SubElement(num_el, f"{W}abstractNumId")
        ref.set(f"{W}val", str(abs_id))

        self._mark("word/numbering.xml")
        return {
            "abstract_num_id": abs_id,
            "num_id": num_id,
            "name": name,
            "level_count": len(levels),
        }

    def restart_numbering(
        self,
        para_id: str,
        level: int,
        start: int = 1,
    ) -> dict:
        """Restart numbering at a paragraph by adding lvlOverride with startOverride.

        The paragraph must already have w:numPr (i.e., be part of a list).
        Adds/updates w:num/w:lvlOverride/w:startOverride in numbering.xml.
        Returns {"para_id": str, "level": int, "start": int}.
        Raises ValueError if paragraph not found or has no numPr.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        num_id_el = para.find(f".//{W}numId")
        if num_id_el is None:
            raise ValueError(f"Paragraph '{para_id}' has no numPr/numId")

        num_id_val = num_id_el.get(f"{W}val")

        num_tree = self._tree("word/numbering.xml")
        if num_tree is None:
            raise ValueError("numbering.xml not found")

        # Find the w:num element matching num_id
        target_num = None
        for n in num_tree.findall(f"{W}num"):
            if n.get(f"{W}numId") == num_id_val:
                target_num = n
                break
        if target_num is None:
            raise ValueError(f"w:num with numId='{num_id_val}' not found in numbering.xml")

        # Check for existing lvlOverride at this level
        existing_override = None
        for ov in target_num.findall(f"{W}lvlOverride"):
            if ov.get(f"{W}ilvl") == str(level):
                existing_override = ov
                break

        if existing_override is None:
            existing_override = etree.SubElement(target_num, f"{W}lvlOverride")
            existing_override.set(f"{W}ilvl", str(level))

        # Set/update startOverride
        start_ov = existing_override.find(f"{W}startOverride")
        if start_ov is None:
            start_ov = etree.SubElement(existing_override, f"{W}startOverride")
        start_ov.set(f"{W}val", str(start))

        self._mark("word/numbering.xml")
        return {"para_id": para_id, "level": level, "start": start}

    def get_lists(self) -> list[dict]:
        """Return all list definitions from word/numbering.xml.

        Each abstractNum element becomes a dict with:
          - abstract_num_id: int
          - num_format: str  (w:numFmt w:val of level 0, or "" if absent)
          - levels: int      (count of w:lvl children)

        Returns [] if numbering.xml doesn't exist.
        """
        num_tree = self._tree("word/numbering.xml")
        if num_tree is None:
            return []

        result = []
        for abstract in num_tree.findall(f"{W}abstractNum"):
            abs_id_str = abstract.get(f"{W}abstractNumId", "0")
            try:
                abs_id = int(abs_id_str)
            except ValueError:
                abs_id = 0

            lvl_elements = abstract.findall(f"{W}lvl")
            levels = len(lvl_elements)

            # num_format from level 0
            num_format = ""
            if lvl_elements:
                fmt_el = lvl_elements[0].find(f"{W}numFmt")
                if fmt_el is not None:
                    num_format = fmt_el.get(f"{W}val", "")

            result.append({
                "abstract_num_id": abs_id,
                "num_format": num_format,
                "levels": levels,
            })

        return result

    def promote_list_item(self, para_id: str) -> dict:
        """Decrease the list indentation level (ilvl) of a paragraph by 1, min 0.

        Returns {"para_id": para_id, "ilvl": new_value}.
        Raises ValueError if paragraph is not a list item.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ilvl_el = para.find(f".//{W}numPr/{W}ilvl")
        if ilvl_el is None:
            raise ValueError("Paragraph is not a list item")

        current = int(ilvl_el.get(f"{W}val", "0"))
        new_val = max(0, current - 1)
        ilvl_el.set(f"{W}val", str(new_val))
        self._mark("word/document.xml")
        return {"para_id": para_id, "ilvl": new_val}

    def demote_list_item(self, para_id: str) -> dict:
        """Increase the list indentation level (ilvl) of a paragraph by 1, max 8.

        Returns {"para_id": para_id, "ilvl": new_value}.
        Raises ValueError if paragraph is not a list item.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ilvl_el = para.find(f".//{W}numPr/{W}ilvl")
        if ilvl_el is None:
            raise ValueError("Paragraph is not a list item")

        current = int(ilvl_el.get(f"{W}val", "0"))
        new_val = min(8, current + 1)
        ilvl_el.set(f"{W}val", str(new_val))
        self._mark("word/document.xml")
        return {"para_id": para_id, "ilvl": new_val}

    def suppress_numbering(self, para_id: str) -> dict:
        """Remove list numbering from a paragraph by setting numId to 0.

        Sets w:numPr/w:numId/@w:val = "0" in the paragraph's pPr.
        Creates w:numPr if not present.
        Returns {"para_id": str, "suppressed": True}.
        Raises ValueError if paragraph not found.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        num_pr = ppr.find(f"{W}numPr")
        if num_pr is None:
            num_pr = etree.SubElement(ppr, f"{W}numPr")
            ilvl_el = etree.SubElement(num_pr, f"{W}ilvl")
            ilvl_el.set(f"{W}val", "0")
            nid_el = etree.SubElement(num_pr, f"{W}numId")
            nid_el.set(f"{W}val", "0")
        else:
            nid_el = num_pr.find(f"{W}numId")
            if nid_el is None:
                nid_el = etree.SubElement(num_pr, f"{W}numId")
            nid_el.set(f"{W}val", "0")

        self._mark("word/document.xml")
        return {"para_id": para_id, "suppressed": True}
