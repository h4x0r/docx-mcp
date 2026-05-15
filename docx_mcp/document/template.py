"""TemplateMixin: fill SDT content controls from a data dict."""

from __future__ import annotations

import copy

from lxml import etree

from .base import W14, W, _preserve


def _set_sdt_text(sdt: etree._Element, val: str) -> None:
    """Replace sdtContent with a plain-text paragraph run."""
    sdtContent = sdt.find(f"{W}sdtContent")
    if sdtContent is None:
        sdtContent = etree.SubElement(sdt, f"{W}sdtContent")
    for child in list(sdtContent):
        sdtContent.remove(child)
    p = etree.SubElement(sdtContent, f"{W}p")
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = str(val)
    _preserve(t, str(val))


class TemplateMixin:
    """Template filling operations using SDT content controls."""

    def fill_template(
        self,
        data: dict[str, str | list[str]],
        remove_empty: bool = False,
    ) -> dict:
        """Fill SDT content controls from data dict.

        - data: {"CLIENT_NAME": "Acme Corp", "DATE": "2026-06-01"}
          For repeating sections: {"ITEMS": ["Item A", "Item B"]}
        - remove_empty: if True, remove SDTs with no matching key in data
        - Returns: {"filled": int, "unfilled": list[str]}
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        self._mark("word/document.xml")  # type: ignore[attr-defined]

        filled = 0
        unfilled: list[str] = []

        # Collect SDTs first (iteration over a live tree during mutation is unsafe)
        sdts = list(doc.iter(f"{W}sdt"))

        for sdt in sdts:
            # Skip if element was already removed from tree
            if sdt.getparent() is None:
                continue

            sdtPr = sdt.find(f"{W}sdtPr")
            tag_el = sdtPr.find(f"{W}tag") if sdtPr is not None else None
            tag = tag_el.get(f"{W}val", "") if tag_el is not None else ""

            if not tag:
                continue

            if tag in data:
                val = data[tag]
                if isinstance(val, list):
                    # Repeating section: first item in existing SDT, rest as clones
                    if val:
                        _set_sdt_text(sdt, val[0])
                        parent = sdt.getparent()
                        if parent is not None:
                            idx = list(parent).index(sdt)
                            for i, item in enumerate(val[1:], 1):
                                clone = copy.deepcopy(sdt)
                                _set_sdt_text(clone, item)
                                parent.insert(idx + i, clone)
                    filled += 1
                else:
                    _set_sdt_text(sdt, str(val))
                    filled += 1
            else:
                if remove_empty:
                    parent = sdt.getparent()
                    if parent is not None:
                        parent.remove(sdt)
                unfilled.append(tag)

        return {"filled": filled, "unfilled": unfilled}

    def list_template_fields(self) -> list[dict]:
        """List all SDT placeholders in the document.

        Returns: [{"tag": str, "label": str, "type": str}]
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]

        results = []
        for sdt in doc.iter(f"{W}sdt"):
            sdtPr = sdt.find(f"{W}sdtPr")
            if sdtPr is None:
                continue
            tag_el = sdtPr.find(f"{W}tag")
            if tag_el is None:
                continue
            tag = tag_el.get(f"{W}val", "")
            if not tag:
                continue
            alias_el = sdtPr.find(f"{W}alias")
            label = alias_el.get(f"{W}val", "") if alias_el is not None else ""
            # Detect type
            type_ = "text"
            if sdtPr.find(f"{W14}checkbox") is not None:
                type_ = "checkbox"
            elif sdtPr.find(f"{W}dropDownList") is not None:
                type_ = "dropdown"
            elif sdtPr.find(f"{W}date") is not None:
                type_ = "date"
            results.append({"tag": tag, "label": label, "type": type_})
        return results

    def validate_template_data(self, data: dict) -> dict:
        """Check data covers all required template fields.

        Returns: {"valid": bool, "missing": list[str], "extra": list[str]}
        missing = fields in doc but not in data
        extra = keys in data not matching any doc field
        """
        fields = self.list_template_fields()
        doc_tags = {f["tag"] for f in fields}
        data_keys = set(data.keys())
        missing = sorted(doc_tags - data_keys)
        extra = sorted(data_keys - doc_tags)
        return {"valid": len(missing) == 0, "missing": missing, "extra": extra}
