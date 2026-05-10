"""Styles mixin: enumerate document styles."""

from __future__ import annotations

import re

from lxml import etree

from .base import W


class StylesMixin:
    """Style inspection."""

    def get_styles(self) -> list[dict]:
        """Get all defined styles."""
        tree = self._tree("word/styles.xml")
        if tree is None:
            return []
        result = []
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            based_el = s.find(f"{W}basedOn")
            result.append(
                {
                    "id": s.get(f"{W}styleId", ""),
                    "name": name_el.get(f"{W}val", "") if name_el is not None else "",
                    "type": s.get(f"{W}type", ""),
                    "base_style": based_el.get(f"{W}val", "") if based_el is not None else "",
                }
            )
        return result

    def create_style(
        self,
        name: str,
        style_type: str,
        based_on: str | None = None,
        next_style: str | None = None,
    ) -> dict:
        tree = self._require("word/styles.xml")
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            existing = name_el.get(f"{W}val", "") if name_el is not None else ""
            if existing.lower() == name.lower():
                raise ValueError(f"Style '{name}' already exists")
        style_id = re.sub(r"\s+", "", name)
        style_el = etree.SubElement(tree, f"{W}style")
        style_el.set(f"{W}type", style_type)
        style_el.set(f"{W}styleId", style_id)
        name_el = etree.SubElement(style_el, f"{W}name")
        name_el.set(f"{W}val", name)
        if based_on is not None:
            based_el = etree.SubElement(style_el, f"{W}basedOn")
            based_el.set(f"{W}val", based_on)
        if next_style is not None:
            next_el = etree.SubElement(style_el, f"{W}next")
            next_el.set(f"{W}val", next_style)
        self._mark("word/styles.xml")
        return {"style_id": style_id, "name": name, "type": style_type}

    def update_style(
        self,
        name: str,
        based_on: str | None = None,
        next_style: str | None = None,
    ) -> dict:
        tree = self._require("word/styles.xml")
        target = None
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            val = name_el.get(f"{W}val", "") if name_el is not None else ""
            if val.lower() == name.lower():
                target = s
                break
            if s.get(f"{W}styleId", "").lower() == name.lower():
                target = s
                break
        if target is None:
            raise ValueError(f"Style '{name}' not found")
        style_id = target.get(f"{W}styleId", "")
        changed = False
        if based_on is not None:
            old = target.find(f"{W}basedOn")
            if old is not None:
                target.remove(old)
            based_el = etree.SubElement(target, f"{W}basedOn")
            based_el.set(f"{W}val", based_on)
            changed = True
        if next_style is not None:
            old = target.find(f"{W}next")
            if old is not None:
                target.remove(old)
            next_el = etree.SubElement(target, f"{W}next")
            next_el.set(f"{W}val", next_style)
            changed = True
        if changed:
            self._mark("word/styles.xml")
        return {"style_id": style_id, "name": name}

    def delete_style(self, name: str) -> dict:
        tree = self._require("word/styles.xml")
        target = None
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            val = name_el.get(f"{W}val", "") if name_el is not None else ""
            if val.lower() == name.lower() or s.get(f"{W}styleId", "").lower() == name.lower():
                target = s
                break
        if target is None:
            raise ValueError(f"Style '{name}' not found")
        if target.get(f"{W}type") == "paragraph" and target.get(f"{W}default") == "1":
            raise ValueError(f"Cannot delete default paragraph style '{name}'")
        style_id = target.get(f"{W}styleId", "")
        tree.remove(target)
        self._mark("word/styles.xml")
        return {"deleted": style_id}
