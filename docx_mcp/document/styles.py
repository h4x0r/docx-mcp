"""Styles mixin: enumerate document styles."""

from __future__ import annotations

import copy
import re

from lxml import etree

from .base import W
from .errors import DocxMcpError, ErrCode


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

    # ── helpers ──────────────────────────────────────────────────────────────

    def _find_style(self, tree, name_or_id: str):
        """Find style element by name or styleId (case-insensitive). Returns element or None."""
        lo = name_or_id.lower()
        for s in tree.findall(f"{W}style"):
            if s.get(f"{W}styleId", "").lower() == lo:
                return s
            name_el = s.find(f"{W}name")
            if name_el is not None and name_el.get(f"{W}val", "").lower() == lo:
                return s
        return None

    # ── extended style API ────────────────────────────────────────────────────

    def get_style(self, name_or_id: str) -> dict:
        """Find a style by name or styleId (case-insensitive).

        Returns:
            {"style_id": str, "name": str, "type": str, "base_style": str, "next_style": str}

        Raises:
            ValueError: if style not found.
        """
        tree = self._require("word/styles.xml")
        s = self._find_style(tree, name_or_id)
        if s is None:
            raise ValueError(f"Style '{name_or_id}' not found")
        name_el = s.find(f"{W}name")
        based_el = s.find(f"{W}basedOn")
        next_el = s.find(f"{W}next")
        return {
            "style_id": s.get(f"{W}styleId", ""),
            "name": name_el.get(f"{W}val", "") if name_el is not None else "",
            "type": s.get(f"{W}type", ""),
            "base_style": based_el.get(f"{W}val", "") if based_el is not None else "",
            "next_style": next_el.get(f"{W}val", "") if next_el is not None else "",
        }

    def copy_style(self, source_name_or_id: str, new_name: str) -> dict:
        """Deep-copy an existing style under a new name.

        Returns:
            {"style_id": str, "name": str, "type": str}

        Raises:
            ValueError: if source not found or new_name already exists.
        """
        tree = self._require("word/styles.xml")
        source = self._find_style(tree, source_name_or_id)
        if source is None:
            raise ValueError(f"Style '{source_name_or_id}' not found")
        new_id = re.sub(r"\s+", "", new_name)
        # Check both display name and computed styleId to avoid duplicate IDs
        if self._find_style(tree, new_name) is not None or self._find_style(tree, new_id) is not None:
            raise ValueError(f"Style '{new_name}' already exists")

        new_el = copy.deepcopy(source)
        new_el.attrib.pop(f"{W}default", None)  # prevent duplicate default style
        new_el.set(f"{W}styleId", new_id)

        name_el = new_el.find(f"{W}name")
        if name_el is None:
            name_el = etree.SubElement(new_el, f"{W}name")
        name_el.set(f"{W}val", new_name)

        tree.append(new_el)
        self._mark("word/styles.xml")
        return {
            "style_id": new_id,
            "name": new_name,
            "type": new_el.get(f"{W}type", ""),
        }

    def apply_style_to_range(self, para_ids: list[str], style_name_or_id: str) -> dict:
        """Apply a style to a list of paragraphs identified by their paraIds.

        Returns:
            {"applied": int, "style_id": str, "para_ids": list[str]}

        Raises:
            ValueError: if style not found.
            DocxMcpError: if a para_id is not found.
        """
        styles_tree = self._require("word/styles.xml")
        style_el = self._find_style(styles_tree, style_name_or_id)
        if style_el is None:
            raise ValueError(f"Style '{style_name_or_id}' not found")
        canonical_id = style_el.get(f"{W}styleId", "")

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        for para_id in para_ids:
            para = self._find_para(body, para_id)
            if para is None:
                raise DocxMcpError(
                    ErrCode.PARA_NOT_FOUND,
                    f"Paragraph with paraId '{para_id}' not found.",
                )
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.SubElement(para, f"{W}pPr")
                para.insert(0, ppr)
            pstyle = ppr.find(f"{W}pStyle")
            if pstyle is None:
                pstyle = etree.SubElement(ppr, f"{W}pStyle")
                ppr.insert(0, pstyle)
            pstyle.set(f"{W}val", canonical_id)

        if para_ids:
            self._mark("word/document.xml")
        return {
            "applied": len(para_ids),
            "style_id": canonical_id,
            "para_ids": list(para_ids),
        }
