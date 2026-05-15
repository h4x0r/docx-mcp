"""Theme mixin: read and write Word theme colors from word/theme/theme1.xml."""

from __future__ import annotations

from lxml import etree

from .base import A

_THEME_PATH = "word/theme/theme1.xml"

_VALID_SLOTS = {
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
}

_ANS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class ThemeMixin:
    def get_theme_colors(self) -> dict:
        theme = self._tree(_THEME_PATH)
        if theme is None:
            return {}
        clr_scheme = theme.find(f"{A}themeElements/{A}clrScheme")
        if clr_scheme is None:
            return {}
        result: dict[str, str] = {}
        for slot in _VALID_SLOTS:
            slot_el = clr_scheme.find(f"{A}{slot}")
            if slot_el is None:
                continue
            srgb = slot_el.find(f"{A}srgbClr")
            if srgb is not None:
                result[slot] = srgb.get("val", "")
                continue
            sys_clr = slot_el.find(f"{A}sysClr")
            if sys_clr is not None:
                result[slot] = sys_clr.get("lastClr", "")
        return result

    def set_theme_color(self, slot: str, hex_color: str) -> dict:
        if slot not in _VALID_SLOTS:
            raise ValueError(f"unknown slot '{slot}': must be one of {sorted(_VALID_SLOTS)}")
        theme = self._tree(_THEME_PATH)
        if theme is None:
            raise RuntimeError("No theme file found in document")
        clr_scheme = theme.find(f"{A}themeElements/{A}clrScheme")
        if clr_scheme is None:
            raise RuntimeError("Theme XML missing clrScheme element")
        slot_el = clr_scheme.find(f"{A}{slot}")
        if slot_el is None:
            slot_el = etree.SubElement(clr_scheme, f"{A}{slot}")
        for child in list(slot_el):
            slot_el.remove(child)
        srgb = etree.SubElement(slot_el, f"{A}srgbClr")
        srgb.set("val", hex_color)
        self._mark(_THEME_PATH)
        return {"slot": slot, "hex_color": hex_color}
