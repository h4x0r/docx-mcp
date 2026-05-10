"""Markdown export mixin."""
from __future__ import annotations
import os
import re
from lxml import etree
from .base import W

_HEADING_RE = re.compile(r"^heading\s*([123])$", re.IGNORECASE)


def _heading_level(ppr) -> int | None:
    if ppr is None:
        return None
    ps = ppr.find(f"{W}pStyle")
    if ps is None:
        return None
    val = (ps.get(f"{W}val") or "").strip()
    m = _HEADING_RE.match(val)
    if m:
        return int(m.group(1))
    return None


def _has_num_pr(ppr) -> bool:
    if ppr is None:
        return False
    numpr = ppr.find(f"{W}numPr")
    if numpr is None:
        return False
    numid = numpr.find(f"{W}numId")
    if numid is None:
        return False
    val = numid.get(f"{W}val") or "0"
    return val != "0"


def _run_text(run) -> str:
    rpr = run.find(f"{W}rPr")
    parts = []
    for t in run.iter(f"{W}t"):
        parts.append(t.text or "")
    text = "".join(parts)
    if not text:
        return ""
    bold = rpr is not None and rpr.find(f"{W}b") is not None
    italic = rpr is not None and rpr.find(f"{W}i") is not None
    if bold:
        text = f"**{text}**"
    elif italic:
        text = f"*{text}*"
    return text


def _para_to_md(para) -> str:
    ppr = para.find(f"{W}pPr")
    level = _heading_level(ppr)
    if level is not None:
        text = "".join(t.text or "" for t in para.iter(f"{W}t"))
        return "#" * level + " " + text

    runs = para.findall(f"{W}r")
    if not runs:
        return ""

    text = "".join(_run_text(r) for r in runs)
    if not text:
        return ""

    if _has_num_pr(ppr):
        return f"- {text}"

    return text


def _cell_text(tc) -> str:
    parts = []
    for t in tc.iter(f"{W}t"):
        parts.append(t.text or "")
    return "".join(parts)


def _table_to_md(tbl) -> str:
    rows = tbl.findall(f"{W}tr")
    if not rows:
        return ""
    lines = []
    for i, row in enumerate(rows):
        cells = row.findall(f"{W}tc")
        values = [_cell_text(c) for c in cells]
        lines.append("| " + " | ".join(values) + " |")
        if i == 0:
            lines.append("| " + " | ".join("---" for _ in values) + " |")
    return "\n".join(lines)


class MarkdownExportMixin:
    def export_markdown(self, output_path: str = "") -> dict:
        if not output_path:
            output_path = str(self.workdir / "export.md")

        root = self._tree("word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            body = root

        skip_tags = {f"{W}sectPr", f"{W}tblPr", f"{W}tblGrid"}
        lines: list[str] = []
        para_count = 0
        table_count = 0

        for child in body:
            if child.tag in skip_tags:
                continue
            if child.tag == f"{W}tbl":
                table_count += 1
                lines.append(_table_to_md(child))
            elif child.tag == f"{W}p":
                md = _para_to_md(child)
                lines.append(md)
                para_count += 1

        content = "\n".join(lines)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)

        return {"output_path": output_path, "paragraphs": para_count, "tables": table_count}
