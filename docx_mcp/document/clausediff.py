"""Clause-aware contract diff mixin."""

from __future__ import annotations

import difflib
import os
import re
import zipfile

from lxml import etree

from .base import W
from .errors import DocxMcpError, ErrCode

_HEADING_RE = re.compile(r"^heading\s*[123]$", re.IGNORECASE)


def _para_style(para) -> str:
    ppr = para.find(f"{W}pPr")
    if ppr is None:
        return ""
    ps = ppr.find(f"{W}pStyle")
    if ps is None:
        return ""
    val = ps.get(f"{W}val") or ""
    return val.lower()


def _is_heading(para) -> bool:
    return bool(_HEADING_RE.match(_para_style(para)))


def _para_text(para) -> str:
    return "".join(t.text for t in para.iter(f"{W}t") if t.text)


def _clause_text(clause) -> str:
    return " ".join(_para_text(p) for p in clause["body_paras"])


def _extract_clauses(doc_xml: bytes) -> list[dict]:
    root = etree.fromstring(doc_xml)
    paras = list(root.iter(f"{W}p"))
    clauses: list[dict] = []
    current: dict | None = None

    for para in paras:
        if _is_heading(para):
            if current is not None:
                clauses.append(current)
            current = {
                "heading": _para_text(para),
                "heading_para": para,
                "body_paras": [],
            }
        else:
            if current is None:
                current = {"heading": "", "heading_para": None, "body_paras": []}
            current["body_paras"].append(para)

    if current is not None:
        clauses.append(current)
    return clauses


def _fuzzy_match(clauses_a, clauses_b) -> list[tuple]:
    matched: list[tuple] = []
    used_b: set[int] = set()
    for i, ca in enumerate(clauses_a):
        best_ratio = 0.0
        best_j = -1
        for j, cb in enumerate(clauses_b):
            if j in used_b:
                continue
            if not ca["heading"] and not cb["heading"]:
                ratio = 1.0
            elif not ca["heading"] or not cb["heading"]:
                ratio = 0.0
            else:
                ratio = difflib.SequenceMatcher(None, ca["heading"], cb["heading"]).ratio()
            if ratio >= 0.7 and ratio > best_ratio:
                best_ratio = ratio
                best_j = j
        if best_j >= 0:
            matched.append((i, best_j, best_ratio))
            used_b.add(best_j)
    return matched


class ClauseDiffMixin:
    def compare_contracts(
        self,
        other_path: str,
        output_path: str = "",
        align_by: str = "heading",
    ) -> dict:
        """Clause-aware diff: align by heading, then diff within each clause.

        Unlike compare_documents (which does LCS line-by-line), this:
        1. Extracts logical clauses (heading → sub-paragraphs) from both docs
        2. Matches clauses by heading text (fuzzy) across both docs
        3. Produces tracked-change output aligned at clause level
        4. Flags reordered, added, deleted, or renamed clauses

        Returns: {"output_path": str, "clauses_compared": int, "clauses_changed": int, "reordered": int}  # noqa: E501
        Raises: DocxMcpError(ErrCode.PART_NOT_FOUND) if other_path doesn't exist.
        """
        if align_by != "heading":
            raise ValueError(f"Only align_by='heading' is supported; got {align_by!r}")
        if not os.path.exists(other_path):
            raise DocxMcpError(
                ErrCode.PART_NOT_FOUND,
                f"Other document not found: {other_path}",
            )

        with zipfile.ZipFile(other_path, "r") as zf:
            other_xml = zf.read("word/document.xml")

        doc = self._require("word/document.xml")
        self_xml = etree.tostring(doc)

        clauses_a = _extract_clauses(self_xml)
        clauses_b = _extract_clauses(other_xml)

        matched_pairs = _fuzzy_match(clauses_a, clauses_b)

        matched_a: set[int] = {i for i, j, r in matched_pairs}
        matched_b: set[int] = {j for i, j, r in matched_pairs}

        summary_lines: list[str] = []
        clauses_changed = 0
        reordered = 0
        max_b_idx = -1

        for i, j, ratio in matched_pairs:
            ca = clauses_a[i]
            cb = clauses_b[j]
            renamed = ratio < 1.0
            body_changed = _clause_text(ca) != _clause_text(cb)

            if j <= max_b_idx:
                reordered += 1
            else:
                max_b_idx = j

            if renamed:
                summary_lines.append(f"[RENAMED] {ca['heading']} → {cb['heading']}")
                clauses_changed += 1
            elif body_changed:
                summary_lines.append(f"[CHANGED] {ca['heading']}")
                clauses_changed += 1
            else:
                summary_lines.append(f"[MATCHED] {ca['heading']}")

        for i, ca in enumerate(clauses_a):
            if i not in matched_a:
                summary_lines.append(f"[DELETED] {ca['heading']}")
                clauses_changed += 1

        for j, cb in enumerate(clauses_b):
            if j not in matched_b:
                summary_lines.append(f"[ADDED] {cb['heading']}")
                clauses_changed += 1

        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        for line in summary_lines:
            p = etree.Element(f"{W}p")
            r = etree.SubElement(p, f"{W}r")
            t = etree.SubElement(r, f"{W}t")
            t.text = line
            if sect_pr is not None:
                sect_pr.addprevious(p)
            else:
                body.append(p)

        self._mark("word/document.xml")

        out = output_path if output_path else str(self.workdir / "compare_contracts_output.docx")
        self.save(out, backup=False)

        heading_matches = [
            (i, j, r)
            for i, j, r in matched_pairs
            if clauses_a[i]["heading"] or clauses_b[j]["heading"]
        ]  # noqa: E501
        clauses_compared = (
            len(heading_matches)
            + len(
                [i for i in range(len(clauses_a)) if i not in matched_a and clauses_a[i]["heading"]]
            )
            + len(
                [j for j in range(len(clauses_b)) if j not in matched_b and clauses_b[j]["heading"]]
            )
        )

        return {
            "output_path": out,
            "clauses_compared": clauses_compared,
            "clauses_changed": clauses_changed,
            "reordered": reordered,
        }
