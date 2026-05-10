"""Review merge mixin: consolidate tracked changes from N reviewer copies."""

from __future__ import annotations

import copy
import zipfile

from lxml import etree

from .base import W, W14
from .errors import DocxMcpError, ErrCode


def _para_text(para: etree._Element) -> str:
    return "".join(t.text for t in para.iter(f"{W}t") if t.text)


def _ins_text(ins: etree._Element) -> str:
    return "".join(t.text for t in ins.iter(f"{W}t") if t.text)


def _del_text(del_el: etree._Element) -> str:
    return "".join(dt.text for dt in del_el.iter(f"{W}delText") if dt.text)


def _preceding_text(para: etree._Element, paras: list[etree._Element]) -> str:
    idx = paras.index(para)
    if idx == 0:
        return ""
    return _para_text(paras[idx - 1])[:40]


def _extract_changes(doc_xml: bytes) -> list[dict]:
    """Parse reviewer document XML and return list of change descriptors."""
    root = etree.fromstring(doc_xml)
    paras = list(root.iter(f"{W}p"))
    changes = []
    for para in paras:
        preceding = _preceding_text(para, paras)
        for child in para:
            tag = child.tag
            if tag == f"{W}ins":
                text = _ins_text(child)
                changes.append({
                    "kind": "ins",
                    "preceding": preceding,
                    "text": text,
                    "author": child.get(f"{W}author", ""),
                    "date": child.get(f"{W}date", ""),
                    "element": child,
                    "fingerprint": ("ins", preceding, text),
                })
            elif tag == f"{W}del":
                text = _del_text(child)
                changes.append({
                    "kind": "del",
                    "preceding": preceding,
                    "text": text,
                    "author": child.get(f"{W}author", ""),
                    "date": child.get(f"{W}date", ""),
                    "element": child,
                    "fingerprint": ("del", preceding, text),
                })
    return changes


class ReviewMergeMixin:
    def merge_review_rounds(
        self,
        reviewer_paths: list[str],
        base_path: str | None = None,
    ) -> dict:
        """Merge tracked changes from N reviewer copies into the open document.

        Algorithm:
        1. For each reviewer doc, extract w:ins/w:del with author+date
        2. Deduplicate identical changes (same text + position)
        3. Merge non-conflicting changes into open doc
        4. Flag conflicts (same range, different content) for manual resolution

        Returns: {"merged": int, "conflicts": list[dict], "skipped_duplicates": int}
        Raises: DocxMcpError(ErrCode.PART_NOT_FOUND) if a reviewer path doesn't exist.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        all_changes: list[dict] = []
        for rpath in reviewer_paths:
            import os
            if not os.path.exists(rpath):
                raise DocxMcpError(
                    ErrCode.PART_NOT_FOUND,
                    f"Reviewer document not found: {rpath}",
                )
            with zipfile.ZipFile(rpath, "r") as zf:
                xml_bytes = zf.read("word/document.xml")
            all_changes.extend(_extract_changes(xml_bytes))

        # Deduplicate and detect conflicts
        seen_fps: set[tuple] = set()
        range_texts: dict[tuple, str] = {}  # (kind, preceding) → first accepted text
        unique: list[dict] = []
        skipped = 0
        conflicts: list[dict] = []

        for ch in all_changes:
            fp = ch["fingerprint"]
            range_key = (fp[0], fp[1])  # (kind, preceding)
            if fp in seen_fps:
                skipped += 1
                continue
            first_text = range_texts.get(range_key)
            if first_text is not None and first_text != ch["text"]:
                conflicts.append({
                    "kind": ch["kind"],
                    "preceding": ch["preceding"],
                    "text": ch["text"],
                    "author": ch["author"],
                    "date": ch["date"],
                })
            else:
                seen_fps.add(fp)
                range_texts[range_key] = ch["text"]
                unique.append(ch)

        # Merge unique non-conflicting changes into open doc
        doc_paras = list(doc.iter(f"{W}p"))
        merged = 0
        next_id = self._next_markup_id(doc)

        for ch in unique:
            # Find matching paragraph by preceding text
            target_para = None
            for i, para in enumerate(doc_paras):
                pre = ""
                if i > 0:
                    pre = _para_text(doc_paras[i - 1])[:40]
                if pre == ch["preceding"]:
                    target_para = para
                    break

            if target_para is None:
                # Fallback: append to last paragraph
                target_para = doc_paras[-1] if doc_paras else None

            if target_para is None:
                continue

            el = copy.deepcopy(ch["element"])
            el.set(f"{W}id", str(next_id))
            next_id += 1
            target_para.append(el)
            merged += 1

        if merged:
            self._mark("word/document.xml")

        return {
            "merged": merged,
            "conflicts": conflicts,
            "skipped_duplicates": skipped,
        }
