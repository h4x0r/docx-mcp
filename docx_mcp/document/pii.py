"""PII scrubbing mixin: detect and permanently redact personal information."""

from __future__ import annotations

import copy
import os
import zipfile
from pathlib import Path

from lxml import etree

from .base import W

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Lazy singleton — loading AnalyzerEngine instantiates spaCy (expensive once)
_analyzer = None


def _get_analyzer():
    global _analyzer
    if _analyzer is None:
        try:
            from presidio_analyzer import AnalyzerEngine
        except ImportError as exc:
            raise ImportError(
                "PII scrubbing requires presidio-analyzer and a spaCy model. "
                "Install with:\n"
                "  pip install presidio-analyzer presidio-anonymizer\n"
                "  python -m spacy download en_core_web_lg"
            ) from exc
        _analyzer = AnalyzerEngine()
    return _analyzer


def _wt(tag: str) -> str:
    return f"{{{_W}}}{tag}"


# ── Run-level helpers ────────────────────────────────────────────────────────


def _make_run(text: str, rPr: etree._Element | None) -> etree._Element:
    """Create a plain w:r with optional rPr (deep-copied)."""
    run = etree.Element(_wt("r"))
    if rPr is not None:
        run.append(copy.deepcopy(rPr))
    t_el = etree.SubElement(run, _wt("t"))
    t_el.set(_XML_SPACE, "preserve")
    t_el.text = text
    return run


def _make_redacted_run(text: str, original_rPr: etree._Element | None) -> etree._Element:
    """Create a w:r with black highlight + black color (redaction bar)."""
    run = etree.Element(_wt("r"))
    # Build rPr: start from copy of original, add/replace highlight + color
    rPr = copy.deepcopy(original_rPr) if original_rPr is not None else etree.Element(_wt("rPr"))

    # Remove any existing highlight or color so we can set them cleanly
    for tag in (_wt("highlight"), _wt("color")):
        existing = rPr.find(tag)
        if existing is not None:
            rPr.remove(existing)

    highlight = etree.SubElement(rPr, _wt("highlight"))
    highlight.set(_wt("val"), "black")
    color = etree.SubElement(rPr, _wt("color"))
    color.set(_wt("val"), "000000")
    run.append(rPr)

    t_el = etree.SubElement(run, _wt("t"))
    t_el.set(_XML_SPACE, "preserve")
    t_el.text = text
    return run


# ── Paragraph redaction ──────────────────────────────────────────────────────


def _build_run_char_map(
    para: etree._Element,
) -> tuple[str, list[tuple[int, int, etree._Element]]]:
    """
    Return (full_text, run_ranges) where:
      full_text  — concatenated w:t text (excluding w:del content)
      run_ranges — [(global_start, global_end, run_el), ...]

    Each run's character range in full_text is [global_start, global_end).
    Only w:r elements NOT inside w:del are included.
    """
    full_text = ""
    run_ranges: list[tuple[int, int, etree._Element]] = []

    for run in para.iter(_wt("r")):
        # Skip deleted runs
        parent = run.getparent()
        if parent is not None and parent.tag == _wt("del"):
            continue

        run_text = "".join(t.text or "" for t in run.findall(_wt("t")))
        if run_text:
            start = len(full_text)
            full_text += run_text
            run_ranges.append((start, len(full_text), run))

    return full_text, run_ranges


def _redact_span(
    run_ranges: list[tuple[int, int, etree._Element]],
    span_start: int,
    span_end: int,
    marker: str,
) -> None:
    """
    Replace the character range [span_start, span_end) in the paragraph
    with redaction markers.  Modifies the XML tree in-place.

    Handles spans that cross multiple runs by splitting each involved run.
    """
    involved = [(s, e, r) for s, e, r in run_ranges if e > span_start and s < span_end]
    if not involved:
        return

    for run_start, run_end, run in involved:
        run_text = "".join(t.text or "" for t in run.findall(_wt("t")))
        rPr = run.find(_wt("rPr"))

        # Portion of this run that falls within the PII span
        pii_in_run_start = max(span_start - run_start, 0)
        pii_in_run_end = min(span_end - run_start, run_end - run_start)

        before = run_text[:pii_in_run_start]
        pii_part = run_text[pii_in_run_start:pii_in_run_end]
        after = run_text[pii_in_run_end:]

        # Insert replacement runs immediately before original, then remove it
        if before:
            run.addprevious(_make_run(before, rPr))
        if pii_part:
            run.addprevious(_make_redacted_run(marker * len(pii_part), rPr))
        if after:
            run.addprevious(_make_run(after, rPr))

        run.getparent().remove(run)


def _find_all_occurrences(text: str, needle: str) -> list[tuple[int, int]]:
    """Return [(start, end), ...] for all case-insensitive occurrences of needle."""
    needle_lower = needle.casefold()
    text_lower = text.casefold()
    spans = []
    start = 0
    while True:
        pos = text_lower.find(needle_lower, start)
        if pos == -1:
            break
        spans.append((pos, pos + len(needle)))
        start = pos + 1
    return spans


def _merge_overlapping(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Merge overlapping or adjacent spans, return sorted non-overlapping list."""
    if not spans:
        return []
    sorted_spans = sorted(spans)
    merged = [sorted_spans[0]]
    for start, end in sorted_spans[1:]:
        if start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


# ── Main mixin ───────────────────────────────────────────────────────────────


class PiiMixin:
    """Detect and permanently redact PII using Presidio + spaCy NER."""

    def scrub_pii(
        self,
        output_path: str,
        *,
        entities: list[str] | None = None,
        confidence_threshold: float = 0.35,
        redaction_marker: str = "█",
        dry_run: bool = False,
        also_sanitize_metadata: bool = True,
        redact_authors_as: str = "REDACTED",
    ) -> dict:
        """Detect and permanently redact PII from the open document.

        Args:
            output_path: Destination path for the scrubbed DOCX.
                         Required when dry_run=False; ignored when dry_run=True.
            entities:    Presidio entity type filter, e.g. ["PERSON","EMAIL_ADDRESS"].
                         None / empty list = all supported types.
            confidence_threshold: Presidio score floor (default 0.35).
            redaction_marker:     Character repeated to replace each PII char (default "█").
            dry_run:     If True, detect only — return entity list, write no file.
            also_sanitize_metadata: When True (default), apply level-3 metadata
                                    sanitization on the output (clears creator,
                                    company, revision, attachedTemplate).
            redact_authors_as: Author label for metadata sanitization (level 2).

        Returns:
            {"path": str|None, "entities": [...]}
        """
        if not dry_run and not output_path:
            raise ValueError("output_path is required when dry_run=False")
        if self.workdir is None:
            raise RuntimeError("No document is open")

        analyzer = _get_analyzer()
        entity_filter = list(entities) if entities else None

        doc_tree = self._trees.get("word/document.xml")
        if doc_tree is None:
            return {"path": None if dry_run else output_path, "entities": []}

        # ── Phase 1: detect PII per paragraph ───────────────────────────────
        body = doc_tree.find(_wt("body"))
        paragraphs = list(body) if body is not None else []

        detected: list[dict] = []  # {type, text, para_index, start, end, score}

        for para_idx, para in enumerate(paragraphs):
            if para.tag != _wt("p"):
                continue
            full_text, _ = _build_run_char_map(para)
            if not full_text.strip():
                continue

            results = analyzer.analyze(
                text=full_text,
                language="en",
                entities=entity_filter,
                score_threshold=confidence_threshold,
            )
            for r in results:
                detected.append({
                    "type": r.entity_type,
                    "text": full_text[r.start:r.end],
                    "para_index": para_idx,
                    "start": r.start,
                    "end": r.end,
                    "score": round(r.score, 3),
                })

        if dry_run:
            return {"path": None, "entities": detected}

        # ── Phase 2: deduplication ───────────────────────────────────────────
        # Collect unique PII strings, then find ALL occurrences across the doc
        unique_pii_texts = {d["text"] for d in detected}

        # Build redaction spans per paragraph (index → [(start, end)])
        redaction_map: dict[int, list[tuple[int, int]]] = {}

        for para_idx, para in enumerate(paragraphs):
            if para.tag != _wt("p"):
                continue
            full_text, _ = _build_run_char_map(para)
            if not full_text:
                continue

            spans: list[tuple[int, int]] = []
            for pii_text in unique_pii_texts:
                spans.extend(_find_all_occurrences(full_text, pii_text))

            if spans:
                redaction_map[para_idx] = _merge_overlapping(spans)

        # ── Phase 3: redact on a deep copy of the document tree ─────────────
        doc_copy = copy.deepcopy(doc_tree)
        body_copy = doc_copy.find(_wt("body"))
        copy_paragraphs = list(body_copy) if body_copy is not None else []

        for para_idx, para in enumerate(copy_paragraphs):
            if para.tag != _wt("p"):
                continue
            if para_idx not in redaction_map:
                continue

            full_text, run_ranges = _build_run_char_map(para)
            # Redact spans in reverse order to preserve offsets
            for span_start, span_end in reversed(redaction_map[para_idx]):
                _redact_span(run_ranges, span_start, span_end, redaction_marker)
                # Rebuild run_ranges after each modification (indices shift)
                full_text, run_ranges = _build_run_char_map(para)

        out_doc_bytes = etree.tostring(
            doc_copy, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # ── Phase 4: write output zip ────────────────────────────────────────
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for root_dir, _dirs, files in os.walk(self.workdir):
                for fname in sorted(files):
                    fpath = Path(root_dir) / fname
                    arcname = str(fpath.relative_to(self.workdir))

                    if arcname == "word/document.xml":
                        out_zip.writestr(arcname, out_doc_bytes)
                    elif also_sanitize_metadata and arcname in (
                        "word/settings.xml",
                        "docProps/core.xml",
                    ):
                        if arcname in self._trees:
                            el = copy.deepcopy(self._trees[arcname])
                            if arcname == "word/settings.xml":
                                self._sanitize_settings_el(el)
                            else:
                                self._sanitize_core_el(el)
                            out_zip.writestr(
                                arcname,
                                etree.tostring(
                                    el, xml_declaration=True, encoding="UTF-8", standalone=True
                                ),
                            )
                        else:
                            out_zip.writestr(arcname, fpath.read_bytes())
                    elif also_sanitize_metadata and arcname == "docProps/app.xml":
                        raw = fpath.read_bytes()
                        app_el = etree.fromstring(raw)
                        self._sanitize_app_el(app_el)
                        out_zip.writestr(
                            arcname,
                            etree.tostring(
                                app_el, xml_declaration=True, encoding="UTF-8", standalone=True
                            ),
                        )
                    elif arcname in self._trees:
                        out_zip.writestr(
                            arcname,
                            etree.tostring(
                                self._trees[arcname],
                                xml_declaration=True,
                                encoding="UTF-8",
                                standalone=True,
                            ),
                        )
                    else:
                        out_zip.writestr(arcname, fpath.read_bytes())

        return {"path": output_path, "entities": detected}
