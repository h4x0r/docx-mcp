"""DOCX document model with cached DOM for efficient MCP tool operations.

Opens a .docx (ZIP archive of XML), caches parsed lxml trees in memory,
and provides high-level methods for editing with track changes, comments,
footnotes, and structural validation.
"""

from __future__ import annotations

import contextlib
import os
import random
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

# ── OOXML namespace constants ───────────────────────────────────────────────
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

REL_TYPES = {
    "comments": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "commentsExtended": "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "footnotes": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
}

CT_TYPES = {
    "comments": ("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
    "commentsExtended": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    ),
}


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _preserve(t_el: etree._Element, text: str) -> None:
    """Set text on a <w:t> or <w:delText> element with xml:space=preserve."""
    t_el.text = text
    t_el.set(XML_SPACE, "preserve")


class DocxDocument:
    """Represents an open DOCX document with cached XML trees."""

    def __init__(self, path: str):
        self.source_path = Path(path).resolve()
        self.workdir: Path | None = None
        self._trees: dict[str, etree._Element] = {}
        self._modified: set[str] = set()

    # ── Open / Close ────────────────────────────────────────────────────────

    def open(self) -> dict:
        """Unpack DOCX and parse XML files. Returns document info."""
        if not self.source_path.exists():
            raise FileNotFoundError(f"File not found: {self.source_path}")
        if self.source_path.suffix.lower() != ".docx":
            raise ValueError(f"Not a .docx file: {self.source_path}")

        self.workdir = Path(tempfile.mkdtemp(prefix="docx_mcp_"))
        with zipfile.ZipFile(self.source_path, "r") as zf:
            zf.extractall(self.workdir)

        # Discover and parse XML files
        xml_files = ["[Content_Types].xml"]
        word_dir = self.workdir / "word"
        if word_dir.exists():
            for name in [
                "document.xml",
                "footnotes.xml",
                "endnotes.xml",
                "comments.xml",
                "commentsExtended.xml",
                "styles.xml",
                "numbering.xml",
                "settings.xml",
            ]:
                if (word_dir / name).exists():
                    xml_files.append(f"word/{name}")
            # Headers and footers
            for f in word_dir.iterdir():
                if f.name.startswith(("header", "footer")) and f.suffix == ".xml":
                    xml_files.append(f"word/{f.name}")

        # Relationship files
        rels_dir = word_dir / "_rels"
        if rels_dir.exists():
            for f in rels_dir.iterdir():
                if f.suffix == ".rels":
                    xml_files.append(f"word/_rels/{f.name}")

        for rel_path in xml_files:
            full_path = self.workdir / rel_path
            if full_path.exists():
                try:
                    parser = etree.XMLParser(remove_blank_text=False)
                    tree = etree.parse(str(full_path), parser)
                    self._trees[rel_path] = tree.getroot()
                except etree.XMLSyntaxError:
                    pass

        return self.get_info()

    def close(self) -> None:
        """Clean up temporary files."""
        if self.workdir and self.workdir.exists():
            shutil.rmtree(self.workdir, ignore_errors=True)
        self._trees.clear()
        self._modified.clear()
        self.workdir = None

    # ── Info ────────────────────────────────────────────────────────────────

    def get_info(self) -> dict:
        """Get document overview stats."""
        info: dict = {"path": str(self.source_path)}
        with contextlib.suppress(OSError):
            info["size_bytes"] = self.source_path.stat().st_size

        doc = self._tree("word/document.xml")
        if doc is not None:
            body = doc.find(f"{W}body")
            paras = list((body if body is not None else doc).iter(f"{W}p"))
            info["paragraph_count"] = len(paras)
            info["heading_count"] = len(self._find_headings(doc))
            info["image_count"] = len(list(doc.iter(f"{W}drawing")))

        fn = self._tree("word/footnotes.xml")
        if fn is not None:
            info["footnote_count"] = len(self._real_footnotes(fn))

        cm = self._tree("word/comments.xml")
        if cm is not None:
            info["comment_count"] = len(cm.findall(f"{W}comment"))

        info["parts"] = sorted(self._trees.keys())
        return info

    # ── Headings ────────────────────────────────────────────────────────────

    def get_headings(self) -> list[dict]:
        doc = self._require("word/document.xml")
        return self._find_headings(doc)

    def _find_headings(self, root: etree._Element) -> list[dict]:
        headings = []
        for para in root.iter(f"{W}p"):
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                continue
            pstyle = ppr.find(f"{W}pStyle")
            if pstyle is None:
                continue
            style = pstyle.get(f"{W}val", "")
            m = re.match(r"^Heading(\d+)$", style)
            if not m:
                continue
            headings.append(
                {
                    "level": int(m.group(1)),
                    "text": self._text(para),
                    "style": style,
                    "paraId": para.get(f"{W14}paraId", ""),
                }
            )
        return headings

    # ── Search ──────────────────────────────────────────────────────────────

    def search_text(self, query: str, *, regex: bool = False) -> list[dict]:
        """Search for text across document body, footnotes, and comments."""
        results = []
        targets = [
            ("document", "word/document.xml"),
            ("footnotes", "word/footnotes.xml"),
            ("comments", "word/comments.xml"),
        ]
        for source, rel_path in targets:
            tree = self._tree(rel_path)
            if tree is None:
                continue
            for para in tree.iter(f"{W}p"):
                text = self._text(para)
                if not text:
                    continue
                if regex:
                    matches = list(re.finditer(query, text))
                    if not matches:
                        continue
                    match_info = [
                        {"start": m.start(), "end": m.end(), "match": m.group()} for m in matches
                    ]
                else:
                    if query.lower() not in text.lower():
                        continue
                    match_info = None
                results.append(
                    {
                        "source": source,
                        "paraId": para.get(f"{W14}paraId", ""),
                        "text": text[:300],
                        "matches": match_info,
                    }
                )
        return results

    # ── Paragraph access ────────────────────────────────────────────────────

    def get_paragraph(self, para_id: str) -> dict:
        """Get full text and metadata for a paragraph by paraId."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        ppr = para.find(f"{W}pPr")
        style = ""
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                style = ps.get(f"{W}val", "")
        return {
            "paraId": para_id,
            "style": style,
            "text": self._text(para),
        }

    # ── Footnotes ───────────────────────────────────────────────────────────

    def get_footnotes(self) -> list[dict]:
        fn_tree = self._tree("word/footnotes.xml")
        if fn_tree is None:
            return []
        result = []
        for fn in self._real_footnotes(fn_tree):
            result.append(
                {
                    "id": int(fn.get(f"{W}id", "0")),
                    "text": self._text(fn),
                }
            )
        return result

    def add_footnote(self, para_id: str, text: str) -> dict:
        """Add a footnote to a paragraph. Returns the new footnote ID."""
        doc = self._require("word/document.xml")
        fn_tree = self._require("word/footnotes.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # Next ID
        existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        next_id = max(existing | {0}) + 1

        # Build footnote in footnotes.xml
        fn_el = etree.SubElement(fn_tree, f"{W}footnote")
        fn_el.set(f"{W}id", str(next_id))

        fn_para = etree.SubElement(fn_el, f"{W}p")
        fn_para.set(f"{W14}paraId", self._new_para_id())
        fn_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(fn_para, f"{W}pPr")
        ps = etree.SubElement(ppr, f"{W}pStyle")
        ps.set(f"{W}val", "FootnoteText")

        # Footnote ref mark
        ref_run = etree.SubElement(fn_para, f"{W}r")
        ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
        ref_style = etree.SubElement(ref_rpr, f"{W}rStyle")
        ref_style.set(f"{W}val", "FootnoteReference")
        etree.SubElement(ref_run, f"{W}footnoteRef")

        # Space
        sp_run = etree.SubElement(fn_para, f"{W}r")
        sp_t = etree.SubElement(sp_run, f"{W}t")
        _preserve(sp_t, " ")

        # Text
        txt_run = etree.SubElement(fn_para, f"{W}r")
        txt_t = etree.SubElement(txt_run, f"{W}t")
        _preserve(txt_t, text)

        self._mark("word/footnotes.xml")

        # Add reference in document paragraph
        r = etree.SubElement(para, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "FootnoteReference")
        fref = etree.SubElement(r, f"{W}footnoteReference")
        fref.set(f"{W}id", str(next_id))
        self._mark("word/document.xml")

        return {"footnote_id": next_id, "para_id": para_id}

    def validate_footnotes(self) -> dict:
        """Cross-reference footnote IDs between document.xml and footnotes.xml."""
        doc = self._tree("word/document.xml")
        fn_tree = self._tree("word/footnotes.xml")
        if doc is None:
            return {"error": "No document open"}
        if fn_tree is None:
            return {"valid": True, "references": 0, "definitions": 0}

        ref_ids = set()
        for ref in doc.iter(f"{W}footnoteReference"):
            fid = ref.get(f"{W}id")
            if fid:
                ref_ids.add(int(fid))

        def_ids = {int(f.get(f"{W}id", "0")) for f in self._real_footnotes(fn_tree)}

        missing = sorted(ref_ids - def_ids)
        orphans = sorted(def_ids - ref_ids)
        return {
            "valid": not missing and not orphans,
            "references": len(ref_ids),
            "definitions": len(def_ids),
            "missing_definitions": missing,
            "orphan_definitions": orphans,
        }

    # ── ParaId validation ───────────────────────────────────────────────────

    def validate_paraids(self) -> dict:
        """Check paraId uniqueness across all document parts."""
        all_ids: dict[str, list[str]] = {}
        for rel_path, tree in self._trees.items():
            if not rel_path.endswith(".xml"):
                continue
            for elem in tree.iter():
                pid = elem.get(f"{W14}paraId")
                if pid:
                    all_ids.setdefault(pid, []).append(rel_path)

        duplicates = {k: v for k, v in all_ids.items() if len(v) > 1}
        invalid = []
        for pid in all_ids:
            try:
                if int(pid, 16) >= 0x80000000:
                    invalid.append(pid)
            except ValueError:
                invalid.append(pid)

        return {
            "valid": not duplicates and not invalid,
            "total": len(all_ids),
            "duplicates": duplicates,
            "out_of_range": invalid,
        }

    # ── Watermark removal ───────────────────────────────────────────────────

    def remove_watermark(self) -> dict:
        """Remove VML watermarks (e.g., DRAFT) from all header XML files."""
        removed = []
        for rel_path, tree in self._trees.items():
            if "header" not in rel_path:
                continue
            for para in list(tree.iter(f"{W}p")):
                for run in list(para.findall(f"{W}r")):
                    pict = run.find(f"{W}pict")
                    if pict is None:
                        continue
                    for shape in pict.iter(f"{V}shape"):
                        tp = shape.find(f"{V}textpath")
                        if tp is not None:
                            wm_text = tp.get("string", "")
                            para.remove(run)
                            removed.append({"header": rel_path, "text": wm_text})
                            self._mark(rel_path)
                            break
        return {"removed": len(removed), "details": removed}

    # ── Track changes ───────────────────────────────────────────────────────

    def insert_text(
        self,
        para_id: str,
        text: str,
        *,
        position: str = "end",
        author: str = "Claude",
    ) -> dict:
        """Insert text with Word track-changes markup (w:ins).

        Args:
            para_id: Target paragraph paraId.
            text: Text to insert.
            position: 'end', 'start', or a substring after which to insert.
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        ins = etree.Element(f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        r = etree.SubElement(ins, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        if position == "start":
            ppr = para.find(f"{W}pPr")
            if ppr is not None:
                ppr.addnext(ins)
            else:
                para.insert(0, ins)
        elif position == "end":
            para.append(ins)
        else:
            placed = False
            for run_el in para.findall(f"{W}r"):
                if position in self._text(run_el):
                    run_el.addnext(ins)
                    placed = True
                    break
            if not placed:
                para.append(ins)

        self._mark("word/document.xml")
        return {"change_id": cid, "type": "insertion", "author": author, "date": now}

    def delete_text(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Mark text as deleted with Word track-changes markup (w:del).

        Finds the text within runs of the target paragraph and wraps the
        matching portion in <w:del><w:r><w:delText>...</w:delText></w:r></w:del>,
        splitting the run if the match is a substring.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            full = t_el.text
            if text not in full:
                continue

            idx = full.index(text)
            rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(rpr) if rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)

            insert_at = pos

            # Text before
            if idx > 0:
                before = self._make_run(full[:idx], rpr_bytes)
                parent.insert(insert_at, before)
                insert_at += 1

            # Deletion
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, text)
            parent.insert(insert_at, del_el)
            insert_at += 1

            # Text after
            end = idx + len(text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark("word/document.xml")
            return {"change_id": cid, "type": "deletion", "author": author, "date": now}

        raise ValueError(
            f"Text '{text}' not found in a single run of paragraph '{para_id}'. "
            "If the text spans multiple runs, try searching for a shorter substring."
        )

    # ── Comments ────────────────────────────────────────────────────────────

    def get_comments(self) -> list[dict]:
        cm = self._tree("word/comments.xml")
        if cm is None:
            return []
        return [
            {
                "id": int(c.get(f"{W}id", "0")),
                "author": c.get(f"{W}author", ""),
                "date": c.get(f"{W}date", ""),
                "text": self._text(c),
            }
            for c in cm.findall(f"{W}comment")
        ]

    def add_comment(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Add a comment anchored to a paragraph."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cm_tree = self._tree("word/comments.xml")
        if cm_tree is None:
            cm_tree = self._create_comments_part()

        comment_id = self._next_comment_id(cm_tree)
        now = _now_iso()
        initials = "".join(w[0].upper() for w in author.split() if w) or "C"

        # Add to comments.xml
        c = etree.SubElement(cm_tree, f"{W}comment")
        c.set(f"{W}id", str(comment_id))
        c.set(f"{W}author", author)
        c.set(f"{W}date", now)
        c.set(f"{W}initials", initials)

        cp = etree.SubElement(c, f"{W}p")
        cp.set(f"{W14}paraId", self._new_para_id())
        cp.set(f"{W14}textId", "77777777")

        # Annotation ref
        ar_run = etree.SubElement(cp, f"{W}r")
        ar_rpr = etree.SubElement(ar_run, f"{W}rPr")
        ar_rs = etree.SubElement(ar_rpr, f"{W}rStyle")
        ar_rs.set(f"{W}val", "CommentReference")
        etree.SubElement(ar_run, f"{W}annotationRef")

        # Comment text
        t_run = etree.SubElement(cp, f"{W}r")
        t_el = etree.SubElement(t_run, f"{W}t")
        _preserve(t_el, text)
        self._mark("word/comments.xml")

        # Add range markers in document.xml
        range_start = etree.Element(f"{W}commentRangeStart")
        range_start.set(f"{W}id", str(comment_id))

        ppr = para.find(f"{W}pPr")
        first_run = para.find(f"{W}r")
        if first_run is not None:
            first_run.addprevious(range_start)
        elif ppr is not None:
            ppr.addnext(range_start)
        else:
            para.insert(0, range_start)

        range_end = etree.SubElement(para, f"{W}commentRangeEnd")
        range_end.set(f"{W}id", str(comment_id))

        ref_run = etree.SubElement(para, f"{W}r")
        ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
        ref_rs = etree.SubElement(ref_rpr, f"{W}rStyle")
        ref_rs.set(f"{W}val", "CommentReference")
        cref = etree.SubElement(ref_run, f"{W}commentReference")
        cref.set(f"{W}id", str(comment_id))
        self._mark("word/document.xml")

        return {"comment_id": comment_id, "para_id": para_id, "author": author, "date": now}

    def reply_to_comment(
        self,
        parent_id: int,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Reply to an existing comment."""
        cm_tree = self._require("word/comments.xml")

        # Verify parent exists
        parent_el = None
        for c in cm_tree.findall(f"{W}comment"):
            if c.get(f"{W}id") == str(parent_id):
                parent_el = c
                break
        if parent_el is None:
            raise ValueError(f"Comment {parent_id} not found")

        comment_id = self._next_comment_id(cm_tree)
        now = _now_iso()
        initials = "".join(w[0].upper() for w in author.split() if w) or "C"

        reply = etree.SubElement(cm_tree, f"{W}comment")
        reply.set(f"{W}id", str(comment_id))
        reply.set(f"{W}author", author)
        reply.set(f"{W}date", now)
        reply.set(f"{W}initials", initials)

        rp = etree.SubElement(reply, f"{W}p")
        reply_para_id = self._new_para_id()
        rp.set(f"{W14}paraId", reply_para_id)
        rp.set(f"{W14}textId", "77777777")

        t_run = etree.SubElement(rp, f"{W}r")
        t_el = etree.SubElement(t_run, f"{W}t")
        _preserve(t_el, text)
        self._mark("word/comments.xml")

        # Thread via commentsExtended.xml
        ext = self._tree("word/commentsExtended.xml")
        if ext is not None:
            parent_para = parent_el.find(f"{W}p")
            parent_para_id = parent_para.get(f"{W14}paraId", "") if parent_para is not None else ""
            ce = etree.SubElement(ext, f"{W15}commentEx")
            ce.set(f"{W15}paraId", reply_para_id)
            ce.set(f"{W15}paraIdParent", parent_para_id)
            ce.set(f"{W15}done", "0")
            self._mark("word/commentsExtended.xml")

        return {
            "comment_id": comment_id,
            "parent_id": parent_id,
            "author": author,
            "date": now,
        }

    # ── Audit ───────────────────────────────────────────────────────────────

    def audit(self) -> dict:
        """Run comprehensive structural validation."""
        results: dict = {}

        results["footnotes"] = self.validate_footnotes()
        results["paraids"] = self.validate_paraids()

        # Headings
        doc = self._tree("word/document.xml")
        if doc is not None:
            headings = self._find_headings(doc)
            issues = []
            prev = 0
            for h in headings:
                if h["level"] > prev + 1 and prev > 0:
                    issues.append(
                        {
                            "issue": "level_skip",
                            "heading": h["text"][:60],
                            "expected_max": prev + 1,
                            "actual": h["level"],
                        }
                    )
                prev = h["level"]
            results["headings"] = {"count": len(headings), "issues": issues}

            # Bookmarks
            starts = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkStart") if e.get(f"{W}id")}
            ends = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkEnd") if e.get(f"{W}id")}
            results["bookmarks"] = {
                "total": len(starts),
                "unpaired_starts": len(starts - ends),
                "unpaired_ends": len(ends - starts),
            }

        # Relationships — check targets exist
        rels_tree = self._tree("word/_rels/document.xml.rels")
        rel_issues = []
        if rels_tree is not None:
            for rel in rels_tree.findall(f"{RELS}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                if not (self.workdir / "word" / target).exists():
                    rel_issues.append({"id": rel.get("Id"), "target": target})
        results["relationships"] = {"missing_targets": rel_issues}

        # Images
        img_issues = []
        if doc is not None and rels_tree is not None:
            for blip in doc.iter(f"{A}blip"):
                embed = blip.get(f"{R}embed")
                if not embed:
                    continue
                rel = rels_tree.find(f'{RELS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    target = rel.get("Target", "")
                    if not (self.workdir / "word" / target).exists():
                        img_issues.append({"rId": embed, "target": target})
        results["images"] = {"missing": img_issues}

        # Artifacts
        artifacts = []
        for marker in ["DRAFT", "TODO", "FIXME", "XXX"]:
            for hit in self.search_text(marker):
                artifacts.append(
                    {"marker": marker, "source": hit["source"], "context": hit["text"][:100]}
                )
        results["artifacts"] = artifacts

        # Overall
        results["valid"] = (
            results["footnotes"].get("valid", True)
            and results["paraids"].get("valid", True)
            and not results["headings"].get("issues")
            and results["bookmarks"].get("unpaired_starts", 0) == 0
            and results["bookmarks"].get("unpaired_ends", 0) == 0
            and not rel_issues
            and not img_issues
        )
        return results

    # ── Save ────────────────────────────────────────────────────────────────

    def save(self, output_path: str | None = None) -> dict:
        """Write modified XML back to files and repack into a .docx."""
        if self.workdir is None:
            raise RuntimeError("No document is open")

        output = Path(output_path) if output_path else self.source_path

        # Serialize modified trees
        for rel_path in self._modified:
            tree = self._trees.get(rel_path)
            if tree is None:
                continue
            fp = self.workdir / rel_path
            fp.parent.mkdir(parents=True, exist_ok=True)
            et = etree.ElementTree(tree)
            et.write(
                str(fp),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        # Repack
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(self.workdir):
                for fname in files:
                    fpath = Path(root) / fname
                    arcname = str(fpath.relative_to(self.workdir))
                    zf.write(fpath, arcname)

        modified = sorted(self._modified)
        self._modified.clear()
        return {
            "path": str(output),
            "size_bytes": output.stat().st_size,
            "modified_parts": modified,
        }

    # ── Private helpers ─────────────────────────────────────────────────────

    def _tree(self, rel_path: str) -> etree._Element | None:
        return self._trees.get(rel_path)

    def _require(self, rel_path: str) -> etree._Element:
        t = self._tree(rel_path)
        if t is None:
            raise RuntimeError(f"{rel_path} not found — is a document open?")
        return t

    def _mark(self, rel_path: str) -> None:
        self._modified.add(rel_path)

    @staticmethod
    def _text(element: etree._Element) -> str:
        """Concatenate all <w:t> text descendants."""
        return "".join(t.text for t in element.iter(f"{W}t") if t.text)

    @staticmethod
    def _real_footnotes(fn_root: etree._Element) -> list[etree._Element]:
        """Return footnote elements excluding separators (id 0 and -1)."""
        return [f for f in fn_root.findall(f"{W}footnote") if f.get(f"{W}id") not in ("0", "-1")]

    def _find_para(self, root: etree._Element, para_id: str) -> etree._Element | None:
        for p in root.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                return p
        return None

    def _new_para_id(self) -> str:
        """Generate a unique paraId (8 hex digits, < 0x80000000)."""
        existing: set[str] = set()
        for tree in self._trees.values():
            for el in tree.iter():
                pid = el.get(f"{W14}paraId")
                if pid:
                    existing.add(pid.upper())
        while True:
            val = random.randint(1, 0x7FFFFFFF)
            pid = f"{val:08X}"
            if pid not in existing:
                return pid

    @staticmethod
    def _next_markup_id(doc: etree._Element) -> int:
        """Next available ID for ins/del/comment/bookmark markup."""
        max_id = 0
        for tag in (
            f"{W}ins",
            f"{W}del",
            f"{W}commentRangeStart",
            f"{W}commentRangeEnd",
            f"{W}bookmarkStart",
            f"{W}bookmarkEnd",
        ):
            for el in doc.iter(tag):
                eid = el.get(f"{W}id")
                if eid:
                    with contextlib.suppress(ValueError):
                        max_id = max(max_id, int(eid))
        return max_id + 1

    @staticmethod
    def _next_comment_id(cm_tree: etree._Element) -> int:
        max_id = -1
        for c in cm_tree.findall(f"{W}comment"):
            cid = c.get(f"{W}id")
            if cid:
                with contextlib.suppress(ValueError):
                    max_id = max(max_id, int(cid))
        return max_id + 1

    @staticmethod
    def _make_run(text: str, rpr_bytes: bytes | None) -> etree._Element:
        """Build a <w:r> element with optional copied rPr."""
        r = etree.Element(f"{W}r")
        if rpr_bytes:
            r.append(etree.fromstring(rpr_bytes))
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)
        return r

    def _create_comments_part(self) -> etree._Element:
        """Create word/comments.xml and register it in rels + content types."""
        root = etree.Element(
            f"{W}comments",
            nsmap={"w": NSMAP["w"], "w14": NSMAP["w14"], "r": NSMAP["r"]},
        )
        self._trees["word/comments.xml"] = root

        # Write file so it exists on disk
        fp = self.workdir / "word" / "comments.xml"
        fp.parent.mkdir(parents=True, exist_ok=True)
        etree.ElementTree(root).write(str(fp), xml_declaration=True, encoding="UTF-8")

        # Content type
        ct = self._tree("[Content_Types].xml")
        if ct is not None:
            existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
            if "/word/comments.xml" not in existing:
                ov = etree.SubElement(ct, f"{CT}Override")
                ov.set("PartName", "/word/comments.xml")
                ov.set("ContentType", CT_TYPES["comments"])
                self._mark("[Content_Types].xml")

        # Relationship
        rels = self._tree("word/_rels/document.xml.rels")
        if rels is not None:
            existing_targets = {r.get("Target") for r in rels.findall(f"{RELS}Relationship")}
            if "comments.xml" not in existing_targets:
                max_rid = 0
                for r in rels.findall(f"{RELS}Relationship"):
                    rid = r.get("Id", "")
                    if rid.startswith("rId"):
                        with contextlib.suppress(ValueError):
                            max_rid = max(max_rid, int(rid[3:]))
                rel = etree.SubElement(rels, f"{RELS}Relationship")
                rel.set("Id", f"rId{max_rid + 1}")
                rel.set("Type", REL_TYPES["comments"])
                rel.set("Target", "comments.xml")
                self._mark("word/_rels/document.xml.rels")

        self._mark("word/comments.xml")
        return root
