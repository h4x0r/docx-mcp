"""Comments mixin: get, add, reply."""

from __future__ import annotations

from lxml import etree

from .base import W14, W15, W, _now_iso, _preserve


class CommentsMixin:
    """Comment operations."""

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

        # Thread via commentsExtended.xml — create if needed
        ext = self._tree("word/commentsExtended.xml")
        if ext is None:
            ext = self._create_comments_extended_part()

        parent_para = parent_el.find(f"{W}p")
        parent_para_id = parent_para.get(f"{W14}paraId", "") if parent_para is not None else ""

        # Add commentEx for parent if not already present
        existing_para_ids = {ce.get(f"{W15}paraId") for ce in ext.iter(f"{W15}commentEx")}
        if parent_para_id and parent_para_id not in existing_para_ids:
            parent_ce = etree.SubElement(ext, f"{W15}commentEx")
            parent_ce.set(f"{W15}paraId", parent_para_id)
            parent_ce.set(f"{W15}done", "0")

        # Add commentEx for reply
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

    def _create_comments_extended_part(self) -> etree._Element:
        """Create word/commentsExtended.xml and register in rels + content types."""
        from .base import CT, CT_TYPES, NSMAP, REL_TYPES, RELS

        root = etree.Element(
            f"{W15}commentsEx",
            nsmap={"w15": NSMAP["w15"], "w14": NSMAP["w14"]},
        )
        self._trees["word/commentsExtended.xml"] = root

        fp = self.workdir / "word" / "commentsExtended.xml"
        fp.parent.mkdir(parents=True, exist_ok=True)
        etree.ElementTree(root).write(str(fp), xml_declaration=True, encoding="UTF-8")

        ct = self._tree("[Content_Types].xml")
        if ct is not None:
            existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
            if "/word/commentsExtended.xml" not in existing:
                ov = etree.SubElement(ct, f"{CT}Override")
                ov.set("PartName", "/word/commentsExtended.xml")
                ov.set("ContentType", CT_TYPES["commentsExtended"])
                self._mark("[Content_Types].xml")

        rels = self._tree("word/_rels/document.xml.rels")
        if rels is not None:
            import contextlib

            existing_targets = {r.get("Target") for r in rels.findall(f"{RELS}Relationship")}
            if "commentsExtended.xml" not in existing_targets:
                max_rid = 0
                for r in rels.findall(f"{RELS}Relationship"):
                    rid = r.get("Id", "")
                    if rid.startswith("rId"):
                        with contextlib.suppress(ValueError):
                            max_rid = max(max_rid, int(rid[3:]))
                rel = etree.SubElement(rels, f"{RELS}Relationship")
                rel.set("Id", f"rId{max_rid + 1}")
                rel.set("Type", REL_TYPES["commentsExtended"])
                rel.set("Target", "commentsExtended.xml")
                self._mark("word/_rels/document.xml.rels")

        self._mark("word/commentsExtended.xml")
        return root

    def update_comment(self, comment_id: int, text: str) -> dict:
        """Replace the text of an existing comment."""
        cm_tree = self._tree("word/comments.xml")
        if cm_tree is None:
            raise ValueError(f"Comment {comment_id} not found")

        comment_el = None
        for c in cm_tree.findall(f"{W}comment"):
            if c.get(f"{W}id") == str(comment_id):
                comment_el = c
                break
        if comment_el is None:
            raise ValueError(f"Comment {comment_id} not found")

        # Replace text in first paragraph: remove all w:r runs, add new one
        first_para = comment_el.find(f"{W}p")
        if first_para is not None:
            for run in first_para.findall(f"{W}r"):
                # Keep annotation-ref run (contains w:annotationRef), remove text runs
                if run.find(f"{W}annotationRef") is None:
                    first_para.remove(run)
            new_run = etree.SubElement(first_para, f"{W}r")
            t_el = etree.SubElement(new_run, f"{W}t")
            _preserve(t_el, text)

        self._mark("word/comments.xml")
        return {"comment_id": comment_id, "text": text}

    def delete_comment(self, comment_id: int) -> dict:
        """Delete a comment and all its range markers from the document."""
        cm_tree = self._tree("word/comments.xml")
        if cm_tree is None:
            raise ValueError(f"Comment {comment_id} not found")

        comment_el = None
        for c in cm_tree.findall(f"{W}comment"):
            if c.get(f"{W}id") == str(comment_id):
                comment_el = c
                break
        if comment_el is None:
            raise ValueError(f"Comment {comment_id} not found")

        # Get paraId of comment's first paragraph for commentsExtended cleanup
        first_para = comment_el.find(f"{W}p")
        comment_para_id = first_para.get(f"{W14}paraId") if first_para is not None else None

        # Remove from comments.xml
        cm_tree.remove(comment_el)
        self._mark("word/comments.xml")

        # Remove range markers from document.xml
        doc = self._tree("word/document.xml")
        if doc is not None:
            cid_str = str(comment_id)
            for tag in (
                f"{W}commentRangeStart",
                f"{W}commentRangeEnd",
            ):
                for el in list(doc.iter(tag)):
                    if el.get(f"{W}id") == cid_str:
                        el.getparent().remove(el)
            # commentReference is inside a w:r — remove the whole run
            for run in list(doc.iter(f"{W}r")):
                cref = run.find(f"{W}commentReference")
                if cref is not None and cref.get(f"{W}id") == cid_str:
                    run.getparent().remove(run)
            self._mark("word/document.xml")

        # Remove from commentsExtended.xml if present
        ext = self._tree("word/commentsExtended.xml")
        if ext is not None and comment_para_id:
            for ce in list(ext.iter(f"{W15}commentEx")):
                if ce.get(f"{W15}paraId") == comment_para_id:
                    ce.getparent().remove(ce)
                    self._mark("word/commentsExtended.xml")
                    break

        return {"deleted": comment_id}

    def resolve_comment(self, comment_id: int) -> dict:
        """Mark a comment as resolved in commentsExtended.xml (w15:done='1')."""
        cm_tree = self._require("word/comments.xml")

        comment_el = None
        for c in cm_tree.findall(f"{W}comment"):
            if c.get(f"{W}id") == str(comment_id):
                comment_el = c
                break
        if comment_el is None:
            raise ValueError(f"Comment {comment_id} not found")

        first_para = comment_el.find(f"{W}p")
        para_id = first_para.get(f"{W14}paraId") if first_para is not None else None

        ext = self._tree("word/commentsExtended.xml")
        if ext is None or not para_id:
            return {"resolved": comment_id, "found_extended": False}

        found = False
        for ce in ext.iter(f"{W15}commentEx"):
            if ce.get(f"{W15}paraId") == para_id:
                ce.set(f"{W15}done", "1")
                self._mark("word/commentsExtended.xml")
                found = True
                break

        return {"resolved": comment_id, "found_extended": found}

    def list_comment_threads(self) -> list[dict]:
        """Return threaded comment structure using commentsExtended.xml."""
        cm_tree = self._tree("word/comments.xml")
        if cm_tree is None:
            return []

        # Build comment info dict keyed by comment_id, including paraId
        comments_info: dict[int, dict] = {}
        para_id_to_comment_id: dict[str, int] = {}

        for c in cm_tree.findall(f"{W}comment"):
            cid = int(c.get(f"{W}id", "0"))
            first_para = c.find(f"{W}p")
            para_id = first_para.get(f"{W14}paraId", "") if first_para is not None else ""
            info = {
                "id": cid,
                "author": c.get(f"{W}author", ""),
                "date": c.get(f"{W}date", ""),
                "text": self._text(c),
                "_para_id": para_id,
            }
            comments_info[cid] = info
            if para_id:
                para_id_to_comment_id[para_id] = cid

        if not comments_info:
            return []

        ext = self._tree("word/commentsExtended.xml")

        # Map: reply comment_id → parent comment_id (via paraIdParent)
        reply_to_parent: dict[int, int] = {}

        if ext is not None:
            # Build paraId → commentId map from commentsExtended
            for ce in ext.iter(f"{W15}commentEx"):
                para_id = ce.get(f"{W15}paraId", "")
                parent_para_id = ce.get(f"{W15}paraIdParent", "")
                if parent_para_id and para_id:
                    child_cid = para_id_to_comment_id.get(para_id)
                    parent_cid = para_id_to_comment_id.get(parent_para_id)
                    if child_cid is not None and parent_cid is not None:
                        reply_to_parent[child_cid] = parent_cid

        # Build threads
        threads: dict[int, dict] = {}
        reply_ids: set[int] = set(reply_to_parent.keys())

        for cid, info in comments_info.items():
            if cid not in reply_ids:
                threads[cid] = {
                    "root": {k: v for k, v in info.items() if not k.startswith("_")},
                    "replies": [],
                }

        for reply_cid, parent_cid in reply_to_parent.items():
            if parent_cid in threads and reply_cid in comments_info:
                reply_info = comments_info[reply_cid]
                threads[parent_cid]["replies"].append(
                    {k: v for k, v in reply_info.items() if not k.startswith("_")}
                )

        return list(threads.values())
