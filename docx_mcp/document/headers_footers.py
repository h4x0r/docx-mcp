"""Headers/footers mixin: read and edit."""

from __future__ import annotations

from lxml import etree

from .base import W, _now_iso, _preserve


class HeadersFootersMixin:
    """Header and footer operations."""

    def get_headers_footers(self) -> list[dict]:
        """Get all headers and footers with their text content."""
        results = []
        for rel_path, tree in self._trees.items():
            if not rel_path.startswith("word/header") and not rel_path.startswith("word/footer"):
                continue
            location = "header" if "header" in rel_path else "footer"
            text = self._text(tree)
            results.append(
                {
                    "part": rel_path,
                    "location": location,
                    "text": text,
                }
            )
        return sorted(results, key=lambda x: x["part"])

    def edit_header_footer(
        self,
        location: str,
        old_text: str,
        new_text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Edit text in a header or footer with tracked changes.

        Args:
            location: "header" or "footer" (matches first found).
            old_text: Text to find and replace.
            new_text: Replacement text.
            author: Author name for the revision.
        """
        # Find matching part
        target_part = None
        for rel_path in self._trees:
            if rel_path.startswith(f"word/{location}"):
                target_part = rel_path
                break

        if target_part is None:
            raise ValueError(f"No {location} found in document")

        tree = self._trees[target_part]
        now = _now_iso()
        changes = 0

        for run_el in list(tree.iter(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            if old_text not in t_el.text:
                continue

            full = t_el.text
            idx = full.index(old_text)
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
            cid = self._next_markup_id(tree)
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, old_text)
            parent.insert(insert_at, del_el)
            insert_at += 1
            changes += 1

            # Insertion
            cid = self._next_markup_id(tree)
            ins_el = etree.Element(f"{W}ins")
            ins_el.set(f"{W}id", str(cid))
            ins_el.set(f"{W}author", author)
            ins_el.set(f"{W}date", now)
            ins_run = etree.SubElement(ins_el, f"{W}r")
            if rpr_bytes:
                ins_run.append(etree.fromstring(rpr_bytes))
            ins_t = etree.SubElement(ins_run, f"{W}t")
            _preserve(ins_t, new_text)
            parent.insert(insert_at, ins_el)
            insert_at += 1

            # Text after
            end = idx + len(old_text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark(target_part)
            return {"location": location, "changes": changes}

        raise ValueError(f"Text '{old_text}' not found in {location}")
