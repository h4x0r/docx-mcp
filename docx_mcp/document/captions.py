"""Caption mixin: insert figure/table captions into a Word document."""

from __future__ import annotations

from lxml import etree

from .base import W14, XML_SPACE, W


class CaptionMixin:
    def insert_caption(self, after_para_id: str, text: str, label: str = "Figure") -> dict:
        doc = self._tree("word/document.xml")

        target = self._find_para(doc, after_para_id)
        if target is None:
            raise ValueError(f"Paragraph '{after_para_id}' not found")

        existing_captions = [
            p
            for p in doc.iter(f"{W}p")
            if (
                p.find(f"{W}pPr/{W}pStyle") is not None
                and p.find(f"{W}pPr/{W}pStyle").get(f"{W}val", "").lower() == "caption"
            )
        ]
        seq_num = len(existing_captions) + 1

        new_pid = self._new_para_id()

        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", new_pid)
        new_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(new_para, f"{W}pPr")
        pstyle = etree.SubElement(ppr, f"{W}pStyle")
        pstyle.set(f"{W}val", "Caption")

        run = etree.SubElement(new_para, f"{W}r")
        t = etree.SubElement(run, f"{W}t")
        caption_text = f"{label} {seq_num}: {text}"
        t.text = caption_text
        t.set(XML_SPACE, "preserve")

        parent = target.getparent()
        idx = list(parent).index(target)
        parent.insert(idx + 1, new_para)

        self._mark("word/document.xml")

        return {
            "para_id": new_pid,
            "label": label,
            "seq_num": seq_num,
            "text": text,
        }
