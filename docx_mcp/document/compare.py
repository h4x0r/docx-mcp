"""Document comparison mixin: diff two DOCX files into tracked-change output."""

from __future__ import annotations

import copy
import difflib
import re
import zipfile
from pathlib import Path

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_COMPARE_AUTHOR = "Compare"
_COMPARE_DATE = "2026-01-01T00:00:00Z"


def _wt(tag: str) -> str:
    return f"{{{_W}}}{tag}"


def _parse_doc_root(docx_path: str) -> etree._Element:
    with zipfile.ZipFile(docx_path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def _body_paragraphs(root: etree._Element) -> list[etree._Element]:
    body = root.find(_wt("body"))
    if body is None:
        return []
    return [el for el in body if el.tag == _wt("p")]


def _para_text(para: etree._Element) -> str:
    parts: list[str] = []
    for t in para.iter(_wt("t")):
        parts.append(t.text or "")
    return "".join(parts)


def _del_para(para: etree._Element, cid: int) -> etree._Element:
    """Return a new w:p whose text runs are wrapped in a single w:del."""
    new_para = etree.Element(_wt("p"))
    pPr = para.find(_wt("pPr"))
    if pPr is not None:
        new_para.append(copy.deepcopy(pPr))
    text = _para_text(para)
    if text:
        del_el = etree.SubElement(new_para, _wt("del"))
        del_el.set(_wt("id"), str(cid))
        del_el.set(_wt("author"), _COMPARE_AUTHOR)
        del_el.set(_wt("date"), _COMPARE_DATE)
        run = etree.SubElement(del_el, _wt("r"))
        dt = etree.SubElement(run, _wt("delText"))
        dt.set(_XML_SPACE, "preserve")
        dt.text = text
    return new_para


def _ins_para(para: etree._Element, cid: int) -> etree._Element:
    """Return a new w:p whose text runs are wrapped in a single w:ins."""
    new_para = etree.Element(_wt("p"))
    pPr = para.find(_wt("pPr"))
    if pPr is not None:
        new_para.append(copy.deepcopy(pPr))
    text = _para_text(para)
    if text:
        ins_el = etree.SubElement(new_para, _wt("ins"))
        ins_el.set(_wt("id"), str(cid))
        ins_el.set(_wt("author"), _COMPARE_AUTHOR)
        ins_el.set(_wt("date"), _COMPARE_DATE)
        run = etree.SubElement(ins_el, _wt("r"))
        t = etree.SubElement(run, _wt("t"))
        t.set(_XML_SPACE, "preserve")
        t.text = text
    return new_para


def _word_diff_para(
    base_text: str,
    rev_text: str,
    start_id: int,
) -> tuple[list[etree._Element], int]:
    """
    Build XML children for a word-level tracked-change paragraph.

    Returns (children, next_available_id).
    Children are w:r, w:del, w:ins elements to be appended to a w:p.
    """
    tokens_base = re.split(r"(\s+)", base_text)
    tokens_rev = re.split(r"(\s+)", rev_text)

    matcher = difflib.SequenceMatcher(None, tokens_base, tokens_rev, autojunk=False)
    children: list[etree._Element] = []
    cid = start_id

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            text = "".join(tokens_base[i1:i2])
            if text:
                run = etree.Element(_wt("r"))
                t_el = etree.SubElement(run, _wt("t"))
                t_el.set(_XML_SPACE, "preserve")
                t_el.text = text
                children.append(run)

        elif tag in ("delete", "replace"):
            del_text = "".join(tokens_base[i1:i2])
            if del_text:
                del_el = etree.Element(_wt("del"))
                del_el.set(_wt("id"), str(cid))
                del_el.set(_wt("author"), _COMPARE_AUTHOR)
                del_el.set(_wt("date"), _COMPARE_DATE)
                cid += 1
                run = etree.SubElement(del_el, _wt("r"))
                dt = etree.SubElement(run, _wt("delText"))
                dt.set(_XML_SPACE, "preserve")
                dt.text = del_text
                children.append(del_el)
            if tag == "replace":
                ins_text = "".join(tokens_rev[j1:j2])
                if ins_text:
                    ins_el = etree.Element(_wt("ins"))
                    ins_el.set(_wt("id"), str(cid))
                    ins_el.set(_wt("author"), _COMPARE_AUTHOR)
                    ins_el.set(_wt("date"), _COMPARE_DATE)
                    cid += 1
                    run = etree.SubElement(ins_el, _wt("r"))
                    t_el = etree.SubElement(run, _wt("t"))
                    t_el.set(_XML_SPACE, "preserve")
                    t_el.text = ins_text
                    children.append(ins_el)

        elif tag == "insert":
            ins_text = "".join(tokens_rev[j1:j2])
            if ins_text:
                ins_el = etree.Element(_wt("ins"))
                ins_el.set(_wt("id"), str(cid))
                ins_el.set(_wt("author"), _COMPARE_AUTHOR)
                ins_el.set(_wt("date"), _COMPARE_DATE)
                cid += 1
                run = etree.SubElement(ins_el, _wt("r"))
                t_el = etree.SubElement(run, _wt("t"))
                t_el.set(_XML_SPACE, "preserve")
                t_el.text = ins_text
                children.append(ins_el)

    return children, cid


class CompareMixin:
    """Compare two DOCX files and produce tracked-change output."""

    @staticmethod
    def compare_documents(
        base_path: str,
        revised_path: str,
        output_path: str = "",
    ) -> dict:
        """Diff two DOCX files and write the result as a tracked-change document.

        Paragraph-level LCS diff:
          - Equal paragraphs → copied verbatim from base.
          - Deleted paragraphs (in base, not revised) → wrapped in w:del.
          - Inserted paragraphs (in revised, not base) → wrapped in w:ins.
          - Modified paragraphs (1:1 replace) → word-level del+ins inline.
          - N:M replace blocks → delete all base paras, insert all revised paras.

        output_path is auto-generated (base stem + "_compared.docx") if empty.
        Returns {"path": output_path}.
        """
        if not output_path:
            base_p = Path(base_path)
            output_path = str(base_p.parent / (base_p.stem + "_compared.docx"))

        base_root = _parse_doc_root(base_path)
        rev_root = _parse_doc_root(revised_path)

        base_paras = _body_paragraphs(base_root)
        rev_paras = _body_paragraphs(rev_root)

        base_texts = [_para_text(p) for p in base_paras]
        rev_texts = [_para_text(p) for p in rev_paras]

        matcher = difflib.SequenceMatcher(None, base_texts, rev_texts, autojunk=False)

        cid = 1
        output_paras: list[etree._Element] = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                for para in base_paras[i1:i2]:
                    output_paras.append(copy.deepcopy(para))

            elif tag == "delete":
                for para in base_paras[i1:i2]:
                    output_paras.append(_del_para(para, cid))
                    cid += 1

            elif tag == "insert":
                for para in rev_paras[j1:j2]:
                    output_paras.append(_ins_para(para, cid))
                    cid += 1

            elif tag == "replace":
                if (i2 - i1) == 1 and (j2 - j1) == 1:
                    # 1:1 replace → word-level inline diff
                    new_para = etree.Element(_wt("p"))
                    children, cid = _word_diff_para(base_texts[i1], rev_texts[j1], cid)
                    for child in children:
                        new_para.append(child)
                    output_paras.append(new_para)
                else:
                    # N:M replace → delete block then insert block
                    for para in base_paras[i1:i2]:
                        output_paras.append(_del_para(para, cid))
                        cid += 1
                    for para in rev_paras[j1:j2]:
                        output_paras.append(_ins_para(para, cid))
                        cid += 1

        # Build output document: clone base root, replace body paragraphs
        out_root = copy.deepcopy(base_root)
        body = out_root.find(_wt("body"))
        if body is None:
            body = etree.SubElement(out_root, _wt("body"))

        # Remove existing w:p children, keep sectPr and others
        for child in list(body):
            if child.tag == _wt("p"):
                body.remove(child)

        # Insert before sectPr if present
        body_children = list(body)
        sect_pr = next(
            (c for c in body_children if c.tag == _wt("sectPr")), None
        )
        insert_before = sect_pr if sect_pr is not None else None

        for para in output_paras:
            if insert_before is not None:
                insert_before.addprevious(para)
            else:
                body.append(para)

        out_bytes = etree.tostring(
            out_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # Copy base DOCX, replace word/document.xml with merged content
        with zipfile.ZipFile(base_path) as src_zip:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for name in src_zip.namelist():
                    out_zip.writestr(
                        name,
                        out_bytes if name == "word/document.xml" else src_zip.read(name),
                    )

        return {"path": output_path}
