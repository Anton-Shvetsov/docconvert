"""Align variable-definition paragraphs after «where» blocks (Russian: «где»).

Layout goal (tab-stop based, no manual spaces):

  где[TAB]$x_1$ -- definition;   ← TAB snaps $x_1$ to 1.25 cm
          $x_2$ -- definition.   ← continuation lines start at 1.25 cm

The trigger word is configurable via ``config.strings.where_keyword``.
"""
from __future__ import annotations

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

from ..pipeline import Context

_DEFN_INDENT = Cm(1.25)
_TAB_TWIPS = round(_DEFN_INDENT / 635)   # 1.25 cm in twips ≈ 709
_MAX_DEFN_LEN = 2000


def _has_content(para) -> bool:
    return bool(para.text.strip())


def _starts_with(para, keyword: str) -> bool:
    for run in para.runs:
        t = run.text.strip()
        if t:
            return t.startswith(keyword)
    return False


def _is_definition(para) -> bool:
    text = para.text
    return len(text) <= _MAX_DEFN_LEN and (" -- " in text or " – " in text)


def _zero_spacing(para) -> None:
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")


def _set_tab_stop(para) -> None:
    pPr = para._p.get_or_add_pPr()
    tabs_el = pPr.find(qn("w:tabs"))
    if tabs_el is None:
        tabs_el = OxmlElement("w:tabs")
        pPr.append(tabs_el)
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "left")
    tab.set(qn("w:pos"), str(_TAB_TWIPS))
    tabs_el.append(tab)


def _insert_tab_after_keyword(para, keyword: str) -> None:
    p_el = para._p
    children = list(p_el)
    keyword_idx = None
    for idx, child in enumerate(children):
        if child.tag != qn("w:r"):
            continue
        t_el = child.find(qn("w:t"))
        if t_el is not None and t_el.text and t_el.text.strip().startswith(keyword):
            t_el.text = t_el.text.rstrip()
            tab_el = OxmlElement("w:tab")
            child.append(tab_el)
            keyword_idx = idx
            break
    if keyword_idx is None:
        return
    for child in children[keyword_idx + 1:]:
        if child.tag == qn("w:r"):
            t_el = child.find(qn("w:t"))
            if t_el is not None and t_el.text:
                t_el.text = t_el.text.lstrip()
                if not t_el.text:
                    p_el.remove(child)
        break


def run(ctx: Context) -> None:
    keyword = ctx.config.strings.where_keyword
    doc = Document(ctx.paths.working_docx)
    paras = doc.paragraphs

    i = 0
    fixed = 0
    while i < len(paras):
        if _starts_with(paras[i], keyword):
            head = paras[i]
            head.paragraph_format.left_indent = Cm(0)
            head.paragraph_format.first_line_indent = Cm(0)
            _set_tab_stop(head)
            _insert_tab_after_keyword(head, keyword)
            i += 1
            defns = []
            while i < len(paras):
                p = paras[i]
                if not _has_content(p):
                    i += 1
                    continue
                if _is_definition(p):
                    p.paragraph_format.left_indent = Cm(0)
                    p.paragraph_format.first_line_indent = _DEFN_INDENT
                    defns.append(p)
                    fixed += 1
                    i += 1
                else:
                    break
            for p in [head, *defns]:
                _zero_spacing(p)
        else:
            i += 1

    doc.save(ctx.paths.working_docx)
    print(f"  shifted {fixed} definition paragraph(s) after «{keyword}»")
