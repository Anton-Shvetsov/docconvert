"""Force consistent text styling on TMP/thesis.docx:
- black color on every heading style and every inline run (pandoc / theme fills);
- outline level 9 on the abstract section heading so the auto TOC excludes it;
- Heading 1 style + hanging-indent + 14pt on numbered chapter headings;
- zero space_after on BodyText paragraphs lacking an explicit override.
"""
from __future__ import annotations

import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..docx_utils import force_black_on_rpr, para_text, set_outline_level, strip_run_colors
from ..pipeline import Context

_NUMBERED_HEADING_RE = re.compile(r"^\d+(\.\d+)*")


def _strip_all_run_colors(para_elem) -> None:
    for r in para_elem.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            continue
        color = rPr.find(qn("w:color"))
        if color is not None:
            rPr.remove(color)


def _apply_numbered_heading_format(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "both")
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLine"), "709")  # 1.25 cm
    ind.attrib.pop(qn("w:left"), None)
    ind.attrib.pop(qn("w:hanging"), None)
    ind.attrib.pop(qn("w:firstLineChars"), None)
    # Force 14pt on every run (Heading 1 default is 16pt for structural sections)
    for r in p_elem.iter(qn("w:r")):
        r_rPr = r.find(qn("w:rPr"))
        if r_rPr is None:
            r_rPr = OxmlElement("w:rPr")
            r.insert(0, r_rPr)
        for sz_tag in ("w:sz", "w:szCs"):
            sz_el = r_rPr.find(qn(sz_tag))
            if sz_el is None:
                sz_el = OxmlElement(sz_tag)
                r_rPr.append(sz_el)
            sz_el.set(qn("w:val"), "28")  # 14pt = 28 half-points


def run(ctx: Context) -> None:
    doc = Document(ctx.paths.working_docx)
    strings = ctx.config.strings

    # Pass 1: black on every Heading style definition.
    for style in doc.styles:
        name = getattr(style, "name", "") or ""
        if not name.startswith("Heading"):
            continue
        el = style.element
        if el is None:
            continue
        force_black_on_rpr(el)
        pPr = el.find(qn("w:pPr"))
        if pPr is not None:
            force_black_on_rpr(pPr)

    # Pass 2: strip inline colors, normalise structural headings.
    abstract_aliases = list(strings.abstract_section_aliases)
    references_title = strings.references_section
    for para in doc.paragraphs:
        _strip_all_run_colors(para._p)
        text = para_text(para._p)
        stripped = text.strip()

        if any(stripped.startswith(alias) for alias in abstract_aliases):
            set_outline_level(para._p, 9)

        if references_title in text:
            para.style = doc.styles["Heading 1"]
            _strip_all_run_colors(para._p)

        if para.style and para.style.name == "Heading 1":
            if _NUMBERED_HEADING_RE.match(stripped):
                _strip_all_run_colors(para._p)
                _apply_numbered_heading_format(para._p)

    # Pass 8: zero space_after on BodyText paragraphs lacking an explicit value.
    for para in doc.paragraphs:
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("w:pStyle"))
        if ps is None or ps.get(qn("w:val"), "") != "BodyText":
            continue
        sp = pPr.find(qn("w:spacing"))
        if sp is not None and sp.get(qn("w:after")) is not None:
            continue
        if sp is None:
            sp = OxmlElement("w:spacing")
            pPr.append(sp)
        sp.set(qn("w:after"), "0")

    doc.save(ctx.paths.working_docx)
    print("  styles enforced: headings, colors, BodyText spacing")
