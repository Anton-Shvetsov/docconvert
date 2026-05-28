"""Shared helpers for python-docx XML manipulation."""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def force_font(style, name: str) -> None:
    """Set font on a style, clearing theme overrides that override font.name."""
    style.font.name = name
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.attrib.pop(qn("w:asciiTheme"), None)
    rFonts.attrib.pop(qn("w:hAnsiTheme"), None)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)


def force_color_black(style) -> None:
    """Force a style's color to black, stripping any theme color override."""
    rPr = style.element.get_or_add_rPr()
    color_el = rPr.find(qn("w:color"))
    if color_el is None:
        color_el = OxmlElement("w:color")
        rPr.append(color_el)
    color_el.set(qn("w:val"), "000000")
    color_el.attrib.pop(qn("w:themeColor"), None)
    color_el.attrib.pop(qn("w:themeTint"), None)
    color_el.attrib.pop(qn("w:themeShade"), None)


def force_black_on_rpr(rpr_parent) -> None:
    """Ensure the rPr child of rpr_parent has w:color = 000000 (no theme)."""
    if rpr_parent is None:
        return
    rPr = rpr_parent.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rpr_parent.insert(0, rPr)
    color_el = rPr.find(qn("w:color"))
    if color_el is None:
        color_el = OxmlElement("w:color")
        rPr.append(color_el)
    color_el.set(qn("w:val"), "000000")
    color_el.attrib.pop(qn("w:themeColor"), None)
    color_el.attrib.pop(qn("w:themeTint"), None)
    color_el.attrib.pop(qn("w:themeShade"), None)


def strip_run_colors(container) -> None:
    """Remove inline <w:color> from every run inside container."""
    for r in container.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            continue
        color_el = rPr.find(qn("w:color"))
        if color_el is not None:
            rPr.remove(color_el)


def set_outline_level(para_elem, level: int) -> None:
    pPr = para_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_elem.insert(0, pPr)
    outline = pPr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        pPr.append(outline)
    outline.set(qn("w:val"), str(level))


def para_text(para_elem) -> str:
    parts = [t.text or "" for t in para_elem.iter(qn("w:t"))]
    return "".join(parts)


def p_style_id(p_elem) -> str | None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return None
    ps = pPr.find(qn("w:pStyle"))
    if ps is None:
        return None
    return ps.get(qn("w:val"))


def set_p_style(p_elem, style_id: str) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    ps = pPr.find(qn("w:pStyle"))
    if ps is None:
        ps = OxmlElement("w:pStyle")
        pPr.insert(0, ps)
    ps.set(qn("w:val"), style_id)
