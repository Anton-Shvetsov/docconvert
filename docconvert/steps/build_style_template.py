"""Produce reference.docx — the style template pandoc applies via --reference-doc.

All style parameters (font, size, spacing, margins, indent) come from
``config.style``. The output is written to ``TMP/reference.docx``.
"""
from __future__ import annotations

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from ..docx_utils import force_color_black, force_font
from ..pipeline import Context


def _apply_heading_style(style, size_pt, *, centered, page_break_before, indent_cm, head_sp):
    force_font(style, "Times New Roman")
    force_color_black(style)
    style.font.size = Pt(size_pt)
    style.font.bold = True
    style.font.all_caps = False
    pf = style.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(indent_cm)
    pf.left_indent = Cm(0)
    pf.page_break_before = page_break_before
    pf.space_before = head_sp
    pf.space_after = head_sp


def _setup_table_grid(doc, font_pt):
    tg_name = "Table Grid"
    style_names = [s.name for s in doc.styles]
    tg = doc.styles[tg_name] if tg_name in style_names else doc.styles.add_style(tg_name, WD_STYLE_TYPE.TABLE)
    force_font(tg, "Times New Roman")
    tg.font.size = Pt(font_pt)

    tblPr = tg.element.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tg.element.append(tblPr)

    for child_tag in ("w:tblW", "w:tblBorders", "w:tblLayout", "w:tblCellMar"):
        existing = tblPr.find(qn(child_tag))
        if existing is not None:
            tblPr.remove(existing)

    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.insert(0, tblW)

    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    tblPr.append(tblBorders)

    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "auto")
    tblPr.append(tblLayout)

    tblCellMar = OxmlElement("w:tblCellMar")
    for side, val in (("top", "55"), ("left", "108"), ("bottom", "55"), ("right", "108")):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val)
        m.set(qn("w:type"), "dxa")
        tblCellMar.append(m)
    tblPr.append(tblCellMar)


def _setup_source_code(doc):
    sc_name = "Source Code"
    style_names = [s.name for s in doc.styles]
    sc = doc.styles[sc_name] if sc_name in style_names else doc.styles.add_style(sc_name, WD_STYLE_TYPE.PARAGRAPH)
    force_font(sc, "Courier New")
    sc.font.size = Pt(10)
    sc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sc.paragraph_format.first_line_indent = Cm(0)
    sc.paragraph_format.space_before = Pt(0)
    sc.paragraph_format.space_after = Pt(0)
    sc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    rPr = sc.element.get_or_add_rPr()
    for tag, attr, val in [("w:spacing", "w:val", "0"), ("w:kern", "w:val", "0")]:
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn(attr), val)


def _setup_verbatim_char(doc, font_pt):
    vc_name = "VerbatimChar"
    style_names = [s.name for s in doc.styles]
    vc = doc.styles[vc_name] if vc_name in style_names else doc.styles.add_style(vc_name, WD_STYLE_TYPE.CHARACTER)
    force_font(vc, "Courier New")
    vc.font.size = Pt(font_pt)


def _setup_footer(section):
    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    for tag, ftype in [("w:fldChar", "begin"), ("w:instrText", None), ("w:fldChar", "end")]:
        el = OxmlElement(tag)
        if ftype:
            el.set(qn("w:fldCharType"), ftype)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = " PAGE "
        run._r.append(el)


def run(ctx: Context) -> None:
    style_cfg = ctx.config.style
    margins = style_cfg.margins_mm
    indent_cm = style_cfg.paragraph_indent_cm

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(margins["left"] / 10)
    section.right_margin = Cm(margins["right"] / 10)
    section.top_margin = Cm(margins["top"] / 10)
    section.bottom_margin = Cm(margins["bottom"] / 10)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    section.different_first_page_header_footer = True

    # Extra spacing around headings: (2 - line_spacing) * font_size_pt
    head_sp = Pt((2 - style_cfg.line_spacing) * style_cfg.font_size_pt)

    normal = doc.styles["Normal"]
    normal.font.name = style_cfg.font
    normal.font.size = Pt(style_cfg.font_size_pt)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    normal.paragraph_format.first_line_indent = Cm(indent_cm)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    _apply_heading_style(
        doc.styles["Heading 1"],
        size_pt=16, centered=True, page_break_before=True, indent_cm=0, head_sp=head_sp,
    )
    for name in ("Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6"):
        style_names = [s.name for s in doc.styles]
        h = doc.styles[name] if name in style_names else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        _apply_heading_style(
            h,
            size_pt=style_cfg.font_size_pt,
            centered=False,
            page_break_before=False,
            indent_cm=indent_cm,
            head_sp=head_sp,
        )

    _setup_footer(section)
    _setup_table_grid(doc, font_pt=style_cfg.font_size_pt)
    _setup_source_code(doc)
    _setup_verbatim_char(doc, font_pt=style_cfg.font_size_pt)

    ctx.paths.reference_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(ctx.paths.reference_docx)
    print(f"  reference.docx → {ctx.paths.reference_docx}")
