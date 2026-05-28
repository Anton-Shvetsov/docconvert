"""Fill the figures/tables/sources counts in the abstract section line:

    Расчетно-пояснительная записка: ___ с., N рис., M табл., K источников.

All trigger strings are taken from ``config.strings``. Page count is left as
``___`` for manual entry.
"""
from __future__ import annotations

import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..pipeline import Context

_NBSP = " "


def run(ctx: Context) -> None:
    strings = ctx.config.strings
    figure_label = strings.figure_label
    table_label = strings.table_label
    bib_title = strings.references_section
    line_prefix = strings.rpz_line_prefix

    fig_re = re.compile(rf"^{re.escape(figure_label)}\s+\S")
    tbl_re = re.compile(rf"^{re.escape(table_label)}\s+\S")
    bib_entry_re = re.compile(r"^\d+\.")

    doc = Document(ctx.paths.working_docx)
    figures = tables = sources = 0
    in_bib = False
    target_para = None

    for para in doc.paragraphs:
        text = para.text.strip()
        style = para.style.name if para.style else ""

        if bib_title in text:
            in_bib = True
            continue

        if in_bib:
            if "Heading" in style or (text and text.isupper() and len(text) > 3):
                in_bib = False
            elif bib_entry_re.match(text):
                sources += 1

        if fig_re.match(text):
            figures += 1
        if tbl_re.match(text):
            tables += 1
        if line_prefix in text:
            target_para = para

    new_text = (
        f"{line_prefix} ___{_NBSP}с., "
        f"{figures}{_NBSP}рис., {tables}{_NBSP}табл., {sources}{_NBSP}источников."
    )

    if target_para is None:
        print(f"  WARNING: '{line_prefix}' line not found — skipping")
    else:
        p_elem = target_para._p
        for r in list(p_elem.findall(qn("w:r"))):
            p_elem.remove(r)
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = new_text
        r.append(t)
        p_elem.append(r)
        print(f"  filled: ___ с., {figures} рис., {tables} табл., {sources} источников.")

    doc.save(ctx.paths.working_docx)
