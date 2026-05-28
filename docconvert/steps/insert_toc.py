"""Insert an updatable Word TOC field into TMP/thesis.docx.

Also:
- Forces full-width / auto-layout / borders on every non-empty table (pandoc
  sometimes omits these).
- Detects table caption paragraphs, moves below-table captions above, and
  prepends 'Таблица N — ' with a SEQ field if absent.
- Normalises TOC1/TOC2/TOC3 styles and per-paragraph formatting for GOST.

The TOC is inserted immediately after the abstract section heading (its title
listed in ``config.strings.abstract_section_aliases``). Plain-text appendix
notes — one per ``config.appendices`` entry — are written after the TOC field
so they can be copy-pasted into the auto-generated TOC by hand (Word renders
only the heading text inside the TOC field itself).
"""
from __future__ import annotations

import re
from typing import Iterable

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..pipeline import Context

_CM_TO_TWIPS = 360000 / 635
_TEXT_WIDTH_TWIPS = round(16.5 * _CM_TO_TWIPS)   # ≈ 9354 twips, A4 with GOST margins
_TOC_ALIGN_TWIPS = round(1.25 * _CM_TO_TWIPS)    # ≈ 709 twips

_TOC_NUM_RE = re.compile(r"^(\d[\d.]*)\s")
_TOC_STYLE_IDS = {"TOC1", "TOC2", "TOC3", "toc1", "toc2", "toc3"}


def _set_toc_para_format(pPr) -> None:
    """Apply TOC alignment / indent / tabs / spacing to a w:pPr element."""
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "left")

    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:firstLine"), "0")
    ind.attrib.pop(qn("w:hanging"), None)

    tabs = pPr.find(qn("w:tabs"))
    if tabs is not None:
        pPr.remove(tabs)
    tabs = OxmlElement("w:tabs")
    for val, pos, leader in [
        ("clear", _TOC_ALIGN_TWIPS, None),
        ("left", _TOC_ALIGN_TWIPS, None),
        ("right", _TEXT_WIDTH_TWIPS, "dot"),
    ]:
        t = OxmlElement("w:tab")
        t.set(qn("w:val"), val)
        t.set(qn("w:pos"), str(pos))
        if leader:
            t.set(qn("w:leader"), leader)
        tabs.append(t)
    pPr.append(tabs)

    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "360")
    sp.set(qn("w:lineRule"), "auto")


def _fix_toc_styles(doc) -> None:
    """Define / patch TOC1, TOC2, TOC3 styles in the style table."""
    styles_root = doc.styles.element
    for style_id, style_name in [("TOC1", "toc 1"), ("TOC2", "toc 2"), ("TOC3", "toc 3")]:
        s = next(
            (x for x in styles_root.findall(qn("w:style")) if x.get(qn("w:styleId")) == style_id),
            None,
        )
        if s is None:
            s = OxmlElement("w:style")
            s.set(qn("w:type"), "paragraph")
            s.set(qn("w:styleId"), style_id)
            name_el = OxmlElement("w:name")
            name_el.set(qn("w:val"), style_name)
            s.append(name_el)
            styles_root.append(s)

        pPr = s.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            s.append(pPr)
        _set_toc_para_format(pPr)

        rPr = s.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            s.append(rPr)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr in (qn("w:ascii"), qn("w:hAnsi"), qn("w:cs")):
            rFonts.set(attr, "Times New Roman")
        for sz_name in ("w:sz", "w:szCs"):
            el = rPr.find(qn(sz_name))
            if el is None:
                el = OxmlElement(sz_name)
                rPr.append(el)
            el.set(qn("w:val"), "28")  # 14pt


def _insert_tab_after_num(p_elem) -> None:
    containers = [p_elem] + list(p_elem.findall(qn("w:hyperlink")))
    for container in containers:
        for run in container.findall(qn("w:r")):
            t = run.find(qn("w:t"))
            if t is None or not (t.text or "").strip():
                continue
            m = _TOC_NUM_RE.match(t.text)
            if not m:
                return
            num = m.group(1)
            rest = t.text[m.end():]
            t.text = num
            tab_el = OxmlElement("w:tab")
            t.addnext(tab_el)
            if rest:
                rest_t = OxmlElement("w:t")
                rest_t.set(qn("xml:space"), "preserve")
                rest_t.text = rest
                tab_el.addnext(rest_t)
            return


def _fix_toc_paragraphs(doc) -> None:
    fixed = 0
    for para in doc.paragraphs:
        p = para._p
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None or pStyle.get(qn("w:val"), "") not in _TOC_STYLE_IDS:
            continue
        _set_toc_para_format(pPr)
        _insert_tab_after_num(p)
        fixed += 1
    print(f"  TOC paragraphs reformatted: {fixed}")


def _fix_table(table) -> None:
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    existing_w = tblPr.find(qn("w:tblW"))
    if existing_w is not None:
        tblPr.remove(existing_w)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.insert(0, tblW)

    existing_layout = tblPr.find(qn("w:tblLayout"))
    if existing_layout is not None:
        tblPr.remove(existing_layout)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "auto")
    tblPr.append(layout)

    if tblPr.find(qn("w:tblBorders")) is None:
        borders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)


def _make_toc_field_paragraph():
    para = OxmlElement("w:p")
    specs = [("begin", None), (None, r' TOC \o "1-3" \h \z \u '), ("separate", None), ("end", None)]
    for fc_type, text in specs:
        r = OxmlElement("w:r")
        if fc_type is not None:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), fc_type)
            if fc_type == "begin":
                el.set(qn("w:dirty"), "true")
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        r.append(el)
        para.append(r)
    return para


def _make_appendix_note_paragraph(text: str):
    para = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)
    para.append(pPr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    para.append(r)
    return para


def _make_pagebreak_paragraph():
    para = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    para.append(r)
    return para


def _make_toc_manual_checklist(appendices):
    """Build the checklist of manual TOC fixups (Word can't do them automatically).
    Returns a list of plain paragraphs to insert right after the TOC field."""
    paras = []
    if appendices:
        paras.append(_make_appendix_note_paragraph("Вставьте названия Приложений в Содержание:"))
        for a in appendices:
            paras.append(_make_appendix_note_paragraph(f"{a.code}. {a.title}"))
        paras.append(_make_appendix_note_paragraph(""))
    paras.append(_make_appendix_note_paragraph(
        "Удалите пункт Содержание из Содержания и удалите данную пометку"
    ))
    return paras


def _find_insert_index_after_abstract(body, aliases: Iterable[str]) -> int | None:
    aliases = list(aliases)
    found_abstract = False
    for i, child in enumerate(list(body)):
        if child.tag != qn("w:p"):
            continue
        pPr = child.find(qn("w:pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("w:pStyle"))
        if ps is None:
            continue
        if not ps.get(qn("w:val"), "").lower().startswith("heading"):
            continue
        text = "".join(t.text or "" for t in child.iter(qn("w:t")))
        if any(alias in text for alias in aliases):
            found_abstract = True
        elif found_abstract:
            return i
    return None


def _is_caption_style_p(p_elem) -> bool:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return False
    pStyle = pPr.find(qn("w:pStyle"))
    return pStyle is not None and "caption" in pStyle.get(qn("w:val"), "").lower()


def _p_text_content(p_elem) -> str:
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))


def _apply_tbl_cap_style(cap_p) -> None:
    cap_pPr = cap_p.find(qn("w:pPr"))
    if cap_pPr is None:
        cap_pPr = OxmlElement("w:pPr")
        cap_p.insert(0, cap_pPr)
    jc = cap_pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        cap_pPr.append(jc)
    jc.set(qn("w:val"), "left")
    ind = cap_pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        cap_pPr.append(ind)
    ind.attrib.pop(qn("w:firstLineChars"), None)
    ind.set(qn("w:firstLine"), "0")


def _prepend_caption_label(para_elem, label_prefix: str, seq_name: str, placeholder_num: int) -> None:
    def _mk_run(text):
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        return r

    def _mk_fldchar(ftype):
        r = OxmlElement("w:r")
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), ftype)
        r.append(fc)
        return r

    def _mk_instrtext(text):
        r = OxmlElement("w:r")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = text
        r.append(it)
        return r

    first_r = next((c for c in para_elem if c.tag != qn("w:pPr")), None)
    new_elems = [
        _mk_run(label_prefix + " "),
        _mk_fldchar("begin"),
        _mk_instrtext(f" SEQ {seq_name} \\* ARABIC "),
        _mk_fldchar("separate"),
        _mk_run(str(placeholder_num)),
        _mk_fldchar("end"),
        _mk_run(" — "),
    ]
    if first_r is not None:
        for el in new_elems:
            first_r.addprevious(el)
    else:
        for el in new_elems:
            para_elem.append(el)


def _fix_table_captions(doc, table_label: str) -> None:
    body = doc.element.body
    snapshot = list(body)
    already_re = re.compile(rf"^{re.escape(table_label)}\s+\d")
    counter = 0
    for ti, el in enumerate(snapshot):
        if el.tag != qn("w:tbl"):
            continue
        cap_p = None
        if ti > 0 and snapshot[ti - 1].tag == qn("w:p") and _is_caption_style_p(snapshot[ti - 1]):
            cap_p = snapshot[ti - 1]
        if cap_p is None and ti + 1 < len(snapshot):
            nxt = snapshot[ti + 1]
            if nxt.tag == qn("w:p") and _is_caption_style_p(nxt):
                cap_p = nxt
                body.remove(cap_p)
                el.addprevious(cap_p)
        if cap_p is None:
            continue
        if already_re.match(_p_text_content(cap_p).strip()):
            _apply_tbl_cap_style(cap_p)
            continue
        counter += 1
        _prepend_caption_label(cap_p, table_label, table_label, counter)
        _apply_tbl_cap_style(cap_p)


def run(ctx: Context) -> None:
    doc = Document(ctx.paths.working_docx)
    body = doc.element.body

    heading_para = doc.add_paragraph("СОДЕРЖАНИЕ", style="Heading 1")
    heading_elem = heading_para._element
    body.remove(heading_elem)

    toc_elem = _make_toc_field_paragraph()
    checklist = _make_toc_manual_checklist(ctx.config.appendices)
    pb_elem = _make_pagebreak_paragraph()

    insert_idx = _find_insert_index_after_abstract(body, ctx.config.strings.abstract_section_aliases)
    insert_order = [heading_elem, toc_elem, *checklist, pb_elem]
    target_idx = insert_idx if insert_idx is not None else 0
    for el in reversed(insert_order):
        body.insert(target_idx, el)

    for t in doc.tables:
        if any(cell.text.strip() for row in t.rows for cell in row.cells):
            _fix_table(t)

    _fix_table_captions(doc, ctx.config.strings.table_label)
    _fix_toc_styles(doc)
    _fix_toc_paragraphs(doc)

    doc.save(ctx.paths.working_docx)
    print(f"  TOC inserted into {ctx.paths.working_docx.name}")
