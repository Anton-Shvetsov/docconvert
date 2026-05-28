"""Number display equations in TMP/thesis.docx using Word's m:eqArr scheme.

For each m:oMathPara found, wraps its last m:oMath child in:

  m:eqArr
    m:eqArrPr (m:maxDist val="1")
    m:e
      [original equation children]
      m:r " #"
      m:d (N)              ← parenthesised number, pushed to the right margin

Also inserts a blank paragraph before/after the equation paragraph if the
neighbour is non-empty body text (GOST normcontrol requirement).
"""
from __future__ import annotations

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..pipeline import Context

_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_OMATH_PARA = f"{{{_M}}}oMathPara"
_OMATH = f"{{{_M}}}oMath"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _m(tag: str):
    return OxmlElement(f"m:{tag}")


def _w(tag: str):
    return OxmlElement(f"w:{tag}")


def _mval(el, v) -> None:
    el.set(f"{{{_M}}}val", str(v))


def _cambria_rpr():
    rPr = _w("rPr")
    fonts = _w("rFonts")
    fonts.set(qn("w:ascii"), "Cambria Math")
    fonts.set(qn("w:hAnsi"), "Cambria Math")
    rPr.append(fonts)
    return rPr


def _cambria_rpr_italic():
    rPr = _cambria_rpr()
    rPr.append(_w("i"))
    return rPr


def _ctrlPr_italic():
    ctrl = _m("ctrlPr")
    ctrl.append(_cambria_rpr_italic())
    return ctrl


def _make_eqArr(eq_children, number: int):
    eqArrPr = _m("eqArrPr")
    maxDist = _m("maxDist")
    _mval(maxDist, 1)
    eqArrPr.append(maxDist)
    eqArrPr.append(_ctrlPr_italic())

    e = _m("e")
    for child in eq_children:
        e.append(child)

    r_hash = _m("r")
    r_hash.append(_cambria_rpr())
    t = _m("t")
    t.set(_XML_SPACE, "preserve")
    t.text = " #"
    r_hash.append(t)
    e.append(r_hash)

    d = _m("d")
    dPr = _m("dPr")
    dPr.append(_ctrlPr_italic())
    d.append(dPr)
    e_inner = _m("e")
    r_num = _m("r")
    r_num.append(_cambria_rpr())
    t_num = _m("t")
    t_num.text = str(number)
    r_num.append(t_num)
    e_inner.append(r_num)
    d.append(e_inner)
    e.append(d)

    eqArr = _m("eqArr")
    eqArr.append(eqArrPr)
    eqArr.append(e)
    return eqArr


def _make_empty_para():
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "Normal")
    pPr.append(pStyle)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    pPr.append(sp)
    p.append(pPr)
    return p


def _p_is_empty(p_elem) -> bool:
    return not "".join(t.text or "" for t in p_elem.iter(qn("w:t"))).strip()


def _p_is_heading(p_elem) -> bool:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return False
    ps = pPr.find(qn("w:pStyle"))
    return ps is not None and ps.get(qn("w:val"), "").lower().startswith("heading")


def run(ctx: Context) -> None:
    doc = Document(ctx.paths.working_docx)
    counter = 0

    body_children = list(doc.element.body)
    body_p_ids = {id(ch) for ch in body_children if ch.tag == qn("w:p")}

    for para in doc.paragraphs:
        p_elem = para._p
        omp = p_elem.find(_OMATH_PARA)
        if omp is None:
            continue
        omaths = omp.findall(_OMATH)
        if not omaths:
            continue

        counter += 1
        omath = omaths[-1]
        children = list(omath)
        for ch in children:
            omath.remove(ch)
        omath.append(_make_eqArr(children, counter))

        if id(p_elem) in body_p_ids:
            prev = p_elem.getprevious()
            if (
                prev is not None
                and prev.tag == qn("w:p")
                and not _p_is_empty(prev)
                and not _p_is_heading(prev)
            ):
                p_elem.addprevious(_make_empty_para())
            nxt = p_elem.getnext()
            if nxt is not None and nxt.tag == qn("w:p") and not _p_is_empty(nxt):
                p_elem.addnext(_make_empty_para())

    doc.save(ctx.paths.working_docx)
    print(f"  numbered {counter} display equation(s)")
