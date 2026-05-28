"""Prepend the user-supplied input/titlist.docx onto TMP/thesis.docx as the
first page, using Word's "Different First Page" (titlePg) mechanism:

  first-page header  = university header copied from titlist
  default header     = one empty paragraph (GOST normcontrol requirement)
  first-page footer  = absent → title page has no page number
  default footer     = kept from the working docx (PAGE field via reference.docx)
"""
from __future__ import annotations

import copy

from docx import Document
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..pipeline import Context

_HEADER_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
_HDR_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
_EMPTY_HEADER_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:p/>'
    b'</w:hdr>'
)


def _copy_first_header_part(title_part, main_part) -> str | None:
    existing = {str(p.partname) for p in main_part.package.iter_parts()}
    for rel in title_part.rels.values():
        if rel.is_external or rel.reltype != _HEADER_RELTYPE:
            continue
        src = rel.target_part
        partname = str(src.partname)
        if partname in existing:
            base, ext = partname.rsplit(".", 1)
            n = 1
            while f"{base}_{n}.{ext}" in existing:
                n += 1
            partname = f"{base}_{n}.{ext}"
        new_part = Part(PackURI(partname), src.content_type, src.blob, main_part.package)
        return main_part.relate_to(new_part, _HEADER_RELTYPE)
    return None


def _add_empty_header_part(main_part) -> str:
    existing = {str(p.partname) for p in main_part.package.iter_parts()}
    partname = "/word/header_empty.xml"
    if partname in existing:
        partname = "/word/header_empty2.xml"
    part = Part(PackURI(partname), _HDR_CONTENT_TYPE, _EMPTY_HEADER_XML, main_part.package)
    return main_part.relate_to(part, _HEADER_RELTYPE)


def run(ctx: Context) -> None:
    if not ctx.paths.titlepage_docx.exists():
        raise FileNotFoundError(
            f"Title page docx not found: {ctx.paths.titlepage_docx}. "
            f"Place a prepared {ctx.paths.titlepage_docx.name} in input/."
        )

    title_doc = Document(ctx.paths.titlepage_docx)
    main_doc = Document(ctx.paths.working_docx)
    title_part = title_doc.part
    main_part = main_doc.part

    first_hdr_rId = _copy_first_header_part(title_part, main_part)
    empty_hdr_rId = _add_empty_header_part(main_part)

    main_body = main_doc.element.body
    main_sectPr = main_body.find(qn("w:sectPr"))
    if main_sectPr is not None:
        for tag in (qn("w:headerReference"), qn("w:titlePg")):
            for elem in main_sectPr.findall(tag):
                main_sectPr.remove(elem)
        if first_hdr_rId:
            ref = OxmlElement("w:headerReference")
            ref.set(qn("w:type"), "first")
            ref.set(qn("r:id"), first_hdr_rId)
            main_sectPr.insert(0, ref)
        for htype in ("default", "even"):
            ref = OxmlElement("w:headerReference")
            ref.set(qn("w:type"), htype)
            ref.set(qn("r:id"), empty_hdr_rId)
            main_sectPr.insert(0, ref)
        main_sectPr.append(OxmlElement("w:titlePg"))

    title_body = title_doc.element.body
    title_elems = [copy.deepcopy(e) for e in title_body if e.tag != qn("w:sectPr")]

    last_para = next((e for e in reversed(title_elems) if e.tag == qn("w:p")), None)
    if last_para is None:
        last_para = OxmlElement("w:p")
        title_elems.append(last_para)
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    last_para.append(r)

    for elem in reversed(title_elems):
        main_body.insert(0, elem)

    main_doc.save(ctx.paths.working_docx)
    print(f"  title page from {ctx.paths.titlepage_docx.name} prepended")
