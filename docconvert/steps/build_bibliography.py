"""Replace whatever citeproc put under the references heading with cleanly
formatted [1] … [N] paragraphs per ГОСТ Р 7.0.100-2018.

Citation order is recovered from the LaTeX source: ``input/main.tex`` is parsed
for ``\\input{...}`` directives and every ``\\cite{key}`` inside those files
contributes its keys in first-mention order. Keys not present in
``bibliography.bib`` are skipped.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..pipeline import Context


# --------------------------------------------------
# Citation-order extraction from LaTeX source
# --------------------------------------------------

def _citation_order(main_tex: Path) -> list[str]:
    if not main_tex.exists():
        return []
    input_files: list[Path] = []
    with main_tex.open(encoding="utf-8", errors="replace") as f:
        for line in f:
            m = re.match(r"\s*\\input\{([^}]+)\}", line)
            if m:
                fname = m.group(1).strip()
                if not fname.endswith(".tex"):
                    fname += ".tex"
                input_files.append(main_tex.parent / fname)
    cite_re = re.compile(r"\\cite[tp]?\*?\{([^}]+)\}")
    seen: list[str] = []
    for path in input_files:
        if not path.exists():
            continue
        with path.open(encoding="utf-8", errors="replace") as f:
            for line in f:
                if line.lstrip().startswith("%"):
                    continue
                line = re.sub(r"(?<!\\)%.*", "", line)
                for m in cite_re.finditer(line):
                    for key in m.group(1).split(","):
                        key = key.strip()
                        if key and key not in seen:
                            seen.append(key)
    return seen


# --------------------------------------------------
# .bib parser (minimal)
# --------------------------------------------------

_LATEX_REPLACEMENTS = [
    (r'\{\\"\{?([aouAOU])\}?\}', lambda m: {"a": "ä", "o": "ö", "u": "ü", "A": "Ä", "O": "Ö", "U": "Ü"}.get(m.group(1), m.group(1))),
    (r'\\"\{?([aouAOU])\}?', lambda m: {"a": "ä", "o": "ö", "u": "ü", "A": "Ä", "O": "Ö", "U": "Ü"}.get(m.group(1), m.group(1))),
    (r"\\'\{?([aeiouAEIOU])\}?", lambda m: m.group(1)),
    (r"\\`\{?([aeiouAEIOU])\}?", lambda m: m.group(1)),
    (r"\\~\{?([nN])\}?", lambda m: {"n": "ñ", "N": "Ñ"}.get(m.group(1), m.group(1))),
    (r"\\c\{?([cC])\}?", lambda m: {"c": "ç", "C": "Ç"}.get(m.group(1), m.group(1))),
    (r"\\v\{?([sSzZcC])\}?", lambda m: {"s": "š", "S": "Š", "z": "ž", "Z": "Ž", "c": "č", "C": "Č"}.get(m.group(1), m.group(1))),
    (r"\\textit\{([^}]*)\}", r"\1"),
    (r"\\textbf\{([^}]*)\}", r"\1"),
    (r"\\emph\{([^}]*)\}", r"\1"),
    (r"\\L\b", "Ł"),
    (r"--", "–"),
    (r"\{([^{}]*)\}", r"\1"),
    (r"\\&", "&"),
    (r"~", " "),
]


def _clean_latex(s: str) -> str:
    for pat, repl in _LATEX_REPLACEMENTS:
        s = re.sub(pat, repl, s)
    return s.strip()


def _extract_brace_value(text: str) -> tuple[str, str]:
    text = text.lstrip()
    if not text:
        return "", ""
    if text[0] == "{":
        depth = 0
        for i, c in enumerate(text):
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    return text[1:i], text[i + 1:]
        return text[1:], ""
    if text[0] == '"':
        end = text.find('"', 1)
        if end == -1:
            return text[1:], ""
        return text[1:end], text[end + 1:]
    m = re.match(r"([^,}\s]+)", text)
    if m:
        return m.group(1), text[m.end():]
    return "", text


def _parse_bib(path: Path):
    raw = path.read_text(encoding="utf-8", errors="replace")
    entries = []
    pos = 0
    while pos < len(raw):
        at = raw.find("@", pos)
        if at == -1:
            break
        brace_open = raw.find("{", at)
        if brace_open == -1:
            break
        etype = raw[at + 1:brace_open].strip().lower()
        depth = 1
        j = brace_open + 1
        while j < len(raw) and depth > 0:
            if raw[j] == "{":
                depth += 1
            elif raw[j] == "}":
                depth -= 1
            j += 1
        body = raw[brace_open + 1:j - 1]
        pos = j
        if etype in ("comment", "string", "preamble"):
            continue
        comma = body.find(",")
        if comma == -1:
            continue
        key = body[:comma].strip()
        rest = body[comma + 1:]
        fields = {}
        while rest:
            rest = rest.lstrip(" \t\r\n,")
            m = re.match(r"(\w+)\s*=\s*", rest)
            if not m:
                break
            fname = m.group(1).lower()
            rest = rest[m.end():]
            val, rest = _extract_brace_value(rest)
            fields[fname] = _clean_latex(val)
        entries.append((key, etype, fields))
    return entries


# --------------------------------------------------
# GOST author formatting
# --------------------------------------------------

def _initials(first_str: str) -> str:
    parts = first_str.strip().split()
    out = []
    for p in parts:
        p = p.strip(".")
        if p:
            out.append(p[0].upper() + ".")
    return "".join(out)


def _format_author(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        return ""
    if "," in raw:
        last, _, rest = raw.partition(",")
        ins = _initials(rest.strip())
        return f"{last.strip()} {ins}" if ins else last.strip()
    parts = raw.split()
    if len(parts) == 1:
        return parts[0]
    last = parts[-1]
    first = " ".join(parts[:-1])
    ins = _initials(first)
    return f"{last} {ins}" if ins else last


def _authors_gost(raw: str, max_full: int = 3) -> str:
    if not raw:
        return ""
    parts = [a.strip() for a in re.split(r"\s+and\s+", raw, flags=re.IGNORECASE)]
    et_al = len(parts) > max_full
    shown = [_format_author(p) for p in parts[:max_full]]
    s = ", ".join(shown)
    if et_al:
        s += " [et al.]"
    return s


def _format_entry(n: int, etype: str, fields: dict) -> str:
    def f(k):
        return fields.get(k, "").strip()

    auth = _authors_gost(f("author") or f("editor"))
    title = f("title") or "(no title)"
    year = f("year")
    url = f("url") or f("howpublished")
    if auth:
        auth_prefix = f"{auth} " if auth.endswith(".") else f"{auth}. "
    else:
        auth_prefix = ""

    if etype == "article":
        journal = f("journal")
        vol = f("volume")
        num = f("number")
        pages = f("pages")
        s = f"{auth_prefix}{title}"
        if journal:
            s += f" // {journal}"
        if year:
            s += f". — {year}"
        loc = []
        if vol:
            loc.append(f"Vol. {vol}")
        if num:
            loc.append(f"No. {num}")
        if pages:
            loc.append(f"P. {pages}")
        if loc:
            s += ". — " + ", ".join(loc)
    elif etype in ("book", "inbook"):
        publisher = f("publisher")
        address = f("address")
        pages = f("pages")
        s = f"{auth_prefix}{title}"
        loc = ", ".join(x for x in [address, publisher] if x)
        if loc:
            s += f". — {loc}"
        if year:
            s += f", {year}"
        if pages:
            s += f". — {pages} с."
    elif etype in ("inproceedings", "incollection", "conference"):
        booktitle = f("booktitle")
        pages = f("pages")
        s = f"{auth_prefix}{title}"
        if booktitle:
            s += f" // {booktitle}"
        if year:
            s += f". — {year}"
        if pages:
            s += f". — P. {pages}"
    elif etype == "phdthesis":
        school = f("school")
        s = f"{auth_prefix}{title}"
        if school:
            s += f". — {school}"
        if year:
            s += f", {year}"
    else:
        s = f"{auth_prefix}{title} [Электронный ресурс]"
        if url:
            s += f". — URL: {url}"
        urldate = f("urldate")
        if urldate:
            try:
                urldate = datetime.strptime(urldate, "%Y-%m-%d").strftime("%d.%m.%Y")
            except ValueError:
                pass
            s += f" (дата обращения: {urldate})"

    s = s.strip().rstrip(".")
    return f"{n}. {s}."


def _make_para_elem(text: str):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "Normal")
    pPr.append(pStyle)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "426")
    ind.set(qn("w:hanging"), "426")
    pPr.append(ind)
    p.append(pPr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _para_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(qn("w:t")))


def run(ctx: Context) -> None:
    doc = Document(ctx.paths.working_docx)
    body = doc.element.body
    children = list(body)
    title = ctx.config.strings.references_section

    bib_elem = None
    for child in children:
        if child.tag == qn("w:p") and title in _para_text(child):
            bib_elem = child
            break

    if bib_elem is None:
        raise RuntimeError(f"bibliography heading '{title}' not found in working docx")

    cited = _citation_order(ctx.paths.main_tex)
    print(f"  {len(cited)} citation key(s) from main.tex (first: {cited[:3]})")

    # Drop everything that follows the bibliography heading until the next Heading.
    to_remove = []
    node = bib_elem.getnext()
    while node is not None:
        nxt = node.getnext()
        if node.tag == qn("w:p"):
            pPr = node.find(qn("w:pPr"))
            if pPr is not None:
                ps = pPr.find(qn("w:pStyle"))
                if ps is not None and ps.get(qn("w:val"), "").lower().startswith("heading"):
                    break
        if node.tag != qn("w:bookmarkEnd"):
            to_remove.append(node)
        node = nxt
    for n in to_remove:
        body.remove(n)

    all_entries = _parse_bib(ctx.paths.bibliography_bib)
    entry_map = {key: (key, etype, fields) for key, etype, fields in all_entries}

    if cited:
        entries = [(k, entry_map[k][1], entry_map[k][2]) for k in cited if k in entry_map]
        if not entries:
            entries = all_entries
            source = "fallback: bib file order"
        else:
            source = "LaTeX first-mention order"
    else:
        entries = all_entries
        source = "fallback: bib file order"

    for n, (_key, etype, fields) in reversed(list(enumerate(entries, 1))):
        p = _make_para_elem(_format_entry(n, etype, fields))
        bib_elem.addnext(p)

    doc.save(ctx.paths.working_docx)
    print(f"  {len(entries)} bibliography entries ({source})")
