"""Load and validate config.yaml into typed dataclasses."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class StyleConfig:
    font: str = "Times New Roman"
    font_size_pt: int = 14
    line_spacing: float = 1.5
    margins_mm: dict[str, int] = field(default_factory=lambda: {"left": 30, "right": 10, "top": 20, "bottom": 20})
    paragraph_indent_cm: float = 1.25


@dataclass
class PandocConfig:
    crossref_yaml: str = "crossref.yaml"
    csl: str = "numeric.csl"
    reference_doc: str = "reference.docx"
    extra_args: list[str] = field(default_factory=lambda: ["--number-sections"])


@dataclass
class StringsConfig:
    abstract_section_aliases: list[str] = field(default_factory=lambda: ["РЕФЕРАТ", "АННОТАЦИЯ"])
    intro_section: str = "ВВЕДЕНИЕ"
    conclusion_section: str = "ЗАКЛЮЧЕНИЕ"
    references_section: str = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    figure_label: str = "Рисунок"
    table_label: str = "Таблица"
    where_keyword: str = "где"
    rpz_line_prefix: str = "Расчетно-пояснительная записка:"


@dataclass
class AppendixConfig:
    code: str
    title: str
    source: str
    convert: bool


@dataclass
class PageNumberingConfig:
    start_at: int = 6


@dataclass
class Config:
    raw: dict[str, Any]
    main_tex: str
    bibliography_bib: str
    titlepage_docx: str
    output_name: str
    cleanup_tmp: bool
    appendices: list[AppendixConfig]
    steps: dict[str, bool]
    strings: StringsConfig
    pandoc: PandocConfig
    style: StyleConfig
    page_numbering: PageNumberingConfig


def load_config(path: Path) -> Config:
    with path.open(encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    return Config(
        raw=raw,
        main_tex=raw.get("main_tex", "main.tex"),
        bibliography_bib=raw.get("bibliography_bib", "bibliography.bib"),
        titlepage_docx=raw.get("titlepage_docx", "titlist.docx"),
        output_name=raw.get("output_name", "thesis"),
        cleanup_tmp=bool(raw.get("cleanup_tmp", True)),
        appendices=[AppendixConfig(**a) for a in raw.get("appendices", [])],
        steps=dict(raw.get("steps", {})),
        strings=StringsConfig(**raw.get("strings", {})),
        pandoc=PandocConfig(**raw.get("pandoc", {})),
        style=StyleConfig(**raw.get("style", {})),
        page_numbering=PageNumberingConfig(**raw.get("page_numbering", {})),
    )
