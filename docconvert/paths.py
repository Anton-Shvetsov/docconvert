"""Path resolution for the pipeline.

Centralises every absolute path the steps need so individual steps don't have
to know about config layout or repo root.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class Paths:
    project_root: Path
    input_dir: Path
    tmp_dir: Path
    output_dir: Path
    templates_dir: Path

    main_tex: Path
    bibliography_bib: Path
    titlepage_docx: Path

    flat_tex: Path
    reference_docx: Path
    working_docx: Path

    output_name: str

    @classmethod
    def from_config(cls, project_root: Path, cfg: dict) -> "Paths":
        root = project_root.resolve()
        paths_cfg = cfg.get("paths", {})
        input_dir = (root / paths_cfg.get("input_dir", "input")).resolve()
        tmp_dir = (root / paths_cfg.get("tmp_dir", "TMP")).resolve()
        output_dir = (root / paths_cfg.get("output_dir", "output")).resolve()
        templates_dir = (root / "templates").resolve()

        return cls(
            project_root=root,
            input_dir=input_dir,
            tmp_dir=tmp_dir,
            output_dir=output_dir,
            templates_dir=templates_dir,
            main_tex=input_dir / cfg.get("main_tex", "main.tex"),
            bibliography_bib=input_dir / cfg.get("bibliography_bib", "bibliography.bib"),
            titlepage_docx=input_dir / cfg.get("titlepage_docx", "titlist.docx"),
            flat_tex=tmp_dir / "main_flat.tex",
            reference_docx=tmp_dir / "reference.docx",
            working_docx=tmp_dir / "thesis.docx",
            output_name=cfg.get("output_name", "thesis"),
        )

    def appendix_source(self, source: str) -> Path:
        return self.input_dir / source

    def appendix_flat_tex(self, stem: str) -> Path:
        return self.tmp_dir / f"{stem}_flat.tex"

    def appendix_docx(self, stem: str) -> Path:
        return self.tmp_dir / f"{stem}.docx"

    def final_output_path(self, timestamp: datetime | None = None) -> Path:
        ts = (timestamp or datetime.now()).strftime("%Y-%m-%d_%H-%M-%S")
        return self.output_dir / f"{self.output_name}_{ts}.docx"

    def ensure_dirs(self) -> None:
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
