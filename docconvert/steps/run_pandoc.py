"""Convert TMP/main_flat.tex → TMP/thesis.docx via pandoc.

Runs from input/ so \\includegraphics{Figures/...} paths resolve. The flat tex
lives in TMP, the reference template (built earlier) lives in TMP too.
"""
from __future__ import annotations

import os

from ..pipeline import Context
from ..tools import require, run as run_cmd


def _resolve_template(filename: str, project_root, templates_dir) -> str:
    """Resolve a pandoc-related asset: prefer project root, fall back to templates/."""
    candidate = project_root / filename
    if candidate.exists():
        return str(candidate)
    fallback = templates_dir / filename
    if fallback.exists():
        return str(fallback)
    raise FileNotFoundError(
        f"Pandoc asset not found: {filename} (looked in {project_root} and {templates_dir})"
    )


def run(ctx: Context) -> None:
    pandoc = require("pandoc")
    require("pandoc-crossref")  # filter must be on PATH

    p = ctx.paths
    cfg = ctx.config.pandoc

    crossref = _resolve_template(cfg.crossref_yaml, p.project_root, p.templates_dir)
    csl = _resolve_template(cfg.csl, p.project_root, p.templates_dir)
    bib = p.bibliography_bib

    # Paths to pass to pandoc: relative to its cwd (input/).
    flat_tex_rel = os.path.relpath(p.flat_tex, p.input_dir)
    out_rel = os.path.relpath(p.working_docx, p.input_dir)
    reference_rel = os.path.relpath(p.reference_docx, p.input_dir)

    cmd = [
        pandoc,
        flat_tex_rel,
        "-o", out_rel,
        f"--bibliography={bib.name}",
        "--filter", "pandoc-crossref",
        f"--metadata-file={crossref}",
        "--citeproc",
        f"--csl={csl}",
        f"--reference-doc={reference_rel}",
    ] + list(cfg.extra_args)

    run_cmd(cmd, cwd=p.input_dir)
    print(f"  → {p.working_docx}")
