"""Build PDF via xelatex + bibtex + xelatex × 2.

Runs from ``input/`` so that \\input{} and \\includegraphics{Figures/...} resolve
correctly, and writes auxiliary files into ``TMP/`` via -output-directory.
"""
from __future__ import annotations

import os
import subprocess

from ..pipeline import Context
from ..tools import require, run as run_cmd


def _xelatex(ctx: Context, xelatex: str) -> None:
    run_cmd(
        [
            xelatex,
            f"-output-directory={ctx.paths.tmp_dir}",
            "-interaction=nonstopmode",
            ctx.paths.main_tex.name,
        ],
        cwd=ctx.paths.input_dir,
    )


def _bibtex(ctx: Context, bibtex: str) -> None:
    # bibtex finds main.aux in TMP/; TEXINPUTS/BIBINPUTS point back to input/
    # so it can resolve bibliography.bib referenced from the .aux file.
    env = os.environ.copy()
    env["BIBINPUTS"] = f"{ctx.paths.input_dir}{os.pathsep}"
    env["TEXINPUTS"] = f"{ctx.paths.input_dir}{os.pathsep}"
    cmd = [bibtex, ctx.paths.main_tex.stem]
    print(f"  $ {' '.join(cmd)}  (cwd={ctx.paths.tmp_dir.name})")
    result = subprocess.run(cmd, cwd=ctx.paths.tmp_dir, env=env)
    if result.returncode != 0:
        raise RuntimeError(f"bibtex failed (exit {result.returncode})")


def run(ctx: Context) -> None:
    xelatex = require("xelatex")
    bibtex = require("bibtex")
    _xelatex(ctx, xelatex)
    _bibtex(ctx, bibtex)
    _xelatex(ctx, xelatex)
    _xelatex(ctx, xelatex)
