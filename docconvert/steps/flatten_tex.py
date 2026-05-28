"""Run latexpand on input/main.tex → TMP/main_flat.tex.

Runs from input/ so \\input{} paths resolve. latexpand prints to stdout, so we
capture into the target file.
"""
from __future__ import annotations

from ..pipeline import Context
from ..tools import require, run as run_cmd


def run(ctx: Context) -> None:
    latexpand = require("latexpand")
    run_cmd(
        [latexpand, ctx.paths.main_tex.name],
        cwd=ctx.paths.input_dir,
        stdout_file=ctx.paths.flat_tex,
    )
    print(f"  → {ctx.paths.flat_tex}")
