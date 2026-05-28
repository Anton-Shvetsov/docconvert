"""Final step: copy TMP/thesis.docx → output/{name}_{timestamp}.docx and clean TMP/.

The TMP/ directory is removed only when ``config.cleanup_tmp`` is true and
``--keep-tmp`` was not passed on the CLI.
"""
from __future__ import annotations

import shutil

from ..pipeline import Context


def run(ctx: Context) -> None:
    src = ctx.paths.working_docx
    if not src.exists():
        raise FileNotFoundError(
            f"Working docx not found: {src}. Run the full pipeline (or --from run_pandoc)."
        )

    ctx.paths.output_dir.mkdir(parents=True, exist_ok=True)
    dst = ctx.paths.final_output_path()
    shutil.copy2(src, dst)
    print(f"  → {dst}")

    if ctx.config.cleanup_tmp and not ctx.keep_tmp:
        shutil.rmtree(ctx.paths.tmp_dir, ignore_errors=True)
        print(f"  TMP/ removed (set cleanup_tmp=false or pass --keep-tmp to keep)")
