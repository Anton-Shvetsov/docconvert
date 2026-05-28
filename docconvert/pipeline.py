"""Pipeline orchestration: ordered list of steps + execution context."""
from __future__ import annotations

import importlib
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .config import Config
from .paths import Paths


# Canonical step order. Each entry: (key, module-name-under-docconvert.steps).
STEP_ORDER: list[tuple[str, str]] = [
    ("build_pdf", "build_pdf"),
    ("build_style_template", "build_style_template"),
    ("flatten_tex", "flatten_tex"),
    ("run_pandoc", "run_pandoc"),
    ("insert_toc", "insert_toc"),
    ("merge_appendices", "merge_appendices"),
    ("enforce_styles", "enforce_styles"),
    ("fix_tables", "fix_tables"),
    ("fix_blocks", "fix_blocks"),
    ("fix_sections", "fix_sections"),
    ("number_equations", "number_equations"),
    ("align_where_blocks", "align_where_blocks"),
    ("build_bibliography", "build_bibliography"),
    ("prepend_titlepage", "prepend_titlepage"),
    ("fill_referencat", "fill_referencat"),
    ("finalize", "finalize"),
]


@dataclass
class Context:
    config: Config
    paths: Paths
    keep_tmp: bool = False


def _load_step(module_name: str) -> Callable[[Context], None]:
    mod = importlib.import_module(f"docconvert.steps.{module_name}")
    if not hasattr(mod, "run"):
        raise RuntimeError(f"Step module docconvert.steps.{module_name} has no run() function")
    return mod.run


def list_steps() -> list[str]:
    return [key for key, _ in STEP_ORDER]


def _select_steps(
    cfg_steps: dict[str, bool],
    only: list[str] | None,
    skip: list[str] | None,
    from_step: str | None,
) -> list[tuple[str, str]]:
    if only:
        return [(k, m) for k, m in STEP_ORDER if k in only]

    selected = [(k, m) for k, m in STEP_ORDER if cfg_steps.get(k, True)]
    if skip:
        selected = [(k, m) for k, m in selected if k not in skip]
    if from_step:
        idx = next((i for i, (k, _) in enumerate(selected) if k == from_step), None)
        if idx is None:
            raise ValueError(f"--from step '{from_step}' is disabled in config or unknown")
        selected = selected[idx:]
    return selected


def _prepare_tmp(paths: Paths, from_step: str | None, only: list[str] | None) -> None:
    """Clean TMP at the start of a full run; keep it for partial runs."""
    paths.ensure_dirs()
    is_partial = bool(from_step or only)
    if not is_partial and paths.tmp_dir.exists():
        for child in paths.tmp_dir.iterdir():
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()


def run_pipeline(
    ctx: Context,
    *,
    only: list[str] | None = None,
    skip: list[str] | None = None,
    from_step: str | None = None,
) -> None:
    _prepare_tmp(ctx.paths, from_step, only)
    selected = _select_steps(ctx.config.steps, only, skip, from_step)

    if not selected:
        print("No steps selected — nothing to do.")
        return

    print(f"Running {len(selected)} step(s):")
    for key, _ in selected:
        print(f"  • {key}")
    print()

    for key, module_name in selected:
        t0 = time.perf_counter()
        print(f"== {key} ==")
        step_fn = _load_step(module_name)
        try:
            step_fn(ctx)
        except Exception as e:
            elapsed = time.perf_counter() - t0
            print(f"!! step '{key}' failed after {elapsed:.1f}s: {e}")
            print(f"!! TMP/ preserved for debugging: {ctx.paths.tmp_dir}")
            raise
        elapsed = time.perf_counter() - t0
        print(f"-- {key} done in {elapsed:.1f}s\n")
