"""Command-line entry point for docconvert."""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from .config import load_config
from .paths import Paths
from .pipeline import Context, list_steps, run_pipeline
from .tools import REQUIRED_TOOLS, check_tools


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="docconvert", description="LaTeX→DOCX pipeline for GOST documents.")
    p.add_argument("--config", type=Path, default=Path("config.yaml"), help="Path to config.yaml (default: ./config.yaml)")
    p.add_argument("--workdir", type=Path, default=None, help="chdir here before running (default: directory of --config)")
    p.add_argument("--only", nargs="+", metavar="STEP", help="Run only these steps (in canonical order).")
    p.add_argument("--skip", nargs="+", metavar="STEP", help="Skip these steps.")
    p.add_argument("--from", dest="from_step", metavar="STEP", help="Start from this step and continue.")
    p.add_argument("--keep-tmp", action="store_true", help="Do not delete TMP/ in the finalize step.")
    p.add_argument("--list-steps", action="store_true", help="Print canonical step order and exit.")
    p.add_argument("--check-tools", action="store_true", help="Check that external tools are on PATH and exit.")
    return p


def _print_tool_check() -> int:
    found = check_tools()
    width = max(len(t) for t in REQUIRED_TOOLS)
    missing = 0
    for tool in REQUIRED_TOOLS:
        path = found[tool]
        status = path if path else "NOT FOUND"
        marker = "OK " if path else "!! "
        print(f"  {marker}{tool:<{width}}  {status}")
        if not path:
            missing += 1
    if missing:
        print(f"\n{missing} required tool(s) missing. See README for install instructions.")
        return 1
    print("\nAll required tools found.")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)

    if args.list_steps:
        for s in list_steps():
            print(s)
        return 0

    if args.check_tools:
        return _print_tool_check()

    config_path = args.config.resolve()
    if not config_path.exists():
        print(f"Config file not found: {config_path}", file=sys.stderr)
        return 1

    workdir = args.workdir.resolve() if args.workdir else config_path.parent
    os.chdir(workdir)

    config = load_config(config_path)
    paths = Paths.from_config(workdir, config.raw)

    ctx = Context(config=config, paths=paths, keep_tmp=args.keep_tmp)

    run_pipeline(ctx, only=args.only, skip=args.skip, from_step=args.from_step)
    return 0


if __name__ == "__main__":
    sys.exit(main())
