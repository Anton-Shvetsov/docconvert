"""External tool discovery and subprocess wrappers."""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

REQUIRED_TOOLS = ("xelatex", "bibtex", "latexpand", "pandoc", "pandoc-crossref")


def which(tool: str) -> str | None:
    return shutil.which(tool)


def check_tools() -> dict[str, str | None]:
    return {t: which(t) for t in REQUIRED_TOOLS}


class ToolMissing(RuntimeError):
    pass


def require(tool: str) -> str:
    p = which(tool)
    if not p:
        raise ToolMissing(
            f"Required tool not found on PATH: {tool}. "
            f"Install it and ensure the executable is reachable."
        )
    return p


def run(cmd: list[str], cwd: Path | None = None, stdout_file: Path | None = None) -> None:
    """Run a subprocess and raise on non-zero exit. Streams output to stdout
    unless stdout_file is given (then captures into the file)."""
    pretty = " ".join(cmd)
    print(f"  $ {pretty}" + (f"  > {stdout_file.name}" if stdout_file else ""))
    if stdout_file is not None:
        with stdout_file.open("wb") as f:
            result = subprocess.run(cmd, cwd=cwd, stdout=f, stderr=subprocess.PIPE)
    else:
        result = subprocess.run(cmd, cwd=cwd)
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="replace") if stdout_file else ""
        raise RuntimeError(
            f"Command failed with exit {result.returncode}: {pretty}\n{stderr}"
        )
