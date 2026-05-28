#!/usr/bin/env python3
"""Thin launcher for the docconvert CLI."""
import sys

# Ensure step messages with Unicode (arrows, em-dashes, Cyrillic) render on Windows consoles.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from docconvert.cli import main

if __name__ == "__main__":
    sys.exit(main())
