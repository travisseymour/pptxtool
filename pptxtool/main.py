# SPDX-FileCopyrightText: PPTXTool | Copyright (c) 2026 Travis L. Seymour, PhD <nogard@ucsc.edu>
# SPDX-License-Identifier: GPL-3.0-or-later

import argparse
import os
from pathlib import Path
import sys

from . import __version__

os.environ.setdefault("QT_LOGGING_RULES", "qt.text.font.db=false")


def main() -> None:
    """Entry point for the pptxtool application."""
    parser = argparse.ArgumentParser(prog="pptxtool", description="GUI for preparing presentations and PDFs.")
    parser.add_argument("file", nargs="?", help="a .pptx or .pdf file to open")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args()

    from .gui import run_gui

    # Capture current working directory before Qt changes it
    cwd = str(Path.cwd())
    sys.exit(run_gui(initial_location=cwd, initial_file=args.file))


if __name__ == "__main__":
    main()
