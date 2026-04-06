#!/usr/bin/env python3
"""
Electoral Roll PDF → Excel (GUI) — launcher

The Tkinter UI lives in the website repo:
  Votabase-Website/pdf_to_excel/ui/app.py

From `votabase-backend`, run:
  python ui/app.py

Or from anywhere:
  python /path/to/Votabase-Website/pdf_to_excel/ui/app.py
"""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

_BACKEND_ROOT = Path(__file__).resolve().parent.parent  # votabase-backend
_REPO_ROOT = _BACKEND_ROOT.parent  # e.g. GitHub
_GUI = _REPO_ROOT / "Votabase-Website" / "pdf_to_excel" / "ui" / "app.py"

if not _GUI.is_file():
    print(
        "Electoral Roll GUI not found.\n"
        f"  Expected: {_GUI}\n"
        "  Clone or open the Votabase-Website repo next to votabase-backend, or run:\n"
        "  python pdf_to_excel/ui/app.py\n"
        "  from inside Votabase-Website/pdf_to_excel/",
        file=sys.stderr,
    )
    sys.exit(1)

_TK_HINT = """
Tkinter is not available in this Python build (missing _tkinter).

On macOS with Homebrew Python, Tcl/Tk is often not linked. Try one of:

  1) Install the Homebrew Tk binding for your Python version, then retry:
       brew install python-tk@3.14
     (Use the version that matches: python3 --version)

  2) Use the official installer from https://www.python.org/downloads/
     — it bundles Tcl/Tk and usually includes tkinter.

  3) Skip the GUI and use the CLI (same conversion logic):
       cd Votabase-Website/pdf_to_excel
       python3 scripts/pdf_to_excel.py --pdf /path/to/roll.pdf --out out.xlsx --booth 14
""".strip()


if __name__ == "__main__":
    try:
        runpy.run_path(str(_GUI), run_name="__main__")
    except ModuleNotFoundError as e:
        name = getattr(e, "name", "") or str(e)
        if "_tkinter" in str(e) or name == "_tkinter":
            print(_TK_HINT, file=sys.stderr)
            sys.exit(1)
        raise
