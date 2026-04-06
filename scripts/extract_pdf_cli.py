#!/usr/bin/env python3
"""Extract a voter-list PDF to Excel using the same pipeline as POST /api/extract/pdf-to-excel.

Usage:
  python scripts/extract_pdf_cli.py /path/to/4_9_14_EPUB.pdf
  python scripts/extract_pdf_cli.py ./file.pdf ./out.xlsx --debug

Requires backend venv deps (pdfplumber, openpyxl). For scanned PDFs without AWS Textract,
install poppler + tesseract so OCR fallback can run.
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from app.extract import _run_pdf_extract  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF voter extract to Excel")
    parser.add_argument("pdf", help="Input PDF path")
    parser.add_argument("out", nargs="?", default=None, help="Output .xlsx path (default: same stem as PDF)")
    parser.add_argument("--debug", action="store_true", help="Include DEBUG sheet")
    args = parser.parse_args()

    pdf_path = os.path.abspath(args.pdf)
    if not os.path.isfile(pdf_path):
        print(f"File not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    excel_tmp, row_count = _run_pdf_extract(pdf_path, debug=args.debug)
    out_path = os.path.abspath(args.out or os.path.splitext(pdf_path)[0] + ".xlsx")
    try:
        shutil.copy2(excel_tmp, out_path)
    finally:
        try:
            os.unlink(excel_tmp)
        except OSError:
            pass

    print(f"Rows extracted: {row_count}")
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
