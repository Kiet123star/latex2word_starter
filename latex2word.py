#!/usr/bin/env python3
"""LaTeX to Word CLI & Unified Core Wrapper (`latex2word.py`).

Uses the Unified Core Converter (`backend.core_converter`) to convert LaTeX documents into
editable Microsoft Word (.docx) files with:
- Native Word Shapes / DrawingML (<wpg:wgp>, <wps:wsp>) for TikZ & PGFPlots
- Native Editable Math Tables for `tkz-tab` variation tables, multiple-choice tables,
  statistical/data tables (`\multicolumn`, `\multirow`, `\diagbox`), and `\immini` layouts
- Native Office Math (OMML `<m:oMath>`) for equations
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from backend.core_converter import (
    clean_newlines,
    convert_docx_to_pdf,
    convert_latex_file_unified,
    docx_counts,
    expand_tex_includes,
)


def convert(
    tex: Path,
    output: Path,
    *,
    engine: str = "xelatex",
    dpi: int = 400,
    reference: Path | None = None,
    strict: bool = False,
    qa_pdf: bool = False,
    mode: int = 2,
) -> dict:
    """Convert a .tex file to .docx using the Unified Native DrawingML & Math Tables Core Engine."""
    return convert_latex_file_unified(
        tex,
        output,
        engine=engine,
        dpi=dpi,
        reference=reference,
        strict=strict,
        qa_pdf=qa_pdf,
        mode=mode,
    )


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("tex", type=Path, help="Main LaTeX .tex file")
    ap.add_argument("-o", "--output", type=Path, help="Destination .docx file")
    ap.add_argument("--engine", choices=["pdflatex", "xelatex", "lualatex"], default="xelatex")
    ap.add_argument("--dpi", type=int, default=400, help="Diagram resolution (default: 400)")
    ap.add_argument("--reference", type=Path, help="Custom reference.docx")
    ap.add_argument("--strict", action="store_true", help="Fail on validation issues")
    ap.add_argument("--qa-pdf", action="store_true", help="Render Word as PDF for review (LibreOffice / Word)")
    ap.add_argument("--mode", type=int, choices=[1, 2], default=2, help="Conversion mode (default: 2 - Max Editability)")
    args = ap.parse_args()
    out = args.output or args.tex.with_suffix(".docx")
    try:
        result = convert(
            args.tex,
            out,
            engine=args.engine,
            dpi=args.dpi,
            reference=args.reference,
            strict=args.strict,
            qa_pdf=args.qa_pdf,
            mode=args.mode,
        )
    except (RuntimeError, OSError, UnicodeError, zipfile.BadZipFile) as error:
        print("ERROR:", error, file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
