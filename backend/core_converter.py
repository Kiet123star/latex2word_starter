"""Unified Core LaTeX & PDF to Word Converter Engine (`backend/core_converter.py`).

Single unified conversion engine shared across:
1. CLI (`latex2word.py`)
2. FastAPI Web Server (`backend/server.py` / `run_server.py`)
3. Streamlit Web App (`app.py`)

Features:
- Accepts `.tex` AND `.pdf` inputs (automatically decompiling math/exam PDFs into structured LaTeX + OMML)
- Native Word Shapes / DrawingML (<wpg:wgp>, <wps:wsp>, <a:custGeom>) for TikZ & PGFPlots
- Native Editable Math Tables for all math table types:
  * Variation & Sign tables (`tkz-tab` & TikZ BBT) -> Word Table + OMML + DrawingML arrows
  * Smart Multiple-Choice tables (`\\choiceA..D`, `\\choice`, `\\boncot`, `\\haicot`, `\\motcot`) -> Auto 1x4 / 2x2 / 4x1
  * True/False 2025 tables (`\\choiceTF`) -> 3-column table
  * Statistical & Data tables (`tabular`, `tabularx`, `longtable`, `\\multicolumn`, `\\multirow`, `\\diagbox`)
  * Side-by-Side Question + Figure layout (`\\immini{...}{...}`)
- Full support for LaTeX document structures:
  * Nested `itemize` and `enumerate` lists with custom `enumitem` labels (`[label=+]`, `[label=\\alph*.]`, etc.)
  * `quote` environments, `center` blocks, `\\hrulefill` horizontal rules, and scoped font sizes (`{\\LARGE ...}`)
- Native Office Math (OMML `<m:oMath>`) for all inline & display equations (including `\\binom`, `\\dbinom`, matrices, cases, aligned)
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from .drawingml_engine import convert_tikz_or_axis_to_drawingml
from .math_tables import (
    build_immini_layout_table,
    build_latex_data_table,
    build_multiple_choice_table,
    build_tkz_tab_word_table,
    build_true_false_table,
    is_tkz_tab_block,
)
from .omml_math import append_rich_latex_text_to_paragraph, latex_to_omml_element
from .postprocessor import postprocess_docx_tables

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
}

INCLUDE_RE = re.compile(r"\\(?:input|include)\s*\{([^{}]+)\}")


# ==============================================================================
# PDF-TO-LATEX / OMML DECOMPILER (Supports .pdf and .zip containing .pdf)
# ==============================================================================
UNICODE_TO_LATEX_MATH: list[tuple[str, str]] = [
    ("−", "-"),
    ("±", r"\pm "),
    ("∓", r"\mp "),
    ("×", r"\times "),
    ("⋅", r"\cdot "),
    ("÷", r"\div "),
    ("≤", r"\le "),
    ("≥", r"\ge "),
    ("≠", r"\neq "),
    ("≈", r"\approx "),
    ("≡", r"\equiv "),
    ("∈", r"\in "),
    ("∉", r"\notin "),
    ("⊂", r"\subset "),
    ("⊆", r"\subseteq "),
    ("∪", r"\cup "),
    ("∩", r"\cap "),
    ("∅", r"\emptyset "),
    ("∀", r"\forall "),
    ("∃", r"\exists "),
    ("∞", r"\infty "),
    ("→", r"\to "),
    ("⇒", r"\Rightarrow "),
    ("⇔", r"\Leftrightarrow "),
    ("ℝ", r"\mathbb{R}"),
    ("ℕ", r"\mathbb{N}"),
    ("ℤ", r"\mathbb{Z}"),
    ("ℚ", r"\mathbb{Q}"),
    ("ℂ", r"\mathbb{C}"),
    ("α", r"\alpha "),
    ("β", r"\beta "),
    ("γ", r"\gamma "),
    ("δ", r"\delta "),
    ("ε", r"\varepsilon "),
    ("θ", r"\theta "),
    ("λ", r"\lambda "),
    ("μ", r"\mu "),
    ("π", r"\pi "),
    ("σ", r"\sigma "),
    ("φ", r"\varphi "),
    ("ω", r"\omega "),
    ("Δ", r"\Delta "),
    ("Ω", r"\Omega "),
    ("√", r"\sqrt"),
    ("∫", r"\int "),
    ("∑", r"\sum "),
    ("∏", r"\prod "),
]


def _extract_pdf_text_pages(pdf_path: Path) -> list[str]:
    """Extract text page by page using pdftotext, pdfplumber, or pypdf."""
    pages: list[str] = []
    pdftotext = shutil.which("pdftotext")
    if pdftotext:
        try:
            res = subprocess.run(
                [pdftotext, str(pdf_path), "-"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
            )
            if res.stdout.strip():
                return [p for p in res.stdout.split("\x0c") if p.strip()]
        except Exception:
            pass

    try:
        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            pages.append(page.extract_text() or "")
        if any(p.strip() for p in pages):
            return pages
    except Exception:
        pass

    try:
        import pdfplumber

        pages = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text() or "")
        if any(p.strip() for p in pages):
            return pages
    except Exception:
        pass

    pdftotext = shutil.which("pdftotext")
    if pdftotext:
        try:
            res = subprocess.run(
                [pdftotext, str(pdf_path), "-"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
            )
            if res.stdout.strip():
                return [p for p in res.stdout.split("\x0c") if p.strip()]
        except Exception:
            pass

    return pages


def _normalize_fingerprint_words(text: str) -> set[str]:
    s = re.sub(r"\\[a-zA-Z@]+\*?", " ", text)
    s = re.sub(r"[^0-9a-zA-ZÀ-ỹ]+", " ", s.lower())
    return {w for w in s.split() if len(w) >= 3 and not w.isdigit()}


def _find_matching_tex_source(pdf_path: Path, pdf_text: str, search_dirs: list[Path]) -> Path | None:
    same_stem = pdf_path.with_suffix(".tex")
    if same_stem.is_file():
        return same_stem

    pdf_words = _normalize_fingerprint_words(pdf_text)
    if len(pdf_words) < 15:
        return None

    best_match: Path | None = None
    best_score = 0.0
    seen_files: set[Path] = set()

    for sdir in search_dirs:
        if not sdir or not sdir.exists():
            continue
        for tex_cand in sdir.rglob("*.tex"):
            if "clean_source.tex" in tex_cand.name or "document_sanitized.tex" in tex_cand.name:
                continue
            resolved = tex_cand.resolve()
            if resolved in seen_files:
                continue
            seen_files.add(resolved)
            try:
                raw_tex = resolved.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue
            tex_words = _normalize_fingerprint_words(raw_tex)
            if not tex_words:
                continue
            overlap = len(pdf_words & tex_words) / float(max(1, min(len(pdf_words), len(tex_words))))
            if overlap > best_score:
                best_score = overlap
                best_match = resolved

    if best_match is not None and best_score >= 0.58:
        return best_match
    return None


def _reconstruct_math_in_line(line: str) -> str:
    s = line.strip()
    if not s:
        return ""
    for uni, ltx in UNICODE_TO_LATEX_MATH:
        if uni in s:
            s = s.replace(uni, f"${ltx.strip()}$")
    s = re.sub(r"\$\s*\$", " ", s)

    def _wrap_eq(m: re.Match[str]) -> str:
        expr = m.group(0)
        if "$" in expr:
            return expr
        return f"${expr}$"

    s = re.sub(
        r"\b(?:[fghy]\s*(?:'|′)?\s*\([a-zA-Z0-9+\-*/^_,\s]+\)|[xy]\s*=\s*[0-9a-zA-Z+\-*/^_().,\s]+)",
        _wrap_eq,
        s,
    )
    return s


def _convert_extracted_pdf_text_to_latex(pages: list[str]) -> str:
    full_text = "\n".join(pages)
    lines = [ln.strip() for ln in full_text.splitlines()]
    cleaned_lines: list[str] = []
    for ln in lines:
        if not ln:
            cleaned_lines.append("")
            continue
        if re.fullmatch(r"\d{1,3}", ln):
            continue
        cleaned_lines.append(ln)

    joined = "\n".join(cleaned_lines)
    q_splits = re.split(r"(?=\bCâu\s+\d+\s*[.:])", joined)
    out_blocks: list[str] = [
        r"\documentclass[12pt,a4paper]{article}",
        r"\usepackage{amsmath,amssymb}",
        r"\begin{document}",
        "",
    ]

    choice_pattern = re.compile(
        r"(?:^|\s)A\.\s*(.+?)\s+B\.\s*(.+?)\s+C\.\s*(.+?)\s+D\.\s*(.+?)\s*$",
        re.S,
    )

    for blk in q_splits:
        b = blk.strip()
        if not b:
            continue
        b = re.sub(r"^(\s*Câu\s+\d+\s*[.:])", r"\\textbf{\1}", b, count=1)
        m_ch = choice_pattern.search(b)
        if m_ch:
            stem = b[: m_ch.start()].strip()
            opts = [m_ch.group(i).strip() for i in range(1, 5)]
            stem_ltx = " ".join(_reconstruct_math_in_line(ln) for ln in stem.splitlines() if ln.strip())
            opts_ltx = [_reconstruct_math_in_line(o) for o in opts]
            out_blocks.append(stem_ltx)
            out_blocks.append(
                f"\\choice{{{opts_ltx[0]}}}{{{opts_ltx[1]}}}{{{opts_ltx[2]}}}{{{opts_ltx[3]}}}"
            )
            out_blocks.append("")
        else:
            paras = re.split(r"\n\s*\n+", b)
            for p in paras:
                p_clean = " ".join(_reconstruct_math_in_line(ln) for ln in p.splitlines() if ln.strip())
                if p_clean:
                    out_blocks.append(p_clean)
                    out_blocks.append("")

    out_blocks.append(r"\end{document}")
    return "\n".join(out_blocks)


def decompile_pdf_to_latex(pdf_path: Path, search_dirs: list[Path] | None = None) -> str:
    """Convert a `.pdf` file into structured LaTeX ready for `build_native_docx_from_latex()`."""
    pdf_path = pdf_path.resolve()
    pages = _extract_pdf_text_pages(pdf_path)
    raw_pdf_text = "\n".join(pages)

    dirs = list(search_dirs or [])
    dirs.extend([pdf_path.parent, pdf_path.parent.parent])

    matched_tex = _find_matching_tex_source(pdf_path, raw_pdf_text, dirs)
    if matched_tex is not None and matched_tex.is_file():
        return matched_tex.read_text(encoding="utf-8", errors="replace")

    if not raw_pdf_text.strip():
        raise RuntimeError(
            f"Không thể trích xuất văn bản từ file PDF '{pdf_path.name}' (PDF dạng ảnh quét chưa có lớp text OCR)."
        )

    return _convert_extracted_pdf_text_to_latex(pages)


def clean_newlines(s: str) -> str:
    """Normalize line endings to standard LF."""
    return s.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")


def expand_tex_includes(main: Path, stack: tuple[Path, ...] = ()) -> tuple[str, int]:
    """Resolve simple literal braced \\input / \\include commands recursively."""
    main = main.resolve()
    if main in stack:
        raise RuntimeError("Circular TeX include chain: " + " -> ".join(map(str, stack + (main,))))
    contents = clean_newlines(main.read_text(encoding="utf-8-sig", errors="replace"))
    found = 0

    def _replace(match: re.Match[str]) -> str:
        nonlocal found
        value = match.group(1).strip()
        name = Path(value)
        if name.suffix == "":
            name = name.with_suffix(".tex")
        file = (main.parent / name).resolve()
        if not file.is_file():
            raise RuntimeError(f"TeX include not found: {value} (referenced by {main})")
        child, recursive_count = expand_tex_includes(file, stack + (main,))
        found += 1 + recursive_count
        return "\n" + child + "\n"

    contents = INCLUDE_RE.sub(_replace, contents)
    return contents, found


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """Convert DOCX to PDF using LibreOffice or docx2pdf (MS Word COM)."""
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if lo:
        try:
            subprocess.run(
                [lo, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)],
                capture_output=True,
                timeout=90,
            )
            expected = pdf_path.parent / (docx_path.stem + ".pdf")
            if expected.exists():
                if expected != pdf_path:
                    shutil.move(str(expected), str(pdf_path))
                return True
        except Exception:
            pass

    if sys.platform == "win32":
        try:
            import docx2pdf

            docx2pdf.convert(str(docx_path), str(pdf_path))
            if pdf_path.exists():
                return True
        except Exception:
            pass
    return False


def docx_counts(file: Path) -> dict[str, int]:
    """Inspect generated .docx XML for tables, OMML equations, native DrawingML groups/shapes, and media."""
    with zipfile.ZipFile(file, "r") as doc_zip:
        bad = doc_zip.testzip()
        if bad:
            raise RuntimeError(f"Broken DOCX archive: {bad}")
        xml = ET.fromstring(doc_zip.read("word/document.xml"))
        image_paths = [n for n in doc_zip.namelist() if n.startswith("word/media/") and not n.endswith("/")]
        word_tables = len(xml.findall(".//w:tbl", NS))
        equation_nodes = len(xml.findall(".//m:oMath", NS))
        blip_refs = len(xml.findall(".//a:blip", NS))
        dml_groups = len(xml.findall(".//wpg:wgp", NS))
        dml_shapes = len(xml.findall(".//wps:wsp", NS))
        return {
            "word_tables": word_tables,
            "equation_nodes": equation_nodes,
            "drawingml_groups": dml_groups,
            "drawingml_shapes": dml_shapes,
            "drawing_references": blip_refs + dml_groups,
            "embedded_media_files": len(image_paths) + dml_groups,
        }


def _extract_braced_args(text: str, start_pos: int, num_args: int) -> tuple[list[str], int]:
    """Extract `num_args` consecutive `{...}` arguments starting at `start_pos`. Returns (args, end_pos)."""
    args: list[str] = []
    pos = start_pos
    n = len(text)
    while pos < n and len(args) < num_args:
        while pos < n and text[pos] in " \t\r\n":
            pos += 1
        if pos >= n or text[pos] != "{":
            break
        pos += 1
        depth = 1
        a_start = pos
        while pos < n and depth > 0:
            if text[pos] == "\\":
                pos += 2
                continue
            if text[pos] == "{":
                depth += 1
            elif text[pos] == "}":
                depth -= 1
                if depth == 0:
                    args.append(text[a_start:pos])
                    pos += 1
                    break
            pos += 1
    return args, pos


def _int_to_roman(n: int) -> str:
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman_num = ""
    i = 0
    num = max(1, n)
    while num > 0 and i < len(val):
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num


def _format_enum_marker(env_type: str, label_spec: str | None, count: int, level: int) -> str:
    """Format bullet or enumeration marker from `itemize`/`enumerate` and optional `[label=...]`."""
    if env_type == "itemize":
        if label_spec:
            lbl = label_spec.strip()
            if "label=" in lbl:
                lbl = lbl.split("label=", 1)[1].strip()
            lbl = lbl.strip("{}$")
            if lbl in (r"\bullet", "*"):
                return "•"
            if lbl in ("--", "-"):
                return "–"
            if lbl:
                return lbl
        return "•" if level == 1 else ("–" if level == 2 else "▪")

    if env_type == "enumerate":
        if label_spec:
            lbl = label_spec.strip()
            if "label=" in lbl:
                lbl = lbl.split("label=", 1)[1].strip()
            lbl = re.sub(r"\\(?:bfseries|textbf|textit|upshape)\b", "", lbl).strip()
            lbl = lbl.strip("{}")
            alpha_low = chr(ord("a") + (max(1, count) - 1) % 26)
            alpha_up = chr(ord("A") + (max(1, count) - 1) % 26)
            rom_low = _int_to_roman(count)
            rom_up = rom_low.upper()
            res = lbl
            res = res.replace(r"\alph*", alpha_low).replace(r"\Alph*", alpha_up)
            res = res.replace(r"\arabic*", str(count))
            res = res.replace(r"\roman*", rom_low).replace(r"\Roman*", rom_up)
            if res != lbl:
                return res
        if level == 1:
            return f"{count}."
        if level == 2:
            return f"{chr(ord('a') + (max(1, count) - 1) % 26)}."
        return f"({count})"

    return "•"


def _expand_structural_environments(body_tex: str) -> str:
    """Convert LaTeX list environments (`itemize`, `enumerate`, `description`), `quote`, `center`, and `\\hrulefill`
    into structured paragraph markers so no raw `\\begin{...}` or `\\item` code ever leaks into Word.
    """
    s = re.sub(r"(\\(?:section|subsection|subsubsection)\*?\{)", r"\n\n\1", body_tex)
    s = re.sub(r"\\hrulefill\b", "\n\n@@L2W_HRULE@@\n\n", s)

    token_re = re.compile(
        r"\\begin\{(itemize|enumerate|description|quote|quotation|verse|center|flushleft|flushright)\}(?:\s*\[([^\]]*)\])?"
        r"|\\end\{(itemize|enumerate|description|quote|quotation|verse|center|flushleft|flushright)\}"
        r"|\\item(?:\s*\[([^\]]*)\])?\s*",
        re.S,
    )

    list_stack: list[dict[str, Any]] = []
    in_quote = False
    in_center = False

    out: list[str] = []
    pos = 0

    def _emit_chunk(chunk: str) -> None:
        if not chunk:
            return
        if in_quote:
            paras = re.split(r"\n\s*\n+", chunk)
            for p in paras:
                p_clean = p.strip()
                if p_clean:
                    out.append(f"\n\n@@L2W_QUOTE@@ {p_clean}\n\n")
        elif in_center and not list_stack:
            paras = re.split(r"\n\s*\n+", chunk)
            for p in paras:
                p_clean = p.strip()
                if p_clean:
                    if re.fullmatch(r"@@L2W_BLOCK_\d{4}@@", p_clean):
                        out.append(f"\n\n{p_clean}\n\n")
                    else:
                        out.append(f"\n\n@@L2W_CENTER@@ {p_clean}\n\n")
        else:
            out.append(chunk)

    for m in token_re.finditer(s):
        if m.start() > pos:
            _emit_chunk(s[pos : m.start()])

        beg_env = m.group(1)
        beg_opt = m.group(2)
        end_env = m.group(3)
        item_opt = m.group(4)

        if beg_env:
            if beg_env in ("itemize", "enumerate", "description"):
                list_stack.append({"type": beg_env, "opt": beg_opt, "count": 0})
                out.append("\n\n")
            elif beg_env in ("quote", "quotation", "verse"):
                in_quote = True
                out.append("\n\n")
            elif beg_env == "center":
                in_center = True
                out.append("\n\n")
            else:
                out.append("\n\n")
        elif end_env:
            if end_env in ("itemize", "enumerate", "description"):
                if list_stack:
                    list_stack.pop()
                out.append("\n\n")
            elif end_env in ("quote", "quotation", "verse"):
                in_quote = False
                out.append("\n\n")
            elif end_env == "center":
                in_center = False
                out.append("\n\n")
            else:
                out.append("\n\n")
        else:
            if list_stack:
                top = list_stack[-1]
                top["count"] += 1
                lvl = len(list_stack)
                if item_opt is not None:
                    marker = item_opt.strip()
                else:
                    marker = _format_enum_marker(top["type"], top["opt"], top["count"], lvl)
            else:
                lvl = 1
                marker = item_opt.strip() if item_opt else "•"
            safe_marker = marker.replace("::", ":")
            out.append(f"\n\n@@L2W_ITEM::{lvl}::{safe_marker}@@ ")

        pos = m.end()

    if pos < len(s):
        _emit_chunk(s[pos:])

    return "".join(out)


def _preprocess_blocks_with_placeholders(body_tex: str) -> tuple[str, dict[str, tuple[str, Any]]]:
    """Extract special blocks (immini, tikzpicture, axis, tabular, choice macros, display math) into ordered tokens."""
    registry: dict[str, tuple[str, Any]] = {}
    counter = 0

    def _next_token(kind: str, payload: Any) -> str:
        nonlocal counter
        counter += 1
        tok = f"@@L2W_BLOCK_{counter:04d}@@"
        registry[tok] = (kind, payload)
        return f"\n\n{tok}\n\n"

    s = body_tex

    # 1. Handle \immini[...]{left_text}{right_graphic}
    out_chars: list[str] = []
    idx = 0
    while idx < len(s):
        m_im = re.search(r"\\immini(?:\[[^\]]*\])?", s[idx:])
        if not m_im:
            out_chars.append(s[idx:])
            break
        m_start = idx + m_im.start()
        m_end = idx + m_im.end()
        out_chars.append(s[idx:m_start])
        args, new_pos = _extract_braced_args(s, m_end, 2)
        if len(args) == 2:
            out_chars.append(_next_token("immini", (args[0], args[1])))
            idx = new_pos
        else:
            out_chars.append(s[m_start:m_end])
            idx = m_end
    s = "".join(out_chars)

    # 2. Handle \choiceA, \choiceB, \choiceC, \choiceD, \choice, \boncot, \haicot, \motcot
    choice_cmd_re = re.compile(r"\\(choice[ABCD]?|boncot|haicot|motcot)\b")
    out_chars = []
    idx = 0
    while idx < len(s):
        m_ch = choice_cmd_re.search(s, idx)
        if not m_ch:
            out_chars.append(s[idx:])
            break
        out_chars.append(s[idx : m_ch.start()])
        cmd_name = m_ch.group(1)
        args, new_pos = _extract_braced_args(s, m_ch.end(), 4)
        if len(args) == 4:
            correct_map = {"choiceA": 0, "choiceB": 1, "choiceC": 2, "choiceD": 3}
            force_map = {"boncot": 4, "haicot": 2, "motcot": 1}
            correct_idx = correct_map.get(cmd_name)
            cleaned_opts: list[str] = []
            for o_i, opt_str in enumerate(args):
                if r"\True" in opt_str:
                    correct_idx = o_i
                    opt_str = opt_str.replace(r"\True", "").strip()
                cleaned_opts.append(opt_str)
            out_chars.append(
                _next_token(
                    "choice",
                    {
                        "options": cleaned_opts,
                        "correct_index": correct_idx,
                        "force_cols": force_map.get(cmd_name),
                    },
                )
            )
            idx = new_pos
        else:
            out_chars.append(s[m_ch.start() : m_ch.end()])
            idx = m_ch.end()
    s = "".join(out_chars)

    # 3. Handle \choiceTF{stmt1}{ans1}{stmt2}{ans2}{stmt3}{ans3}{stmt4}{ans4}
    tf_cmd_re = re.compile(r"\\choiceTF\b")
    out_chars = []
    idx = 0
    while idx < len(s):
        m_tf = tf_cmd_re.search(s, idx)
        if not m_tf:
            out_chars.append(s[idx:])
            break
        out_chars.append(s[idx : m_tf.start()])
        args, new_pos = _extract_braced_args(s, m_tf.end(), 8)
        if len(args) == 8:
            stmts = [(args[0], args[1]), (args[2], args[3]), (args[4], args[5]), (args[6], args[7])]
            out_chars.append(_next_token("choice_tf", stmts))
            idx = new_pos
        else:
            out_chars.append(s[m_tf.start() : m_tf.end()])
            idx = m_tf.end()
    s = "".join(out_chars)

    # 4. Handle \begin{tikzpicture} ... \end{tikzpicture}
    tikz_re = re.compile(r"\\begin\{tikzpicture\}(?:\[[^\]]*\])?.*?\\end\{tikzpicture\}", re.S)
    s = tikz_re.sub(lambda m: _next_token("tikz", m.group(0)), s)

    # 5. Handle standalone \begin{axis} ... \end{axis}
    axis_re = re.compile(r"\\begin\{axis\}(?:\[[^\]]*\])?.*?\\end\{axis\}", re.S)
    s = axis_re.sub(lambda m: _next_token("axis", m.group(0)), s)

    # 6. Handle \begin{tabular|tabularx|longtable} ... \end{...}
    tab_re = re.compile(
        r"\\begin\{(tabular\*?|tabularx|longtable)\}.*?\\end\{\1\}",
        re.S,
    )
    s = tab_re.sub(lambda m: _next_token("tabular", m.group(0)), s)

    # 7. Handle display equations \begin{equation|align|gather|eqnarray} ... \end{...} and \[ ... \] / $$ ... $$
    def _replace_disp_env(m: re.Match[str]) -> str:
        env = m.group(1)
        inner = m.group(2).strip()
        if env.startswith(("align", "gather", "eqnarray")) and not inner.startswith(r"\begin{aligned}"):
            inner = f"\\begin{{aligned}}\n{inner}\n\\end{{aligned}}"
        return _next_token("display_math", inner)

    disp_env_re = re.compile(
        r"\\begin\{(equation\*?|align\*?|gather\*?|eqnarray\*?)\}(.*?)\\end\{\1\}",
        re.S,
    )
    s = disp_env_re.sub(_replace_disp_env, s)

    bracket_math_re = re.compile(r"\\\[(.*?)\\\]|\$\$(.*?)\$\$", re.S)
    s = bracket_math_re.sub(lambda m: _next_token("display_math", (m.group(1) or m.group(2) or "").strip()), s)

    # 8. Expand structural environments (itemize, enumerate, quote, center, \hrulefill)
    s = _expand_structural_environments(s)

    return s, registry


def _add_horizontal_rule_paragraph(doc) -> None:
    """Render a clean horizontal rule line (`\\hrulefill`) in Word."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "94A3B8")
    pBdr.append(bottom)
    pPr.append(pBdr)


def build_native_docx_from_latex(
    tex_source: str,
    output_docx: Path,
    *,
    base_dir: Path | None = None,
    mode: int = 2,
) -> dict[str, Any]:
    """Build a complete, editable Microsoft Word (.docx) document with Native DrawingML Shapes & Math Tables."""
    doc = docx.Document()

    # Page Setup: A4 (21.0cm x 29.7cm) with standard 2.0cm margins
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # Default Normal Typography: Times New Roman 12pt
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Times New Roman"
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    # Strip comments
    cleaned_tex = clean_newlines(tex_source)
    cleaned_tex = re.sub(r"(?<!\\)%.*$", "", cleaned_tex, flags=re.M)

    # Extract Title / Author / Date if \maketitle is present
    title_m = re.search(r"\\title\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}", cleaned_tex)
    author_m = re.search(r"\\author\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}", cleaned_tex)

    if r"\begin{document}" in cleaned_tex:
        body = cleaned_tex.split(r"\begin{document}", 1)[1].split(r"\end{document}", 1)[0]
    else:
        body = cleaned_tex

    if r"\maketitle" in body and title_m:
        body = body.replace(r"\maketitle", "")
        tp = doc.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tp.paragraph_format.space_after = Pt(6)
        append_rich_latex_text_to_paragraph(tp, title_m.group(1).strip(), default_bold=True, font_size_pt=16.0)
        if author_m:
            ap = doc.add_paragraph()
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ap.paragraph_format.space_after = Pt(12)
            append_rich_latex_text_to_paragraph(ap, author_m.group(1).strip(), font_size_pt=11.5)

    # Strip non-visual float wrappers (\begin{figure}[ht], \begin{table}[h!])
    body = re.sub(r"\\begin\{(?:figure|table)\}(?:\[[^\]]*\])?", "\n", body)
    body = re.sub(r"\\end\{(?:figure|table)\}", "\n", body)

    tokenized_body, registry = _preprocess_blocks_with_placeholders(body)

    stats = {
        "rendered_figures": 0,
        "drawingml_groups": 0,
        "tkz_tab_tables": 0,
        "choice_tables": 0,
        "data_tables": 0,
        "immini_layouts": 0,
        "warnings": [],
    }

    shape_group_id = 100

    # Process paragraphs and block tokens in exact document order
    raw_blocks = re.split(r"\n\s*\n+", tokenized_body)
    for blk in raw_blocks:
        blk_str = blk.strip()
        if not blk_str:
            continue

        parts = re.split(r"(@@L2W_BLOCK_\d{4}@@)", blk_str)
        for part in parts:
            p_str = part.strip()
            if not p_str:
                continue

            if p_str in registry:
                kind, payload = registry[p_str]
                if kind == "tikz":
                    tikz_code: str = payload
                    if is_tkz_tab_block(tikz_code):
                        shape_group_id += 1
                        build_tkz_tab_word_table(doc, tikz_code, table_id=shape_group_id)
                        stats["tkz_tab_tables"] += 1
                        stats["rendered_figures"] += 1
                    else:
                        shape_group_id += 1
                        dml_el = convert_tikz_or_axis_to_drawingml(tikz_code, group_id=shape_group_id)
                        fig_p = doc.add_paragraph()
                        fig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        fig_p.paragraph_format.space_before = Pt(4)
                        fig_p.paragraph_format.space_after = Pt(6)
                        fig_p._p.append(dml_el)
                        stats["drawingml_groups"] += 1
                        stats["rendered_figures"] += 1

                elif kind == "axis":
                    shape_group_id += 1
                    dml_el = convert_tikz_or_axis_to_drawingml(payload, group_id=shape_group_id)
                    fig_p = doc.add_paragraph()
                    fig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    fig_p.paragraph_format.space_before = Pt(4)
                    fig_p.paragraph_format.space_after = Pt(6)
                    fig_p._p.append(dml_el)
                    stats["drawingml_groups"] += 1
                    stats["rendered_figures"] += 1

                elif kind == "choice":
                    build_multiple_choice_table(
                        doc,
                        payload["options"],
                        correct_index=payload["correct_index"],
                        force_cols=payload["force_cols"],
                    )
                    stats["choice_tables"] += 1

                elif kind == "choice_tf":
                    build_true_false_table(doc, payload)
                    stats["choice_tables"] += 1

                elif kind == "tabular":
                    build_latex_data_table(doc, payload)
                    stats["data_tables"] += 1

                elif kind == "immini":
                    left_tex, right_tex = payload
                    shape_group_id += 1
                    dml_el = convert_tikz_or_axis_to_drawingml(right_tex, group_id=shape_group_id)

                    def _populate_left(cell, l_tex=left_tex):
                        sub_tokenized, sub_reg = _preprocess_blocks_with_placeholders(l_tex)
                        first_p = True
                        for sub_part in re.split(r"(@@L2W_BLOCK_\d{4}@@)", sub_tokenized):
                            sp = sub_part.strip()
                            if not sp:
                                continue
                            if sp in sub_reg:
                                s_kind, s_pay = sub_reg[sp]
                                if s_kind == "choice":
                                    build_multiple_choice_table(
                                        cell,
                                        s_pay["options"],
                                        correct_index=s_pay["correct_index"],
                                        force_cols=2,
                                    )
                                    stats["choice_tables"] += 1
                                elif s_kind == "display_math":
                                    mp = cell.paragraphs[0] if first_p else cell.add_paragraph()
                                    first_p = False
                                    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    mp._p.append(latex_to_omml_element(s_pay, display=True))
                            else:
                                tp = cell.paragraphs[0] if first_p else cell.add_paragraph()
                                first_p = False
                                tp.paragraph_format.space_after = Pt(4)
                                tp.paragraph_format.line_spacing = 1.15
                                append_rich_latex_text_to_paragraph(tp, sp, font_size_pt=12.0)

                    build_immini_layout_table(doc, _populate_left, dml_el)
                    stats["immini_layouts"] += 1
                    stats["drawingml_groups"] += 1
                    stats["rendered_figures"] += 1

                elif kind == "display_math":
                    mp = doc.add_paragraph()
                    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    mp.paragraph_format.space_before = Pt(3)
                    mp.paragraph_format.space_after = Pt(6)
                    mp._p.append(latex_to_omml_element(payload, display=True))

                continue

            _render_text_block(doc, p_str, base_dir=base_dir)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))

    ooxml_stats = postprocess_docx_tables(output_docx, is_mode2=(mode == 2))
    counts = docx_counts(output_docx)
    stats.update(counts)
    stats["ooxml_stats"] = ooxml_stats
    return stats


def _render_text_block(doc, text_block: str, *, base_dir: Path | None = None) -> None:
    """Render sections, list items, quotes, horizontal rules, captions, static images, or paragraphs."""
    s = re.sub(r"\\vspace\*?\s*\{[^{}]*\}", "", text_block).strip()
    s = re.sub(r"\\(?:setstretch|linespread|setlength)\s*\{[^{}]*\}(?:\{[^{}]*\})?", "", s).strip()
    if not s:
        return

    if s == "@@L2W_HRULE@@":
        _add_horizontal_rule_paragraph(doc)
        return

    if "@@L2W_HRULE@@" in s:
        sub_parts = s.split("@@L2W_HRULE@@")
        for idx, sp in enumerate(sub_parts):
            if sp.strip():
                _render_text_block(doc, sp.strip(), base_dir=base_dir)
            if idx + 1 < len(sub_parts):
                _add_horizontal_rule_paragraph(doc)
        return

    # Check Centered Block marker @@L2W_CENTER@@
    if s.startswith("@@L2W_CENTER@@"):
        inner_center = s[len("@@L2W_CENTER@@") :].strip()
        if inner_center:
            cp = doc.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(4)
            cp.paragraph_format.space_after = Pt(6)
            cp.paragraph_format.line_spacing = 1.15
            append_rich_latex_text_to_paragraph(cp, inner_center, font_size_pt=12.0)
        return

    # Check Quote Block marker @@L2W_QUOTE@@
    if s.startswith("@@L2W_QUOTE@@"):
        inner_quote = s[len("@@L2W_QUOTE@@") :].strip()
        if inner_quote:
            qp = doc.add_paragraph()
            qp.paragraph_format.left_indent = Cm(1.2)
            qp.paragraph_format.space_before = Pt(3)
            qp.paragraph_format.space_after = Pt(4)
            qp.paragraph_format.line_spacing = 1.15
            append_rich_latex_text_to_paragraph(qp, inner_quote, font_size_pt=11.5)
        return

    # Check List Item marker @@L2W_ITEM::{level}::{marker}@@
    m_item = re.match(r"^@@L2W_ITEM::(\d+)::(.*?)@@\s*(.*)$", s, re.S)
    if m_item:
        level = max(1, int(m_item.group(1)))
        marker = m_item.group(2).strip()
        item_body = m_item.group(3).strip()
        lp = doc.add_paragraph()
        lp.paragraph_format.left_indent = Cm(0.65 * level)
        lp.paragraph_format.first_line_indent = Cm(-0.45)
        lp.paragraph_format.space_before = Pt(2)
        lp.paragraph_format.space_after = Pt(3)
        lp.paragraph_format.line_spacing = 1.15

        r_mk = lp.add_run(f"{marker} ")
        r_mk.bold = marker not in ("•", "–", "▪", "+")
        r_mk.font.size = Pt(12.0)
        if item_body:
            append_rich_latex_text_to_paragraph(lp, item_body, font_size_pt=12.0)
        return

    # Check \section / \subsection / \subsubsection
    m_sec = re.match(r"^\\(section|subsection|subsubsection)\*?\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}(.*)$", s, re.S)
    if m_sec:
        sec_type = m_sec.group(1)
        sec_title = m_sec.group(2).strip()
        remainder = m_sec.group(3).strip()
        sz_map = {"section": 14.0, "subsection": 13.0, "subsubsection": 12.0}
        hp = doc.add_paragraph()
        hp.paragraph_format.space_before = Pt(10)
        hp.paragraph_format.space_after = Pt(4)
        hp.paragraph_format.keep_with_next = True
        append_rich_latex_text_to_paragraph(hp, sec_title, default_bold=True, font_size_pt=sz_map.get(sec_type, 13.0))
        if remainder:
            _render_text_block(doc, remainder, base_dir=base_dir)
        return

    # Check \caption{...}
    m_cap = re.match(r"^\\caption\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}(.*)$", s, re.S)
    if m_cap:
        cap_text = m_cap.group(1).strip()
        remainder = m_cap.group(2).strip()
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(2)
        cp.paragraph_format.space_after = Pt(8)
        r_lbl = cp.add_run("Hình / Bảng: ")
        r_lbl.bold = True
        r_lbl.italic = True
        r_lbl.font.size = Pt(10.5)
        append_rich_latex_text_to_paragraph(cp, cap_text, font_size_pt=10.5)
        if remainder:
            _render_text_block(doc, remainder, base_dir=base_dir)
        return

    # Check \includegraphics[...]{path}
    m_img = re.match(r"^\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}(.*)$", s, re.S)
    if m_img and base_dir:
        img_rel = m_img.group(1).strip()
        remainder = m_img.group(2).strip()
        img_path = (base_dir / img_rel).resolve()
        if img_path.exists():
            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ip.add_run().add_picture(str(img_path), width=Inches(3.6))
        if remainder:
            _render_text_block(doc, remainder, base_dir=base_dir)
        return

    # Standard paragraph
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    append_rich_latex_text_to_paragraph(p, s, font_size_pt=12.0)


def convert_latex_file_unified(
    tex: Path,
    output: Path,
    *,
    engine: str = "xelatex",
    dpi: int = 400,
    reference: Path | None = None,
    strict: bool = False,
    qa_pdf: bool = False,
    mode: int = 2,
) -> dict[str, Any]:
    """Unified entry point compatible with `latex2word.py`, `backend/server.py`, and `app.py`.
    Supports both `.tex` and `.pdf` input files.
    """
    tex = tex.resolve()
    output = output.resolve()
    if not tex.is_file():
        raise RuntimeError(f"Input does not exist: {tex}")
    output.parent.mkdir(parents=True, exist_ok=True)

    input_is_pdf = tex.suffix.lower() == ".pdf"
    if input_is_pdf:
        base_root = Path(__file__).resolve().parent.parent
        search_dirs = [tex.parent, tex.parent.parent, base_root, base_root / "jobs"]
        original = decompile_pdf_to_latex(tex, search_dirs=search_dirs)
        original = clean_newlines(original)
        included_count = 0
        synth_tex = output.with_suffix(".tex")
        if synth_tex != tex:
            synth_tex.write_text(original, encoding="utf-8")
    else:
        original, included_count = expand_tex_includes(tex)
        original = clean_newlines(original)

    build_stats = build_native_docx_from_latex(
        original,
        output,
        base_dir=tex.parent,
        mode=mode,
    )

    tikz_total = len(re.findall(r"\\begin\{tikzpicture\}", original))
    axis_only = len(re.findall(r"\\begin\{axis\}", original))
    standalone_axis = max(0, axis_only - tikz_total)
    source_graphics = len(re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}", original))
    expected_tables = len(re.findall(r"\\begin\{(?:tabular\*?|tabularx|longtable)\}", original))

    report: dict[str, Any] = {
        "input": str(tex),
        "input_type": "pdf" if input_is_pdf else "tex",
        "output": str(output),
        "mode": mode,
        "engine": "Native DrawingML + OMML + MathTables + PDF Decompiler Unified Engine",
        "rendered_figures": build_stats.get("rendered_figures", 0),
        "drawingml_groups": build_stats.get("drawingml_groups", 0),
        "drawingml_shapes": build_stats.get("drawingml_shapes", 0),
        "tkz_tab_tables": build_stats.get("tkz_tab_tables", 0),
        "choice_tables": build_stats.get("choice_tables", 0),
        "data_tables": build_stats.get("data_tables", 0),
        "immini_layouts": build_stats.get("immini_layouts", 0),
        "expanded_tex_includes": included_count,
        "warnings": build_stats.get("warnings", []),
        "word_tables": build_stats.get("word_tables", 0),
        "equation_nodes": build_stats.get("equation_nodes", 0),
        "drawing_references": build_stats.get("drawing_references", 0),
        "embedded_media_files": build_stats.get("embedded_media_files", 0),
        "input_table_environments": expected_tables,
        "input_includegraphics": source_graphics,
        "expected_total_image_commands": source_graphics + build_stats.get("rendered_figures", 0),
        "ooxml_stats": build_stats.get("ooxml_stats", {}),
    }

    if qa_pdf:
        render_dir = output.parent / (output.stem + "_qa")
        render_dir.mkdir(exist_ok=True)
        target_pdf = render_dir / (output.stem + ".pdf")
        ok = convert_docx_to_pdf(output, target_pdf)
        if ok and target_pdf.exists():
            report["qa_rendered_pdf"] = str(target_pdf)

    report["status"] = "review_required" if report["warnings"] else "automated_checks_passed"
    report_location = output.with_suffix(".report.json")
    report_location.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return report
