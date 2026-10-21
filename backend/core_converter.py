"""Unified Core LaTeX & PDF to Word Converter Engine (`backend/core_converter.py`).

Single unified conversion engine shared across:
1. CLI (`latex2word.py`)
2. FastAPI Web Server (`backend/server.py` / `run_server.py`)
3. Streamlit Web App (`app.py`)

Features:
- Accepts `.tex` AND `.pdf` inputs (with a 100% Pure-Python 2D PDF Math & Text Decompiler using `zlib` + `re`
  that works out-of-the-box on Windows/Linux/macOS without requiring `pypdf`, `pdfplumber`, or `pdftotext`)
- Reconstructs 2D positioned math expressions from TeX-compiled PDFs (`\\max_{[\\alpha;\\beta]}`, stacked fractions
  `\\dfrac{A}{B}`, 2-row bracket/brace systems `\\hoac{...}` / `\\heva{...}`, overset/underset arrows
  `\\xrightarrow[sub]{sup}`, superscripts/subscripts, tall absolute value bars, and multiple-choice options)
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
import zlib
from dataclasses import dataclass
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

_WIN_PATH_REFRESHED = False


def _refresh_windows_path_once() -> None:
    """Refresh PATH from Windows Registry so Streamlit/CLI/FastAPI always find installed Poppler/MiKTeX/LibreOffice."""
    global _WIN_PATH_REFRESHED
    if _WIN_PATH_REFRESHED or sys.platform != "win32":
        return
    _WIN_PATH_REFRESHED = True
    try:
        import winreg

        paths: list[str] = []
        for hkey in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            sub = (
                r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
                if hkey == winreg.HKEY_LOCAL_MACHINE
                else r"Environment"
            )
            try:
                with winreg.OpenKey(hkey, sub) as key:
                    p, _ = winreg.QueryValueEx(key, "Path")
                    paths.append(p)
            except Exception:
                pass
        if paths:
            combined = ";".join(paths) + ";" + os.environ.get("PATH", "")
            seen: set[str] = set()
            clean_paths: list[str] = []
            for part in combined.split(";"):
                part_clean = part.strip()
                if part_clean and part_clean.lower() not in seen:
                    seen.add(part_clean.lower())
                    clean_paths.append(part_clean)
            os.environ["PATH"] = ";".join(clean_paths)
    except Exception:
        pass


_refresh_windows_path_once()


# ==============================================================================
# PURE-PYTHON 2D PDF MATH & TEXT DECOMPILER (Zero External Dependencies)
# ==============================================================================
AGL_MAP: dict[str, str] = {
    "space": " ", "exclam": "!", "quotedbl": '"', "numbersign": "#", "dollar": "$",
    "percent": "%", "ampersand": "&", "quotesingle": "'", "parenleft": "(", "parenright": ")",
    "asterisk": "*", "plus": "+", "comma": ",", "hyphen": "-", "minus": "−", "period": ".",
    "slash": "/", "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
    "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9", "colon": ":",
    "semicolon": ";", "less": "<", "equal": "=", "greater": ">", "question": "?", "at": "@",
    "bracketleft": "[", "backslash": "\\", "bracketright": "]", "asciicircum": "^",
    "underscore": "_", "grave": "`", "braceleft": "{", "bar": "|", "braceright": "}",
    "asciitilde": "~", "bullet": "•", "endash": "–", "emdash": "—",
    "Alpha": "Α", "Beta": "Β", "Gamma": "Γ", "Delta": "Δ", "Epsilon": "Ε", "Zeta": "Ζ",
    "Eta": "Η", "Theta": "Θ", "Iota": "Ι", "Kappa": "Κ", "Lambda": "Λ", "Mu": "Μ",
    "Nu": "Ν", "Xi": "Ξ", "Omicron": "Ο", "Pi": "Π", "Rho": "Ρ", "Sigma": "Σ",
    "Tau": "Τ", "Upsilon": "Υ", "Phi": "Φ", "Chi": "Χ", "Psi": "Ψ", "Omega": "Ω",
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "varepsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "omicron": "ο", "pi": "π", "rho": "ρ",
    "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "ϕ", "varphi": "φ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "lessequal": "≤", "greaterequal": "≥", "notequal": "≠", "approxequal": "≈",
    "equivalence": "≡", "infinity": "∞", "partialdiff": "∂", "integral": "∫",
    "summation": "∑", "product": "∏", "radical": "√", "plusminus": "±", "minusplus": "∓",
    "multiply": "×", "divide": "÷", "periodcentered": "⋅", "dotmath": "⋅",
    "arrowright": "→", "arrowleft": "←", "arrowboth": "↔",
    "arrowdblright": "⇒", "arrowdblleft": "⇐", "arrowdblboth": "⇔",
    "universal": "∀", "existential": "∃", "emptyset": "∅",
    "element": "∈", "notelement": "∉", "propersubset": "⊂", "reflexsubset": "⊆",
    "propersuperset": "⊃", "reflexsuperset": "⊇", "intersection": "∩", "union": "∪",
    "logicaland": "∧", "logicalor": "∨", "logicalnot": "¬", "perp": "⊥", "similar": "∼",
    "prime": "′", "dblverticalbar": "‖",
    "Agrave": "À", "Aacute": "Á", "Acircumflex": "Â", "Atilde": "Ã", "Adieresis": "Ä",
    "Egrave": "È", "Eacute": "É", "Ecircumflex": "Ê", "Igrave": "Ì", "Iacute": "Í",
    "Ograve": "Ò", "Oacute": "Ó", "Ocircumflex": "Ô", "Otilde": "Õ",
    "Ugrave": "Ù", "Uacute": "Ú", "Yacute": "Ý",
    "agrave": "à", "aacute": "á", "acircumflex": "â", "atilde": "ã",
    "egrave": "è", "eacute": "é", "ecircumflex": "ê", "igrave": "ì", "iacute": "í",
    "ograve": "ò", "oacute": "ó", "ocircumflex": "ô", "otilde": "õ",
    "ugrave": "ù", "uacute": "ú", "yacute": "ý",
    "Abreve": "Ă", "abreve": "ă", "Dcroat": "Đ", "dcroat": "đ",
    "Ohorn": "Ơ", "ohorn": "ơ", "Uhorn": "Ư", "uhorn": "ư",
}

MATH_SYMBOL_LATEX: dict[str, str] = {
    "−": "-",
    "±": r"\pm ",
    "∓": r"\mp ",
    "×": r"\times ",
    "⋅": r"\cdot ",
    "·": r"\cdot ",
    "÷": r"\div ",
    "≤": r"\le ",
    "≥": r"\ge ",
    "≠": r"\neq ",
    "≈": r"\approx ",
    "≡": r"\equiv ",
    "∈": r"\in ",
    "∉": r"\notin ",
    "⊂": r"\subset ",
    "⊆": r"\subseteq ",
    "∪": r"\cup ",
    "∩": r"\cap ",
    "∅": r"\emptyset ",
    "∀": r"\forall ",
    "∃": r"\exists ",
    "∞": r"\infty ",
    "→": r"\to ",
    "⇒": r"\Rightarrow ",
    "⇔": r"\Leftrightarrow ",
    "∨": r"\lor ",
    "∧": r"\land ",
    "α": r"\alpha ",
    "β": r"\beta ",
    "γ": r"\gamma ",
    "δ": r"\delta ",
    "ε": r"\varepsilon ",
    "θ": r"\theta ",
    "λ": r"\lambda ",
    "μ": r"\mu ",
    "π": r"\pi ",
    "σ": r"\sigma ",
    "φ": r"\varphi ",
    "ω": r"\omega ",
    "Δ": r"\Delta ",
    "Ω": r"\Omega ",
    "√": r"\sqrt",
    "∫": r"\int ",
    "∑": r"\sum ",
    "∏": r"\prod ",
    "ℝ": r"\mathbb{R}",
    "ℕ": r"\mathbb{N}",
    "ℤ": r"\mathbb{Z}",
    "ℚ": r"\mathbb{Q}",
    "ℂ": r"\mathbb{C}",
}

UNICODE_TO_LATEX_MATH: list[tuple[str, str]] = list(MATH_SYMBOL_LATEX.items())


def _resolve_glyph_name(gname: str) -> str:
    if gname in AGL_MAP:
        return AGL_MAP[gname]
    if gname.startswith("uni") and len(gname) == 7:
        try:
            return chr(int(gname[3:], 16))
        except ValueError:
            pass
    if len(gname) == 1:
        return gname
    return ""


@dataclass
class PdfTextSpan:
    x: float
    y: float
    font_size: float
    font_name: str
    text: str


def _decode_pdf_literal_bytes(lit: str) -> list[int]:
    out: list[int] = []
    i = 0
    n = len(lit)
    while i < n:
        ch = lit[i]
        if ch == "\\":
            i += 1
            if i >= n:
                break
            esc = lit[i]
            if esc in "01234567":
                oct_s = esc
                i += 1
                for _ in range(2):
                    if i < n and lit[i] in "01234567":
                        oct_s += lit[i]
                        i += 1
                    else:
                        break
                out.append(int(oct_s, 8) & 0xFF)
                continue
            esc_map = {"n": 10, "r": 13, "t": 9, "b": 8, "f": 12, "(": 40, ")": 41, "\\": 92}
            if esc in esc_map:
                out.append(esc_map[esc])
            elif esc in "\r\n":
                if esc == "\r" and i + 1 < n and lit[i + 1] == "\n":
                    i += 1
            else:
                out.append(ord(esc) & 0xFF)
            i += 1
        else:
            out.append(ord(ch) & 0xFF)
            i += 1
    return out


def _parse_content_stream_spans(
    content: str,
    p_fonts: dict[str, tuple[str, dict[int, str]]],
) -> list[PdfTextSpan]:
    spans: list[PdfTextSpan] = []
    cur_font_key = ""
    cur_font_size = 12.0

    def decode_str_token(tok: str) -> str:
        _, fmap = p_fonts.get(cur_font_key, ("", {}))
        if tok.startswith("<") and tok.endswith(">"):
            hex_s = re.sub(r"\s+", "", tok[1:-1])
            if len(hex_s) % 2 == 1:
                hex_s += "0"
            b_list = [int(hex_s[i : i + 2], 16) for i in range(0, len(hex_s), 2)]
        else:
            b_list = _decode_pdf_literal_bytes(tok[1:-1])
        chars: list[str] = []
        for b in b_list:
            if b in fmap:
                chars.append(fmap[b])
            elif 32 <= b <= 126:
                chars.append(chr(b))
            else:
                chars.append(chr(b))
        return "".join(chars)

    op_re = re.compile(
        r"/([F]\d+)\s+([+-]?\d+(?:\.\d+)?)\s+Tf"
        r"|([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+Tm"
        r"|([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+(Td|TD)"
        r"|([+-]?\d+(?:\.\d+)?)\s+TL"
        r"|(T\*)"
        r"|\[(.*?)\]\s*TJ"
        r"|(\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]*>)\s*Tj",
        re.S,
    )

    tl = 14.446
    for bt_block in re.findall(r"\bBT\b(.*?)\bET\b", content, re.S):
        tm_a, tm_d = 1.0, 1.0
        tm_e, tm_f = 0.0, 0.0
        lm_e, lm_f = 0.0, 0.0

        for m in op_re.finditer(bt_block):
            if m.group(1):
                cur_font_key = m.group(1)
                cur_font_size = float(m.group(2))
            elif m.group(3):
                tm_a = float(m.group(3))
                tm_d = float(m.group(6))
                tm_e = float(m.group(7))
                tm_f = float(m.group(8))
                lm_e, lm_f = tm_e, tm_f
            elif m.group(9):
                tx = float(m.group(9))
                ty = float(m.group(10))
                op_td = m.group(11)
                if op_td == "TD":
                    tl = -ty
                lm_e += tx * tm_a
                lm_f += ty * tm_d
                tm_e, tm_f = lm_e, lm_f
            elif m.group(12):
                tl = float(m.group(12))
            elif m.group(13):
                lm_f -= tl * tm_d
                tm_e, tm_f = lm_e, lm_f
            elif m.group(14) is not None:
                arr_s = m.group(14)
                font_name, _ = p_fonts.get(cur_font_key, (cur_font_key, {}))
                eff_fs = abs(cur_font_size * tm_d) if abs(tm_d) > 0.01 else cur_font_size
                sub_parts: list[str] = []
                cur_x = tm_e
                span_start_x = cur_x
                for item in re.findall(r"\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]*>|[+-]?\d+(?:\.\d+)?", arr_s):
                    if item.startswith("(") or item.startswith("<"):
                        s_dec = decode_str_token(item)
                        if not sub_parts:
                            span_start_x = cur_x
                        sub_parts.append(s_dec)
                        cur_x += len(s_dec) * eff_fs * 0.48
                    else:
                        kern = float(item)
                        if kern < -180 and sub_parts:
                            txt_piece = "".join(sub_parts)
                            if txt_piece:
                                spans.append(PdfTextSpan(span_start_x, tm_f, eff_fs, font_name, txt_piece))
                            sub_parts = []
                        cur_x -= (kern / 1000.0) * eff_fs
                        if not sub_parts:
                            span_start_x = cur_x
                if sub_parts:
                    txt_piece = "".join(sub_parts)
                    if txt_piece:
                        spans.append(PdfTextSpan(span_start_x, tm_f, eff_fs, font_name, txt_piece))
                tm_e = cur_x
            elif m.group(15):
                tok = m.group(15)
                font_name, _ = p_fonts.get(cur_font_key, (cur_font_key, {}))
                eff_fs = abs(cur_font_size * tm_d) if abs(tm_d) > 0.01 else cur_font_size
                s_dec = decode_str_token(tok)
                if s_dec:
                    spans.append(PdfTextSpan(tm_e, tm_f, eff_fs, font_name, s_dec))
                tm_e += len(s_dec) * eff_fs * 0.48

    return spans


def extract_pdf_spans_pure_python(pdf_path: Path) -> list[list[PdfTextSpan]]:
    """Extract 2D positioned text & math spans from any PDF using only Python standard library (`zlib` + `re`)."""
    try:
        data = pdf_path.read_bytes()
    except Exception:
        return []

    obj_re = re.compile(rb"(\d+)\s+0\s+obj\b(.*?)endobj", re.S)
    objects: dict[int, bytes] = {}
    streams: dict[int, bytes] = {}

    for m in obj_re.finditer(data):
        oid = int(m.group(1))
        body = m.group(2)
        objects[oid] = body
        sm = re.search(rb"stream\r?\n(.*?)\r?\nendstream", body, re.S)
        if sm:
            raw_s = sm.group(1)
            if b"/FlateDecode" in body:
                try:
                    streams[oid] = zlib.decompressobj().decompress(raw_s)
                except Exception:
                    try:
                        streams[oid] = zlib.decompress(raw_s.strip(b"\r\n"))
                    except Exception:
                        streams[oid] = raw_s
            else:
                streams[oid] = raw_s

    cmaps: dict[int, dict[int, str]] = {}
    for oid, st in streams.items():
        if b"beginbfchar" in st or b"beginbfrange" in st:
            txt = st.decode("latin1", errors="ignore")
            cmap: dict[int, str] = {}
            for block in re.findall(r"beginbfchar(.*?)endbfchar", txt, re.S):
                for m_c in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
                    src_c = int(m_c.group(1), 16)
                    dst_hex = m_c.group(2)
                    chars = [
                        chr(int(dst_hex[i : i + 4], 16))
                        for i in range(0, len(dst_hex), 4)
                        if len(dst_hex[i : i + 4]) == 4
                    ]
                    cmap[src_c] = "".join(chars).replace("\x00", "")
            for block in re.findall(r"beginbfrange(.*?)endbfrange", txt, re.S):
                for m_r in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
                    c1, c2, d_start = int(m_r.group(1), 16), int(m_r.group(2), 16), int(m_r.group(3), 16)
                    for offset in range(max(0, min(256, c2 - c1 + 1))):
                        cmap[c1 + offset] = chr(d_start + offset)
                for m_arr in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*?)\]", block, re.S):
                    c1 = int(m_arr.group(1), 16)
                    hex_list = re.findall(r"<([0-9A-Fa-f]+)>", m_arr.group(3))
                    for offset, h_str in enumerate(hex_list):
                        chars = [
                            chr(int(h_str[i : i + 4], 16))
                            for i in range(0, len(h_str), 4)
                            if len(h_str[i : i + 4]) == 4
                        ]
                        cmap[c1 + offset] = "".join(chars).replace("\x00", "")
            if cmap:
                cmaps[oid] = cmap

    enc_diffs: dict[int, dict[int, str]] = {}
    for oid, body in objects.items():
        if b"/Differences" in body:
            txt = body.decode("latin1", errors="ignore")
            m_d = re.search(r"/Differences\s*\[(.*?)\]", txt, re.S)
            if m_d:
                dmap: dict[int, str] = {}
                cur_code = 0
                for tok in re.findall(r"(\d+|/[A-Za-z0-9_.]+)", m_d.group(1)):
                    if tok.isdigit():
                        cur_code = int(tok)
                    elif tok.startswith("/"):
                        g_str = _resolve_glyph_name(tok[1:])
                        if g_str:
                            dmap[cur_code] = g_str
                        cur_code += 1
                if dmap:
                    enc_diffs[oid] = dmap

    font_info: dict[int, tuple[str, dict[int, str]]] = {}
    for oid, body in objects.items():
        if b"/Type" in body and b"/Font" in body:
            txt = body.decode("latin1", errors="ignore")
            m_bf = re.search(r"/BaseFont\s*/([A-Za-z0-9+_\-]+)", txt)
            base_font = m_bf.group(1) if m_bf else f"Font_{oid}"
            fmap: dict[int, str] = {}
            m_enc = re.search(r"/Encoding\s+(\d+)\s+0\s+R", txt)
            if m_enc and int(m_enc.group(1)) in enc_diffs:
                fmap.update(enc_diffs[int(m_enc.group(1))])
            m_tu = re.search(r"/ToUnicode\s+(\d+)\s+0\s+R", txt)
            if m_tu and int(m_tu.group(1)) in cmaps:
                fmap.update(cmaps[int(m_tu.group(1))])
            font_info[oid] = (base_font, fmap)

    def resolve_dict_text(raw_b: bytes) -> str:
        t = raw_b.decode("latin1", errors="ignore")
        for m_ref in re.finditer(r"/Resources\s+(\d+)\s+0\s+R", t):
            rid = int(m_ref.group(1))
            if rid in objects:
                t += "\n" + objects[rid].decode("latin1", errors="ignore")
        for m_fref in re.finditer(r"/Font\s+(\d+)\s+0\s+R", t):
            fid = int(m_fref.group(1))
            if fid in objects:
                t += "\n" + objects[fid].decode("latin1", errors="ignore")
        return t

    global_fonts: dict[str, tuple[str, dict[int, str]]] = {}
    for oid, body in objects.items():
        txt = body.decode("latin1", errors="ignore")
        for m_f in re.finditer(r"/(F\d+)\s+(\d+)\s+0\s+R", txt):
            fn, f_oid = m_f.group(1), int(m_f.group(2))
            if f_oid in font_info:
                global_fonts[fn] = font_info[f_oid]

    page_entries: list[tuple[int, list[int], dict[str, tuple[str, dict[int, str]]]]] = []
    for oid in sorted(objects.keys()):
        body = objects[oid]
        if re.search(rb"/Type\s*/Page\b", body):
            txt = resolve_dict_text(body)
            p_fonts = dict(global_fonts)
            for m_f in re.finditer(r"/(F\d+)\s+(\d+)\s+0\s+R", txt):
                fn, f_oid = m_f.group(1), int(m_f.group(2))
                if f_oid in font_info:
                    p_fonts[fn] = font_info[f_oid]
            c_oids: list[int] = []
            m_c_arr = re.search(r"/Contents\s*\[(.*?)\]", txt, re.S)
            if m_c_arr:
                c_oids = [int(x) for x in re.findall(r"(\d+)\s+0\s+R", m_c_arr.group(1))]
            else:
                m_c_single = re.search(r"/Contents\s+(\d+)\s+0\s+R", txt)
                if m_c_single:
                    c_oids = [int(m_c_single.group(1))]
            if c_oids:
                page_entries.append((oid, c_oids, p_fonts))

    pages_spans: list[list[PdfTextSpan]] = []
    for _, c_oids, p_fonts in page_entries:
        raw_stream = b"\n".join(streams.get(cid, b"") for cid in c_oids)
        content = raw_stream.decode("latin1", errors="ignore")
        spans = _parse_content_stream_spans(content, p_fonts)
        if spans:
            pages_spans.append(spans)
    return pages_spans


def _is_math_font(font_name: str, text: str) -> bool:
    fn = font_name
    t = text.strip()
    if not t:
        return False
    if "CMEX" in fn or "CMSY" in fn or "CMMI" in fn or "CMR" in fn or "MSBM" in fn or "StandardSym" in fn:
        return True
    if "NimbusRomNo9L-ReguItal" in fn:
        return True
    if any(ch in t for ch in "−±∓×⋅·÷≤≥≠≈≡∈∉⊂⊆∪∩∅∀∃∞→⇒⇔∨∧αβγδεθλμπσφωΔΩ|{}[]"):
        return True
    if "VnNimbus" in fn:
        if "Bold" in fn:
            return False
        if t in ("max", "min", "Max", "Min") or re.fullmatch(r"[\d.,;+\-*/=()\[\]|]+", t):
            return True
        return False
    if "NimbusRomNo9L-Regu" in fn:
        return True
    return False


def _span_to_math_latex(sp: PdfTextSpan) -> str:
    t = sp.text
    if "CMEX" in sp.font_name:
        if "\x0c" in t or "|" in t:
            return "|"
        return ""
    if "MSBM" in sp.font_name and t.strip() in ("R", "N", "Z", "Q", "C"):
        return f"\\mathbb{{{t.strip()}}}"
    out: list[str] = []
    for ch in t:
        if ch == "{":
            out.append(r"\{")
        elif ch == "}":
            out.append(r"\}")
        elif ch in MATH_SYMBOL_LATEX:
            out.append(MATH_SYMBOL_LATEX[ch])
        else:
            out.append(ch)
    s = "".join(out)
    s = re.sub(r"\bmax\b", r"\\max", s)
    s = re.sub(r"\bmin\b", r"\\min", s)
    s = re.sub(r"\bMax\b", r"\\text{Max}", s)
    s = re.sub(r"\bMin\b", r"\\text{Min}", s)
    return s


def _render_simple_math_seq(spans: list[PdfTextSpan], base_y: float) -> str:
    parts: list[str] = []
    i = 0
    n = len(spans)
    while i < n:
        sp = spans[i]
        dy = sp.y - base_y
        if sp.font_size < 10.5 and 2.2 <= dy <= 6.5:
            sup_s = [sp]
            i += 1
            while i < n and spans[i].font_size < 10.5 and abs(spans[i].y - sp.y) <= 1.8:
                sup_s.append(spans[i])
                i += 1
            sup_ltx = "".join(_span_to_math_latex(x) for x in sup_s).strip()
            parts.append(f"^{{{sup_ltx}}}")
            continue
        if sp.font_size < 10.5 and -5.0 <= dy <= -1.2:
            sub_s = [sp]
            i += 1
            while i < n and spans[i].font_size < 10.5 and abs(spans[i].y - sp.y) <= 1.8:
                sub_s.append(spans[i])
                i += 1
            sub_ltx = "".join(_span_to_math_latex(x) for x in sub_s).strip()
            parts.append(f"_{{{sub_ltx}}}")
            continue
        if "VnNimbus" in sp.font_name and not _is_math_font(sp.font_name, sp.text):
            parts.append(f"\\text{{ {sp.text.strip()} }}")
        else:
            parts.append(_span_to_math_latex(sp))
        i += 1
    res = "".join(parts)
    while re.search(r"\\text\{\s*([^{}]*?)\s*\}\s*\\text\{\s*([^{}]*?)\s*\}", res):
        res = re.sub(r"\\text\{\s*([^{}]*?)\s*\}\s*\\text\{\s*([^{}]*?)\s*\}", r"\\text{ \1 \2 }", res)
    res = re.sub(r"\s+", " ", res).strip()
    return res


def reconstruct_page_lines(spans: list[PdfTextSpan]) -> list[str]:
    """Reconstruct 2D positioned PDF spans into structured LaTeX paragraphs, math blocks, and choice tables."""
    raw_spans = [
        sp for sp in spans
        if sp.y <= 785.0 and not (sp.y < 75.0 and sp.text.strip().isdigit())
    ]
    clean_spans: list[PdfTextSpan] = []
    for sp in raw_spans:
        if "CMEX" in sp.font_name and "\x0c" in sp.text:
            if clean_spans and clean_spans[-1].text == "|" and abs(sp.x - clean_spans[-1].x) < 1.5:
                continue
            ref_y = clean_spans[-1].y if clean_spans else sp.y
            clean_spans.append(PdfTextSpan(sp.x, ref_y, sp.font_size, "CMSY10", "|"))
        else:
            clean_spans.append(sp)
    if not clean_spans:
        return []

    raw_lines: list[str] = []
    cur_tokens: list[tuple[str, str]] = []

    def flush_line() -> None:
        nonlocal cur_tokens
        if not cur_tokens:
            return
        merged: list[tuple[str, str]] = []
        for kind, val in cur_tokens:
            if not val:
                continue
            if merged and merged[-1][0] == kind:
                sep = (
                    " "
                    if kind in ("text", "bold", "italic")
                    and not merged[-1][1].endswith(" ")
                    and not val.startswith((" ", ".", ",", ":", ";", ")", "?"))
                    else ""
                )
                merged[-1] = (kind, merged[-1][1] + sep + val)
            else:
                merged.append((kind, val))
        cur_tokens = []

        out_parts: list[str] = []
        for idx, (kind, val) in enumerate(merged):
            v = val.strip()
            if not v:
                continue
            if kind == "math":
                trail = ""
                if v.endswith(".") and not v.endswith(r"\right."):
                    v = v[:-1].strip()
                    trail = "."
                elif v.endswith(";") and idx == len(merged) - 1:
                    v = v[:-1].strip()
                    trail = ";"
                if v:
                    out_parts.append(f"${v}${trail}")
                elif trail:
                    out_parts.append(trail)
            elif kind == "bold":
                out_parts.append(f"\\textbf{{{v}}}")
            elif kind == "italic":
                out_parts.append(f"\\textit{{{v}}}")
            else:
                out_parts.append(v)

        line_str = " ".join(out_parts)
        line_str = re.sub(r"\s+([.,:;?!])", r"\1", line_str)
        if line_str.strip() and line_str.strip() != "@@@":
            raw_lines.append(line_str.strip())

    i = 0
    n = len(clean_spans)
    base_y = clean_spans[0].y

    while i < n:
        sp = clean_spans[i]
        t_strip = sp.text.strip()
        if not t_strip:
            i += 1
            continue

        # 1. Detect 2-row system delimiter in CMEX10 ('\x14' = '[' or '(' = '{')
        if "CMEX" in sp.font_name and any(c in sp.text for c in ("\x14", "(")):
            delim_cmd = r"\hoac" if "\x14" in sp.text else r"\heva"
            i += 1
            if i < n:
                y_r1 = clean_spans[i].y
                r1_spans: list[PdfTextSpan] = []
                while i < n and abs(clean_spans[i].y - y_r1) <= 4.8:
                    r1_spans.append(clean_spans[i])
                    i += 1
                r2_spans: list[PdfTextSpan] = []
                if i < n and -20.0 <= (clean_spans[i].y - y_r1) <= -9.5:
                    y_r2 = clean_spans[i].y
                    while i < n and abs(clean_spans[i].y - y_r2) <= 4.8:
                        if clean_spans[i].x < 88.0 and clean_spans[i].y < y_r2 - 2.0:
                            break
                        r2_spans.append(clean_spans[i])
                        i += 1
                r1_ltx = _render_simple_math_seq(r1_spans, y_r1)
                r2_ltx = _render_simple_math_seq(r2_spans, r2_spans[0].y) if r2_spans else ""
                sys_ltx = f"{delim_cmd}{{{r1_ltx} \\\\ {r2_ltx}}}" if r2_ltx else f"{delim_cmd}{{{r1_ltx}}}"
                cur_tokens.append(("math", sys_ltx))
                continue

        # 2. Detect new line when Y drops significantly
        dy = sp.y - base_y
        if cur_tokens and dy < -11.5:
            flush_line()
            base_y = sp.y
            dy = 0.0
        elif not cur_tokens:
            base_y = sp.y
            dy = 0.0

        # 3. Check for stacked fraction or overset/underset arrow when Y jumps UP by +6.0 .. +11.5
        if 6.0 <= dy <= 11.5:
            y_top = sp.y
            top_spans: list[PdfTextSpan] = []
            j = i
            while j < n and abs(clean_spans[j].y - y_top) <= 3.5:
                top_spans.append(clean_spans[j])
                j += 1

            # Case A: Overset/Underset Arrow: next span at base_y is '−−−→'
            if (
                j < n
                and abs(clean_spans[j].y - base_y) <= 3.0
                and "→" in clean_spans[j].text
                and "−" in clean_spans[j].text
            ):
                j += 1
                bot_spans: list[PdfTextSpan] = []
                if j < n and -13.0 <= (clean_spans[j].y - base_y) <= -6.0:
                    y_bot = clean_spans[j].y
                    while j < n and abs(clean_spans[j].y - y_bot) <= 3.5:
                        bot_spans.append(clean_spans[j])
                        j += 1
                top_ltx = _render_simple_math_seq(top_spans, y_top)
                bot_ltx = _render_simple_math_seq(bot_spans, bot_spans[0].y) if bot_spans else ""
                if bot_ltx:
                    cur_tokens.append(("math", f"\\xrightarrow[{bot_ltx}]{{{top_ltx}}}"))
                else:
                    cur_tokens.append(("math", f"\\xrightarrow{{{top_ltx}}}"))
                i = j
                continue

            # Case B: Stacked Fraction: next span is at y_den ≈ base_y - 8.2 (or y_top - 16.3)
            if j < n and -19.5 <= (clean_spans[j].y - y_top) <= -12.5:
                y_bot = clean_spans[j].y
                bot_spans = []
                while j < n and abs(clean_spans[j].y - y_bot) <= 3.5:
                    bot_spans.append(clean_spans[j])
                    j += 1
                num_ltx = _render_simple_math_seq(top_spans, y_top)
                den_ltx = _render_simple_math_seq(bot_spans, y_bot)
                cur_tokens.append(("math", f"\\dfrac{{{num_ltx}}}{{{den_ltx}}}"))
                i = j
                continue

        # 4. Check for max / min with underscript at y ≈ base_y - 8.8
        if (
            t_strip in ("max", "min")
            and i + 1 < n
            and -11.5 <= (clean_spans[i + 1].y - sp.y) <= -5.5
            and clean_spans[i + 1].font_size < 10.5
        ):
            op_cmd = f"\\{t_strip}"
            y_sub = clean_spans[i + 1].y
            j = i + 1
            sub_spans: list[PdfTextSpan] = []
            while j < n and abs(clean_spans[j].y - y_sub) <= 2.5 and clean_spans[j].font_size < 10.5:
                sub_spans.append(clean_spans[j])
                j += 1
            sub_ltx = _render_simple_math_seq(sub_spans, y_sub)
            cur_tokens.append(("math", f"{op_cmd}_{{{sub_ltx}}} "))
            i = j
            continue

        # 5. Check for superscripts & subscripts on current math token
        if sp.font_size < 10.5 and 2.2 <= dy <= 6.2 and cur_tokens:
            y_sup = sp.y
            sup_spans: list[PdfTextSpan] = []
            while i < n and clean_spans[i].font_size < 10.5 and abs(clean_spans[i].y - y_sup) <= 1.8:
                sup_spans.append(clean_spans[i])
                i += 1
            sup_ltx = _render_simple_math_seq(sup_spans, y_sup)
            cur_tokens.append(("math", f"^{{{sup_ltx}}}"))
            continue

        if sp.font_size < 10.5 and -5.0 <= dy <= -1.2 and cur_tokens:
            y_sub = sp.y
            sub_spans = []
            while i < n and clean_spans[i].font_size < 10.5 and abs(clean_spans[i].y - y_sub) <= 1.8:
                sub_spans.append(clean_spans[i])
                i += 1
            sub_ltx = _render_simple_math_seq(sub_spans, y_sub)
            cur_tokens.append(("math", f"_{{{sub_ltx}}}"))
            continue

        # Update base_y smoothly if near baseline
        if abs(dy) <= 2.2 and sp.font_size >= 10.5:
            base_y = sp.y

        # Classify span as bold, italic, text, or math
        if "VnNimbus" in sp.font_name and not _is_math_font(sp.font_name, sp.text):
            if "Bold" in sp.font_name:
                cur_tokens.append(("bold", t_strip))
            elif "Italic" in sp.font_name:
                cur_tokens.append(("italic", t_strip))
            else:
                cur_tokens.append(("text", t_strip))
        else:
            m_ltx = _span_to_math_latex(sp)
            if m_ltx:
                cur_tokens.append(("math", m_ltx))

        i += 1

    flush_line()

    # Merge wrapped continuation lines while preserving distinct logical paragraphs
    merged_lines: list[str] = []
    new_para_re = re.compile(
        r"^(?:\\textbf\{(?:Dạng|Phương pháp|Cách|TH\d|BT\d|Câu|Bài|Ví dụ|[A-D]\.?)|[A-D]\.\s|Vậy\b|Kiểm tra\b|Để\b|Trước tiên\b)"
    )
    for ln in raw_lines:
        if not merged_lines:
            merged_lines.append(ln)
            continue
        if new_para_re.match(ln):
            merged_lines.append(ln)
        else:
            prev = merged_lines[-1]
            if prev.endswith("$") and ln.startswith("$"):
                merged_lines[-1] = prev[:-1] + " " + ln[1:]
            elif prev.endswith("}") and ln.startswith(r"\textbf{"):
                merged_lines[-1] = prev[:-1] + " " + ln[len(r"\textbf{") :]
            else:
                merged_lines[-1] = prev + " " + ln

    # Convert consecutive A., B., C., D. option lines into \choice{...}{...}{...}{...}
    final_lines: list[str] = []
    idx = 0
    opt_re = re.compile(r"^(?:\\textbf\{([A-D])\}\.?|\\textbf\{([A-D])\.\}|([A-D])\.)\s*(.*)$")
    while idx < len(merged_lines):
        if idx + 3 < len(merged_lines):
            m_a = opt_re.match(merged_lines[idx])
            m_b = opt_re.match(merged_lines[idx + 1])
            m_c = opt_re.match(merged_lines[idx + 2])
            m_d = opt_re.match(merged_lines[idx + 3])
            if (
                m_a
                and (m_a.group(1) or m_a.group(2) or m_a.group(3)) == "A"
                and m_b
                and (m_b.group(1) or m_b.group(2) or m_b.group(3)) == "B"
                and m_c
                and (m_c.group(1) or m_c.group(2) or m_c.group(3)) == "C"
                and m_d
                and (m_d.group(1) or m_d.group(2) or m_d.group(3)) == "D"
            ):
                final_lines.append(
                    f"\\choice{{{m_a.group(4).strip()}}}{{{m_b.group(4).strip()}}}{{{m_c.group(4).strip()}}}{{{m_d.group(4).strip()}}}"
                )
                idx += 4
                continue
        final_lines.append(merged_lines[idx])
        idx += 1

    return final_lines


def _extract_pdf_text_pages(pdf_path: Path) -> list[str]:
    """Extract text page by page using Pure-Python 2D Span Extractor, pdftotext, pypdf, or pdfplumber."""
    _refresh_windows_path_once()
    pages_spans = extract_pdf_spans_pure_python(pdf_path)
    if pages_spans:
        pages_text = ["\n".join(sp.text for sp in p_spans if sp.text.strip()) for p_spans in pages_spans]
        if any(p.strip() for p in pages_text):
            return pages_text

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

    return pages


def _normalize_fingerprint_words(text: str) -> set[str]:
    s = re.sub(r"\\[a-zA-Z@]+\*?", " ", text)
    s = re.sub(r"[^0-9a-zA-ZÀ-ỹ]+", " ", s.lower())
    return {w for w in s.split() if len(w) >= 3 and not w.isdigit()}


def _find_matching_tex_source(pdf_path: Path, pdf_text: str, search_dirs: list[Path]) -> Path | None:
    same_stem = pdf_path.with_suffix(".tex")
    if same_stem.is_file() and same_stem.resolve() != pdf_path.resolve():
        # Verify the same_stem .tex was not merely a previously generated output with a different content
        try:
            stem_tex_words = _normalize_fingerprint_words(same_stem.read_text(encoding="utf-8", errors="replace"))
            pdf_words_check = _normalize_fingerprint_words(pdf_text)
            if stem_tex_words and pdf_words_check:
                jacc = len(stem_tex_words & pdf_words_check) / float(max(1, len(stem_tex_words | pdf_words_check)))
                if jacc >= 0.72:
                    return same_stem
        except Exception:
            pass

    pdf_words = _normalize_fingerprint_words(pdf_text)
    if len(pdf_words) < 25:
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
            inter = len(pdf_words & tex_words)
            union = len(pdf_words | tex_words)
            jaccard = inter / float(max(1, union))
            pdf_cov = inter / float(max(1, len(pdf_words)))
            tex_cov = inter / float(max(1, len(tex_words)))
            if pdf_cov >= 0.82 and tex_cov >= 0.72 and jaccard > best_score:
                best_score = jaccard
                best_match = resolved

    if best_match is not None and best_score >= 0.70:
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
    _refresh_windows_path_once()
    pdf_path = pdf_path.resolve()

    # 1. Extract 2D positioned spans using our Pure-Python PDF parser (zero external dependencies)
    pages_spans = extract_pdf_spans_pure_python(pdf_path)
    raw_pdf_text = "\n".join(
        " ".join(sp.text for sp in p_spans if sp.text.strip()) for p_spans in pages_spans
    )

    # 2. Check if an exact matching .tex source file exists in the workspace
    dirs = list(search_dirs or [])
    dirs.extend([pdf_path.parent, pdf_path.parent.parent])
    if raw_pdf_text.strip():
        matched_tex = _find_matching_tex_source(pdf_path, raw_pdf_text, dirs)
        if matched_tex is not None and matched_tex.is_file():
            return matched_tex.read_text(encoding="utf-8", errors="replace")

    # 3. Reconstruct 2D math & text layout directly from the Pure-Python 2D spans
    if pages_spans:
        out_blocks: list[str] = [
            r"\documentclass[12pt,a4paper]{article}",
            r"\usepackage{amsmath,amssymb}",
            r"\begin{document}",
            "",
        ]
        reconstructed_count = 0
        for p_spans in pages_spans:
            p_lines = reconstruct_page_lines(p_spans)
            for ln in p_lines:
                out_blocks.append(ln)
                out_blocks.append("")
                reconstructed_count += 1
        out_blocks.append(r"\end{document}")
        if reconstructed_count > 0:
            return "\n".join(out_blocks)

    # 4. Fallback to plain-text extractors (pdftotext / pypdf / pdfplumber)
    pages = _extract_pdf_text_pages(pdf_path)
    raw_pdf_text = "\n".join(pages)
    if raw_pdf_text.strip():
        matched_tex = _find_matching_tex_source(pdf_path, raw_pdf_text, dirs)
        if matched_tex is not None and matched_tex.is_file():
            return matched_tex.read_text(encoding="utf-8", errors="replace")
        return _convert_extracted_pdf_text_to_latex(pages)

    raise RuntimeError(
        f"Không thể trích xuất văn bản từ file PDF '{pdf_path.name}' (PDF dạng ảnh quét chưa có lớp text OCR)."
    )


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
    import tempfile

    _refresh_windows_path_once()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if lo:
        try:
            with tempfile.TemporaryDirectory(prefix="l2w_pdf_") as tmp_d:
                tmp_dir = Path(tmp_d)
                safe_docx = tmp_dir / "render_input.docx"
                shutil.copy2(str(docx_path), str(safe_docx))
                subprocess.run(
                    [lo, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(safe_docx)],
                    capture_output=True,
                    timeout=90,
                )
                expected = tmp_dir / "render_input.pdf"
                if expected.exists():
                    shutil.copy2(str(expected), str(pdf_path))
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
        "engine": "Native DrawingML + OMML + MathTables + Pure-Python 2D PDF Decompiler Engine",
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
