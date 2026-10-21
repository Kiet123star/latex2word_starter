r"""Unified Core PDF & LaTeX-to-Word Converter Engine (`backend/core_converter.py` v2.4 Universal).

Single unified conversion engine shared across:
1. CLI (`latex2word.py`)
2. FastAPI Web Server (`backend/server.py` / `run_server.py`)
3. Streamlit Web App (`app.py`)

Features:
- Universal 2D Pure-Python Math PDF Decompiler (`decompile_pdf_to_latex`) with zero external tool dependencies:
  * Unpacks PDF 1.5+ Compressed Object Streams (`/Type /ObjStm` with `/N` and `/First`)
  * Walks the `/Catalog -> /Pages -> /Kids` page tree in exact document order with inherited `/Resources`
  * 4-Layer Font Decoder:
    1) `/ToUnicode` CMap (`beginbfchar` / `beginbfrange`, stripping Variation Selectors `\uFE00-\uFE0F`)
    2) `/Encoding /Differences` array
    3) Type 1 `/FontFile` cleartext `dup <code> /<glyph_name> put` parser
    4) Built-in TeX math font tables (`CMSY*`, `CMEX*`, `CMMI*`, `CMR*`, `MSBM*`) for slots 0..127
  * Reconstructs 2D spatial fractions (`\dfrac{A}{B}`), nested radicals (`\sqrt{...}`), extensible systems
    of equations (`\heva{...}`, `\hoac{...}`), arrows with labels (`\xrightarrow[sub]{sup}`), superscripts,
    subscripts, multiple-choice tables (`\choice{...}{...}{...}{...}`), and Vietnamese typography (`VnNimbus*`,
    `VNR*`, `VNBX*`, `VNTI*`, `VNSL*`, `PTM*`, etc.).
- Native Word Shapes / DrawingML (<wpg:wgp>, <wps:wsp>, <a:custGeom>) for TikZ & PGFPlots
- Native Editable Math Tables (`tkz-tab`, `\choice`, `\choiceTF`, `tabular`, `\immini`)
- Native Office Math (OMML `<m:oMath>`) with XML 1.0 control-character sanitization
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import zlib
import zipfile
from dataclasses import dataclass, field
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
from .omml_math import (
    append_rich_latex_text_to_paragraph,
    latex_to_omml_element,
    sanitize_xml_text,
)
from .postprocessor import postprocess_docx_tables

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
}

INCLUDE_RE = re.compile(r"\\(?:input|include)\s*\{([^{}]+)\}")


def clean_newlines(s: str) -> str:
    """Normalize line endings to standard LF and strip illegal XML control chars."""
    s = s.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return sanitize_xml_text(s)


# ==============================================================================
# UNIVERSAL 2D PURE-PYTHON MATH PDF DECOMPILER (NO EXTERNAL DEPENDENCIES)
# ==============================================================================
GLYPH_TO_TEXT: dict[str, str] = {
    "space": " ",
    "exclam": "!",
    "quotedbl": '"',
    "numbersign": "#",
    "dollar": "$",
    "percent": "%",
    "ampersand": "&",
    "quotesingle": "'",
    "quoteright": "'",
    "quoteleft": "`",
    "parenleft": "(",
    "parenright": ")",
    "asterisk": "*",
    "plus": "+",
    "comma": ",",
    "hyphen": "-",
    "minus": "-",
    "period": ".",
    "slash": "/",
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
    "colon": ":",
    "semicolon": ";",
    "less": "<",
    "equal": "=",
    "greater": ">",
    "question": "?",
    "at": "@",
    "bracketleft": "[",
    "backslash": "\\",
    "bracketright": "]",
    "asciicircum": "^",
    "underscore": "_",
    "braceleft": "{",
    "bar": "|",
    "bracketlefttp": "[",
    "bracketleftbt": "[",
    "bracketleftex": "[",
    "bracketrighttp": "]",
    "bracketrightbt": "]",
    "bracketrightex": "]",
    "bracelefttp": "{",
    "braceleftmid": "{",
    "braceleftbt": "{",
    "braceex": "|",
    "braceright": "}",
    "asciitilde": "~",
    "bullet": "•",
    "endash": "–",
    "emdash": "—",
    "quotedblleft": "“",
    "quotedblright": "”",
    "ellipsis": "…",
    # Math & Greek glyphs
    "alpha": r"\alpha ",
    "beta": r"\beta ",
    "gamma": r"\gamma ",
    "delta": r"\delta ",
    "epsilon": r"\epsilon ",
    "epsilon1": r"\varepsilon ",
    "varepsilon": r"\varepsilon ",
    "zeta": r"\zeta ",
    "eta": r"\eta ",
    "theta": r"\theta ",
    "vartheta": r"\vartheta ",
    "iota": r"\iota ",
    "kappa": r"\kappa ",
    "lambda": r"\lambda ",
    "mu": r"\mu ",
    "nu": r"\nu ",
    "xi": r"\xi ",
    "pi": r"\pi ",
    "rho": r"\rho ",
    "sigma": r"\sigma ",
    "tau": r"\tau ",
    "upsilon": r"\upsilon ",
    "phi": r"\phi ",
    "phi1": r"\varphi ",
    "varphi": r"\varphi ",
    "chi": r"\chi ",
    "psi": r"\psi ",
    "omega": r"\omega ",
    "Gamma": r"\Gamma ",
    "Delta": r"\Delta ",
    "Theta": r"\Theta ",
    "Lambda": r"\Lambda ",
    "Xi": r"\Xi ",
    "Pi": r"\Pi ",
    "Sigma": r"\Sigma ",
    "Upsilon": r"\Upsilon ",
    "Phi": r"\Phi ",
    "Psi": r"\Psi ",
    "Omega": r"\Omega ",
    "infinity": r"\infty ",
    "inities": r"\infty ",
    "lessequal": r"\le ",
    "greaterequal": r"\ge ",
    "notequal": r"\neq ",
    "approxequal": r"\approx ",
    "equivalence": r"\equiv ",
    "similar": r"\sim ",
    "similarequal": r"\simeq ",
    "plusminus": r"\pm ",
    "minusplus": r"\mp ",
    "multiply": r"\times ",
    "divide": r"\div ",
    "periodcentered": r"\cdot ",
    " mathasterisk": "*",
    "arrowright": r"\to ",
    "arrowleft": r"\leftarrow ",
    "arrowboth": r"\leftrightarrow ",
    "arrowdblright": r"\Rightarrow ",
    "arrowdblleft": r"\Leftarrow ",
    "arrowdblboth": r"\Leftrightarrow ",
    "mapsto": r"\mapsto ",
    "arrowup": r"\uparrow ",
    "arrowdown": r"\downarrow ",
    "element": r"\in ",
    "notelement": r"\notin ",
    "owner": r"\ni ",
    "propersubset": r"\subset ",
    "reflexsubset": r"\subseteq ",
    "propersuperset": r"\supset ",
    "reflexsuperset": r"\supseteq ",
    "intersection": r"\cap ",
    "union": r"\cup ",
    "universal": r"\forall ",
    "existential": r"\exists ",
    "emptyset": r"\emptyset ",
    "logicaland": r"\land ",
    "logicalor": r"\lor ",
    "logicalnot": r"\neg ",
    "perp": r"\perp ",
    "perpendicular": r"\perp ",
    "parallel": r"\parallel ",
    "bardbl": r"\Vert ",
    "angle": r"\angle ",
    "triangle": r"\triangle ",
    "openbullet": r"\circ ",
    "circle": r"\circ ",
    "prime": "'",
    "partialdiff": r"\partial ",
    " gradient": r"\nabla ",
    "nabla": r"\nabla ",
    "radical": "@@RAD@@",
    "radicalbig": "@@RAD@@",
    "radicalBig": "@@RAD@@",
    "radicalbigg": "@@RAD@@",
    "radicalBigg": "@@RAD@@",
    "radicalvertex": "@@RAD_V@@",
    "radicalbt": "@@RAD_B@@",
    "parenleftbig": "(",
    "parenrightbig": ")",
    "parenleftBig": "(",
    "parenrightBig": ")",
    "parenleftbigg": "(",
    "parenrightbigg": ")",
    "parenleftBigg": "(",
    "parenrightBigg": ")",
    "parenlefttp": "(",
    "parenleftbt": "(",
    "parenleftex": "(",
    "parenrighttp": ")",
    "parenrightbt": ")",
    "parenrightex": ")",
    "bracketleftbig": "[",
    "bracketrightbig": "]",
    "bracketleftBig": "[",
    "bracketrightBig": "]",
    "bracketleftbigg": "[",
    "bracketrightbigg": "]",
    "bracketleftBigg": "[",
    "bracketrightBigg": "]",
    "braceleftbig": "{",
    "bracerightbig": "}",
    "braceleftBig": "{",
    "bracerightBig": "}",
    "braceleftbigg": "{",
    "bracerightbigg": "}",
    "braceleftBigg": "{",
    "bracerightBigg": "}",
    "summationtext": r"\sum ",
    "summationdisplay": r"\sum ",
    "producttext": r"\prod ",
    "productdisplay": r"\prod ",
    "integraltext": r"\int ",
    "integraldisplay": r"\int ",
    "uniontext": r"\bigcup ",
    "uniondisplay": r"\bigcup ",
    "intersectiontext": r"\bigcap ",
    "intersectiondisplay": r"\bigcap ",
    # Vietnamese Adobe / T5 / VnTeX glyph names
    "Agrave": "À",
    "Aacute": "Á",
    "Acircumflex": "Â",
    "Atilde": "Ã",
    "Adieresis": "Ä",
    "Egrave": "È",
    "Eacute": "É",
    "Ecircumflex": "Ê",
    "Igrave": "Ì",
    "Iacute": "Í",
    "Ograve": "Ò",
    "Oacute": "Ó",
    "Ocircumflex": "Ô",
    "Otilde": "Õ",
    "Ugrave": "Ù",
    "Uacute": "Ú",
    "Yacute": "Ý",
    "agrave": "à",
    "aacute": "á",
    "acircumflex": "â",
    "atilde": "ã",
    "egrave": "è",
    "eacute": "é",
    "ecircumflex": "ê",
    "igrave": "ì",
    "iacute": "í",
    "ograve": "ò",
    "oacute": "ó",
    "ocircumflex": "ô",
    "otilde": "õ",
    "ugrave": "ù",
    "uacute": "ú",
    "yacute": "ý",
    "Abreve": "Ă",
    "abreve": "ă",
    "Dcroat": "Đ",
    "dcroat": "đ",
    "DJ": "Đ",
    "dj": "đ",
    "Eth": "Đ",
    "eth": "đ",
    "Ohorn": "Ơ",
    "ohorn": "ơ",
    "Uhorn": "Ư",
    "uhorn": "ư",
    "ihookabove": "ỉ",
    "Ihookabove": "Ỉ",
    "idotbelow": "ị",
    "Idotbelow": "Ị",
    "ahookabove": "ả",
    "Ahookabove": "Ả",
    "adotbelow": "ạ",
    "Adotbelow": "Ạ",
    "acircumflexacute": "ấ",
    "Acircumflexacute": "Ấ",
    "acircumflexgrave": "ầ",
    "Acircumflexgrave": "Ầ",
    "acircumflexhookabove": "ẩ",
    "Acircumflexhookabove": "Ẩ",
    "acircumflextilde": "ẫ",
    "Acircumflextilde": "Ẫ",
    "acircumflexdotbelow": "ậ",
    "Acircumflexdotbelow": "Ậ",
    "abreveacute": "ắ",
    "Abreveacute": "Ắ",
    "abrevegrave": "ằ",
    "Abrevegrave": "Ằ",
    "abrevehookabove": "ẳ",
    "Abrevehookabove": "Ẳ",
    "abrevetilde": "ẵ",
    "Abrevetilde": "Ẵ",
    "abrevedotbelow": "ặ",
    "Abrevedotbelow": "Ặ",
    "ehookabove": "ẻ",
    "Ehookabove": "Ẻ",
    "etilde": "ẽ",
    "Etilde": "Ẽ",
    "edotbelow": "ẹ",
    "Edotbelow": "Ẹ",
    "ecircumflexacute": "ế",
    "Ecircumflexacute": "Ế",
    "ecircumflexgrave": "ề",
    "Ecircumflexgrave": "Ề",
    "ecircumflexhookabove": "ể",
    "Ecircumflexhookabove": "Ể",
    "ecircumflextilde": "ễ",
    "Ecircumflextilde": "Ễ",
    "ecircumflexdotbelow": "ệ",
    "Ecircumflexdotbelow": "Ệ",
    "ohookabove": "ỏ",
    "Ohookabove": "Ỏ",
    "odotbelow": "ọ",
    "Odotbelow": "Ọ",
    "ocircumflexacute": "ố",
    "Ocircumflexacute": "Ố",
    "ocircumflexgrave": "ồ",
    "Ocircumflexgrave": "Ồ",
    "ocircumflexhookabove": "ổ",
    "Ocircumflexhookabove": "Ổ",
    "ocircumflextilde": "ỗ",
    "Ocircumflextilde": "Ỗ",
    "ocircumflexdotbelow": "ộ",
    "Ocircumflexdotbelow": "Ộ",
    "ohornacute": "ớ",
    "Ohornacute": "Ớ",
    "ohorngrave": "ờ",
    "Ohorngrave": "Ờ",
    "ohornhookabove": "ở",
    "Ohornhookabove": "Ở",
    "ohorntilde": "ỡ",
    "Ohorntilde": "Ỡ",
    "ohorndotbelow": "ợ",
    "Ohorndotbelow": "Ợ",
    "uhookabove": "ủ",
    "Uhookabove": "Ủ",
    "utilde": "ũ",
    "Utilde": "Ũ",
    "udotbelow": "ụ",
    "Udotbelow": "Ụ",
    "uhornacute": "ứ",
    "Uhornacute": "Ứ",
    "uhorngrave": "ừ",
    "Uhorngrave": "Ừ",
    "uhornhookabove": "ử",
    "Uhornhookabove": "Ử",
    "uhorntilde": "ữ",
    "Uhorntilde": "Ữ",
    "uhorndotbelow": "ự",
    "Uhorndotbelow": "Ự",
    "ygrave": "ỳ",
    "Ygrave": "Ỳ",
    "yhookabove": "ỷ",
    "Yhookabove": "Ỷ",
    "ytilde": "ỹ",
    "Ytilde": "Ỹ",
    "ydotbelow": "ỵ",
    "Ydotbelow": "Ỵ",
}

# Standard Computer Modern Symbol (CMSY10 / CMSY8 / CMSY7) fallback table for slots 0..127
CMSY_FALLBACK: dict[int, str] = {
    0: "-",
    1: r"\cdot ",
    2: r"\times ",
    3: "*",
    4: r"\div ",
    5: r"\diamond ",
    6: r"\pm ",
    7: r"\mp ",
    8: r"\oplus ",
    9: r"\ominus ",
    10: r"\otimes ",
    11: r"\oslash ",
    12: r"\odot ",
    13: r"\bigcirc ",
    14: r"\circ ",
    15: r"\bullet ",
    16: r"\asymp ",
    17: r"\equiv ",
    18: r"\subseteq ",
    19: r"\supseteq ",
    20: r"\le ",
    21: r"\ge ",
    22: r"\preceq ",
    23: r"\succeq ",
    24: r"\sim ",
    25: r"\approx ",
    26: r"\subset ",
    27: r"\supset ",
    28: r"\ll ",
    29: r"\gg ",
    30: r"\prec ",
    31: r"\succ ",
    32: r"\leftarrow ",
    33: r"\to ",
    34: r"\uparrow ",
    35: r"\downarrow ",
    36: r"\leftrightarrow ",
    37: r"\nearrow ",
    38: r"\searrow ",
    39: r"\simeq ",
    40: r"\Leftarrow ",
    41: r"\Rightarrow ",
    42: r"\Uparrow ",
    43: r"\Downarrow ",
    44: r"\Leftrightarrow ",
    45: r"\nwarrow ",
    46: r"\swarrow ",
    47: r"\propto ",
    48: "'",
    49: r"\infty ",
    50: r"\in ",
    51: r"\ni ",
    52: r"\triangle ",
    53: r"\bigtriangledown ",
    54: "/",
    56: r"\forall ",
    57: r"\exists ",
    58: r"\neg ",
    59: r"\emptyset ",
    60: r"\Re ",
    61: r"\Im ",
    62: r"\top ",
    63: r"\perp ",
    91: r"\cup ",
    92: r"\cap ",
    93: r"\uplus ",
    94: r"\land ",
    95: r"\lor ",
    102: "{",
    103: "}",
    104: r"\langle ",
    105: r"\rangle ",
    106: "|",
    107: r"\Vert ",
    110: r"\setminus ",
    112: "@@RAD@@",
}

# Standard Computer Modern Math Extension (CMEX10) fallback table
CMEX_FALLBACK: dict[int, str] = {
    0: "(",
    1: ")",
    2: "[",
    3: "]",
    8: "{",
    9: "}",
    16: "(",
    17: ")",
    18: "(",
    19: ")",
    20: "[",
    21: "]",
    22: "[",
    23: "]",
    26: "{",
    27: "}",
    28: "{",
    29: "}",
    32: "(",
    33: ")",
    34: "[",
    35: "]",
    40: "{",
    41: "}",
    48: "(",
    50: "[",
    52: "[",
    54: "|",
    56: "{",
    58: "{",
    60: "{",
    62: "|",
    64: ")",
    66: "]",
    68: "]",
    70: "@@RAD@@",
    71: "@@RAD@@",
    72: r"\oint ",
    73: r"\oint ",
    80: r"\sum ",
    81: r"\prod ",
    82: r"\int ",
    83: r"\bigcup ",
    84: r"\bigcap ",
    88: r"\sum ",
    89: r"\prod ",
    90: r"\int ",
    91: r"\bigcup ",
    92: r"\bigcap ",
    104: "[",
    105: "]",
    110: "{",
    111: "}",
    112: "@@RAD@@",
    113: "@@RAD@@",
    114: "@@RAD@@",
    115: "@@RAD@@",
    116: "@@RAD_B@@",
    117: "@@RAD_V@@",
}

# Standard Computer Modern Math Italic (CMMI10 / CMMI8) fallback table for slots 0..63
CMMI_FALLBACK: dict[int, str] = {
    0: r"\Gamma ",
    1: r"\Delta ",
    2: r"\Theta ",
    3: r"\Lambda ",
    4: r"\Xi ",
    5: r"\Pi ",
    6: r"\Sigma ",
    7: r"\Upsilon ",
    8: r"\Phi ",
    9: r"\Psi ",
    10: r"\Omega ",
    11: r"\alpha ",
    12: r"\beta ",
    13: r"\gamma ",
    14: r"\delta ",
    15: r"\epsilon ",
    16: r"\zeta ",
    17: r"\eta ",
    18: r"\theta ",
    19: r"\iota ",
    20: r"\kappa ",
    21: r"\lambda ",
    22: r"\mu ",
    23: r"\nu ",
    24: r"\xi ",
    25: r"\pi ",
    26: r"\rho ",
    27: r"\sigma ",
    28: r"\tau ",
    29: r"\upsilon ",
    30: r"\phi ",
    31: r"\chi ",
    32: r"\psi ",
    33: r"\omega ",
    34: r"\varepsilon ",
    35: r"\vartheta ",
    36: r"\varpi ",
    37: r"\varrho ",
    38: r"\varsigma ",
    39: r"\varphi ",
    58: ".",
    59: ",",
    60: "<",
    61: "/",
    62: ">",
    96: r"\ell ",
    123: r"\imath ",
    124: r"\jmath ",
}

# Standard Computer Modern Roman (CMR10 / CMR8 / CMR7) fallback table for slots 0..31
CMR_FALLBACK: dict[int, str] = {
    0: r"\Gamma ",
    1: r"\Delta ",
    2: r"\Theta ",
    3: r"\Lambda ",
    4: r"\Xi ",
    5: r"\Pi ",
    6: r"\Sigma ",
    7: r"\Upsilon ",
    8: r"\Phi ",
    9: r"\Psi ",
    10: r"\Omega ",
    11: "ff",
    12: "fi",
    13: "fl",
    14: "ffi",
    15: "ffl",
    16: "ı",
    17: "ȷ",
    18: "`",
    19: "´",
    20: "ˇ",
    21: "˘",
    22: "¯",
    23: "˚",
    24: "¸",
    25: "ß",
    26: "æ",
    27: "œ",
    28: "ø",
    29: "Æ",
    30: "Œ",
    31: "Ø",
    34: "”",
    60: "¡",
    62: "¿",
    92: "“",
    123: "–",
    124: "—",
}

# AMS Blackboard Bold (MSBM10)
MSBM_FALLBACK: dict[int, str] = {
    ord("R"): r"\mathbb{R}",
    ord("N"): r"\mathbb{N}",
    ord("Z"): r"\mathbb{Z}",
    ord("Q"): r"\mathbb{Q}",
    ord("C"): r"\mathbb{C}",
    ord("P"): r"\mathbb{P}",
    ord("E"): r"\mathbb{E}",
    ord("F"): r"\mathbb{F}",
    ord("K"): r"\mathbb{K}",
    63: r"\emptyset ",
}


@dataclass
class _PdfSpan:
    page: int
    x: float
    y: float
    text: str
    font_name: str
    font_size: float
    end_x: float = 0.0


@dataclass
class _PdfHBar:
    page: int
    x: float
    y: float
    w: float


def _decompress_pdf_stream(raw_stream: bytes) -> bytes | None:
    raw = raw_stream.strip(b"\r\n")
    for wbits in (zlib.MAX_WBITS, -zlib.MAX_WBITS):
        try:
            return zlib.decompress(raw, wbits)
        except Exception:
            pass
    return None


def _parse_differences_array(diff_str: str) -> dict[int, str]:
    mapping: dict[int, str] = {}
    code = 0
    for tok in re.findall(r"\d+|/[A-Za-z0-9_.]+", diff_str):
        if tok.isdigit():
            code = int(tok)
        elif tok.startswith("/"):
            gname = tok[1:]
            mapping[code] = GLYPH_TO_TEXT.get(gname, gname[:1] if len(gname) == 1 else "")
            code += 1
    return mapping


def _clean_tounicode_target(u_str: str) -> str:
    """Clean Unicode string from a PDF `/ToUnicode` CMap (strip Variation Selectors, soft hyphens, etc.)."""
    if not u_str:
        return ""
    u_str = u_str.replace("\u2423", " ")
    # Check if this is a CMEX radical with Variation Selector (e.g. \u221A\uFE00)
    if "\u221a" in u_str or "√" in u_str:
        return "@@RAD@@"
    # Strip Variation Selectors (\uFE00..\uFE0F), Zero-Width Non-Joiner/Joiner (\u200B..\u200D), Soft Hyphen (\u00AD)
    u_str = re.sub(r"[\uFE00-\uFE0F\u200B-\u200D\u00AD\uFEFF]", "", u_str)
    return sanitize_xml_text(u_str)


def _parse_tounicode_cmap(cmap_text: str) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for block in re.findall(r"beginbfchar(.*?)endbfchar", cmap_text, re.S):
        for src_hex, dst_hex in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            code = int(src_hex, 16)
            chars = "".join(
                chr(int(dst_hex[i : i + 4], 16))
                for i in range(0, len(dst_hex), 4)
                if len(dst_hex[i : i + 4]) == 4
            )
            cleaned = _clean_tounicode_target(chars)
            if cleaned:
                mapping[code] = cleaned
    for block in re.findall(r"beginbfrange(.*?)endbfrange", block_src := cmap_text, re.S):
        for s_hex, e_hex, d_hex in re.findall(
            r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block
        ):
            start_c = int(s_hex, 16)
            end_c = int(e_hex, 16)
            base_u = int(d_hex, 16)
            for offset in range(max(0, end_c - start_c + 1)):
                cleaned = _clean_tounicode_target(chr(base_u + offset))
                if cleaned:
                    mapping[start_c + offset] = cleaned
        # Also support array form: <start> <end> [ <u1> <u2> ... ]
        for s_hex, e_hex, arr_body in re.findall(
            r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*?)\]", block, re.S
        ):
            start_c = int(s_hex, 16)
            hex_items = re.findall(r"<([0-9A-Fa-f]+)>", arr_body)
            for idx, h_item in enumerate(hex_items):
                chars = "".join(
                    chr(int(h_item[i : i + 4], 16))
                    for i in range(0, len(h_item), 4)
                    if len(h_item[i : i + 4]) == 4
                )
                cleaned = _clean_tounicode_target(chars)
                if cleaned:
                    mapping[start_c + idx] = cleaned
    return mapping


def _parse_type1_fontfile_encoding(raw_font_bytes: bytes) -> dict[int, str]:
    """Parse cleartext header of an embedded Type 1 `/FontFile` stream for `dup <code> /<glyph> put`."""
    mapping: dict[int, str] = {}
    text = raw_font_bytes.decode("latin1", errors="ignore")
    for m in re.finditer(r"\bdup\s+(\d+)\s+/([A-Za-z0-9_.]+)\s+put\b", text):
        code = int(m.group(1))
        gname = m.group(2)
        if gname in GLYPH_TO_TEXT:
            mapping[code] = GLYPH_TO_TEXT[gname]
        elif len(gname) == 1 and gname.isprintable():
            mapping[code] = gname
    return mapping


def _builtin_font_fallback(base_font: str) -> dict[int, str]:
    """Return built-in glyph mapping table for standard TeX math fonts."""
    bf = base_font.upper()
    if "CMSY" in bf or "MSAM" in bf:
        return dict(CMSY_FALLBACK)
    if "CMEX" in bf:
        return dict(CMEX_FALLBACK)
    if "CMMI" in bf:
        return dict(CMMI_FALLBACK)
    if "MSBM" in bf:
        return dict(MSBM_FALLBACK)
    if any(k in bf for k in ("CMR", "CMBX", "CMTI", "CMSL", "CMSS", "CMTT")):
        return dict(CMR_FALLBACK)
    return {}


def _decode_pdf_literal_bytes(inner: str) -> list[int]:
    out: list[int] = []
    i = 0
    n = len(inner)
    while i < n:
        ch = inner[i]
        if ch == "\\":
            i += 1
            if i >= n:
                break
            esc = inner[i]
            if esc in "01234567":
                oct_s = esc
                i += 1
                for _ in range(2):
                    if i < n and inner[i] in "01234567":
                        oct_s += inner[i]
                        i += 1
                    else:
                        break
                out.append(int(oct_s, 8) & 0xFF)
                continue
            esc_map = {"n": 10, "r": 13, "t": 9, "b": 8, "f": 12, "(": 40, ")": 41, "\\": 92}
            out.append(esc_map.get(esc, ord(esc) & 0xFF))
            i += 1
        else:
            out.append(ord(ch) & 0xFF)
            i += 1
    return out


def _unpack_pdf_object_streams(
    objects: dict[int, bytes],
    streams: dict[int, bytes],
) -> None:
    """Unpack PDF 1.5+ `/Type /ObjStm` compressed object streams into `objects` dictionary."""
    objstm_ids = [
        oid
        for oid, body in list(objects.items())
        if b"/ObjStm" in body and oid in streams
    ]
    for oid in objstm_ids:
        header_dict = objects[oid].decode("latin1", errors="ignore")
        m_n = re.search(r"/N\s+(\d+)", header_dict)
        m_first = re.search(r"/First\s+(\d+)", header_dict)
        if not m_n or not m_first:
            continue
        num_objs = int(m_n.group(1))
        first_offset = int(m_first.group(1))
        dec = streams[oid]
        if first_offset > len(dec):
            continue
        idx_header = dec[:first_offset].decode("latin1", errors="ignore")
        pairs = re.findall(r"(\d+)\s+(\d+)", idx_header)
        if not pairs:
            continue
        pairs = pairs[:num_objs]
        for i, (sub_oid_str, rel_off_str) in enumerate(pairs):
            sub_oid = int(sub_oid_str)
            start_pos = first_offset + int(rel_off_str)
            end_pos = (
                first_offset + int(pairs[i + 1][1])
                if i + 1 < len(pairs)
                else len(dec)
            )
            if 0 <= start_pos <= end_pos <= len(dec):
                if sub_oid not in objects:
                    objects[sub_oid] = dec[start_pos:end_pos]


def _collect_ordered_page_ids(
    objects: dict[int, bytes],
) -> list[tuple[int, str]]:
    """Walk the `/Catalog -> /Pages -> /Kids` tree in exact document order and return
    `[(page_obj_id, effective_page_dict_str_with_inherited_resources), ...]`.
    """
    obj_strs: dict[int, str] = {
        oid: raw.decode("latin1", errors="ignore") for oid, raw in objects.items()
    }

    root_pages_id: int | None = None
    for oid, s in obj_strs.items():
        if re.search(r"/Type\s*/Catalog\b", s):
            m_p = re.search(r"/Pages\s+(\d+)\s+0\s+R", s)
            if m_p:
                root_pages_id = int(m_p.group(1))
                break

    if root_pages_id is None:
        # Find top-level /Type /Pages without /Parent
        for oid, s in obj_strs.items():
            if re.search(r"/Type\s*/Pages\b", s) and "/Parent" not in s:
                root_pages_id = oid
                break

    ordered_pages: list[tuple[int, str]] = []
    visited: set[int] = set()

    def _walk(node_id: int, inherited_res: str) -> None:
        if node_id in visited or node_id not in obj_strs:
            return
        visited.add(node_id)
        s = obj_strs[node_id]

        # Check if this node defines /Resources
        cur_res = inherited_res
        if "/Resources" in s:
            cur_res = s

        if re.search(r"/Type\s*/Pages\b", s):
            m_kids = re.search(r"/Kids\s*\[(.*?)\]", s, re.S)
            if m_kids:
                for kid_m in re.finditer(r"(\d+)\s+0\s+R", m_kids.group(1)):
                    _walk(int(kid_m.group(1)), cur_res)
        elif re.search(r"/Type\s*/Page\b", s):
            eff_s = s
            if "/Resources" not in s and cur_res:
                eff_s = s + "\n" + cur_res
            ordered_pages.append((node_id, eff_s))

    if root_pages_id is not None:
        _walk(root_pages_id, "")

    if not ordered_pages:
        # Fallback: scan all /Type /Page objects
        for oid in sorted(obj_strs.keys()):
            s = obj_strs[oid]
            if re.search(r"/Type\s*/Page\b", s) and not re.search(r"/Type\s*/Pages\b", s):
                ordered_pages.append((oid, s))

    return ordered_pages


def extract_pdf_spans_pure_python(pdf_path: Path) -> tuple[list[_PdfSpan], list[_PdfHBar]]:
    """Extract 2D positioned text/math spans and horizontal fraction/radical bars from any PDF."""
    data = pdf_path.read_bytes()
    obj_re = re.compile(rb"(\d+)\s+(\d+)\s+obj\b(.*?)endobj", re.S)
    objects: dict[int, bytes] = {}
    streams: dict[int, bytes] = {}

    for m in obj_re.finditer(data):
        oid = int(m.group(1))
        body = m.group(3)
        objects[oid] = body
        sm = re.search(rb"stream\r?\n(.*?)\r?\n?endstream", body, re.S)
        if sm:
            raw_s = sm.group(1)
            dec = _decompress_pdf_stream(raw_s)
            streams[oid] = dec if dec is not None else raw_s

    # Unpack PDF 1.5+ /Type /ObjStm compressed object streams!
    _unpack_pdf_object_streams(objects, streams)

    # Build 4-layer font mapping per font object ID
    font_maps: dict[int, dict[int, str]] = {}
    font_names: dict[int, str] = {}

    for oid, body in objects.items():
        b_str = body.decode("latin1", errors="ignore")
        if "/Font" not in b_str and "/BaseFont" not in b_str:
            continue
        m_bf = re.search(r"/BaseFont\s*/([^\s/\[\]<>()]+)", b_str)
        if not m_bf:
            continue
        bf_raw = m_bf.group(1)
        bf_clean = bf_raw.split("+")[-1]
        font_names[oid] = bf_clean

        # Layer 4: Built-in TeX math font fallback table
        fmap: dict[int, str] = _builtin_font_fallback(bf_clean)

        # Layer 3: Type 1 /FontFile cleartext encoding (`dup <code> /<glyph> put`)
        m_desc = re.search(r"/FontDescriptor\s+(\d+)\s+0\s+R", b_str)
        if m_desc:
            desc_id = int(m_desc.group(1))
            desc_str = objects.get(desc_id, b"").decode("latin1", errors="ignore")
            m_ff = re.search(r"/FontFile[23]?\s+(\d+)\s+0\s+R", desc_str)
            if m_ff:
                ff_id = int(m_ff.group(1))
                if ff_id in streams:
                    fmap.update(_parse_type1_fontfile_encoding(streams[ff_id]))

        # Layer 2: /Encoding /Differences
        m_enc = re.search(r"/Encoding\s+(\d+)\s+0\s+R", b_str)
        if m_enc:
            enc_id = int(m_enc.group(1))
            enc_body = objects.get(enc_id, b"").decode("latin1", errors="ignore")
            m_diff = re.search(r"/Differences\s*\[(.*?)\]", enc_body, re.S)
            if m_diff:
                fmap.update(_parse_differences_array(m_diff.group(1)))
        else:
            m_diff_inline = re.search(r"/Differences\s*\[(.*?)\]", b_str, re.S)
            if m_diff_inline:
                fmap.update(_parse_differences_array(m_diff_inline.group(1)))

        # Layer 1: /ToUnicode CMap (highest priority, with CMEX radical & math operator protection)
        m_tu = re.search(r"/ToUnicode\s+(\d+)\s+0\s+R", b_str)
        if m_tu:
            tu_id = int(m_tu.group(1))
            if tu_id in streams:
                tu_map = _parse_tounicode_cmap(streams[tu_id].decode("latin1", errors="ignore"))
                for k_code, v_str in tu_map.items():
                    # Preserve @@RAD@@, @@RAD_V@@, @@RAD_B@@ in CMEX and CMSY fonts
                    if "CMEX" in bf_clean.upper() and k_code in (70, 71, 112, 113, 114, 115, 116, 117):
                        continue
                    if "CMSY" in bf_clean.upper() and k_code == 112:
                        continue
                    fmap[k_code] = v_str

        font_maps[oid] = fmap

    ordered_pages = _collect_ordered_page_ids(objects)

    spans: list[_PdfSpan] = []
    hbars: list[_PdfHBar] = []

    for page_idx, (oid, page_dict_str) in enumerate(ordered_pages, start=1):
        # Resolve page /Resources -> /Font mapping (/F1 -> font_oid)
        page_fonts: dict[str, int] = {}

        def _collect_fonts_from_str(src_str: str) -> None:
            for fm in re.finditer(r"/(F\d+|T\d+|C\d+_?\d*)\s+(\d+)\s+0\s+R", src_str):
                page_fonts[fm.group(1)] = int(fm.group(2))
            for fdict_m in re.finditer(r"/Font\s*<<(.*?)>>", src_str, re.S):
                for fm in re.finditer(r"/([A-Za-z0-9_.+-]+)\s+(\d+)\s+0\s+R", fdict_m.group(1)):
                    page_fonts[fm.group(1)] = int(fm.group(2))
            for fref_m in re.finditer(r"/Font\s+(\d+)\s+0\s+R", src_str):
                fdict_id = int(fref_m.group(1))
                fdict_body = objects.get(fdict_id, b"").decode("latin1", errors="ignore")
                for fm in re.finditer(r"/([A-Za-z0-9_.+-]+)\s+(\d+)\s+0\s+R", fdict_body):
                    page_fonts[fm.group(1)] = int(fm.group(2))

        _collect_fonts_from_str(page_dict_str)
        for res_m in re.finditer(r"/Resources\s+(\d+)\s+0\s+R", page_dict_str):
            res_id = int(res_m.group(1))
            res_str = objects.get(res_id, b"").decode("latin1", errors="ignore")
            _collect_fonts_from_str(res_str)

        # Resolve /Contents stream(s)
        content_ids: list[int] = []
        m_carr = re.search(r"/Contents\s*\[(.*?)\]", page_dict_str, re.S)
        if m_carr:
            content_ids = [int(x) for x in re.findall(r"(\d+)\s+0\s+R", m_carr.group(1))]
        else:
            m_csingle = re.search(r"/Contents\s+(\d+)\s+0\s+R", page_dict_str)
            if m_csingle:
                content_ids = [int(m_csingle.group(1))]

        content_bytes = b"\n".join(streams.get(cid, b"") for cid in content_ids)
        if not content_bytes:
            continue
        content_str = content_bytes.decode("latin1", errors="ignore")

        # Extract horizontal fraction/radical bars (`x y w h re f` in xelatex AND `0 0 m <w> 0 l S` in pdflatex)
        ctm_stack: list[tuple[float, float]] = []
        cur_x, cur_y = 0.0, 0.0
        for line in content_str.splitlines():
            ls = line.strip()
            if ls == "q":
                ctm_stack.append((cur_x, cur_y))
                continue
            if ls == "Q":
                if ctm_stack:
                    cur_x, cur_y = ctm_stack.pop()
                else:
                    cur_x, cur_y = 0.0, 0.0
                continue
            for m_cm in re.finditer(
                r"1\s+0\s+0\s+1\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+cm\b",
                ls,
            ):
                cur_x += float(m_cm.group(1))
                cur_y += float(m_cm.group(2))
            m_re = re.search(
                r"([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+re\s+f\b",
                ls,
            )
            if m_re:
                rx, ry, rw, rh = (
                    float(m_re.group(1)),
                    float(m_re.group(2)),
                    float(m_re.group(3)),
                    float(m_re.group(4)),
                )
                if 3.0 < rw < 260.0 and abs(rh) < 2.2:
                    hbars.append(_PdfHBar(page=page_idx, x=cur_x + rx, y=cur_y + ry, w=rw))
            for m_ln in re.finditer(
                r"([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+m\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+l\s+S\b",
                ls,
            ):
                lx1, ly1, lx2, ly2 = (
                    float(m_ln.group(1)),
                    float(m_ln.group(2)),
                    float(m_ln.group(3)),
                    float(m_ln.group(4)),
                )
                rw = abs(lx2 - lx1)
                if abs(ly2 - ly1) < 0.2 and 3.0 < rw < 240.0:
                    hbars.append(_PdfHBar(page=page_idx, x=cur_x + min(lx1, lx2), y=cur_y + ly1, w=rw))

        # Parse BT ... ET text blocks
        for bt_block in re.findall(r"\bBT\b(.*?)\bET\b", content_str, re.S):
            cur_font_key = ""
            cur_font_size = 12.0
            tm_x, tm_y = 0.0, 0.0
            line_x, line_y = 0.0, 0.0
            leading = 0.0

            tok_pattern = re.compile(
                r"/(F\d+|T\d+|C\d+_?\d*|[A-Za-z0-9_.+-]+)\s+([+-]?\d+(?:\.\d+)?)\s+Tf"
                r"|([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+Tm"
                r"|([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+Td"
                r"|([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+TD"
                r"|([+-]?\d+(?:\.\d+)?)\s+TL"
                r"|(T\*)"
                r"|\[(.*?)\]\s*TJ"
                r"|\(((?:\\.|[^\\()])*)\)\s*Tj"
                r"|<([0-9A-Fa-f]+)>\s*Tj",
                re.S,
            )

            def decode_str_token(raw_lit: str | None, raw_hex: str | None, font_key: str) -> str:
                f_oid = page_fonts.get(font_key, -1)
                fmap = font_maps.get(f_oid, {})
                b_list: list[int] = []
                if raw_lit is not None:
                    b_list = _decode_pdf_literal_bytes(raw_lit)
                elif raw_hex is not None:
                    hx = re.sub(r"\s+", "", raw_hex)
                    b_list = [int(hx[k : k + 2], 16) for k in range(0, len(hx), 2) if len(hx[k : k + 2]) == 2]
                chars: list[str] = []
                for b in b_list:
                    if b in fmap:
                        chars.append(fmap[b])
                    elif 32 <= b <= 126:
                        chars.append(chr(b))
                    elif b >= 160:
                        chars.append(chr(b))
                    # Never append raw control bytes (0..31 or 127..159) — prevents XMLSyntaxError!
                return sanitize_xml_text("".join(chars))

            for tm in tok_pattern.finditer(bt_block):
                if tm.group(1) is not None:
                    cur_font_key = tm.group(1)
                    cur_font_size = abs(float(tm.group(2))) or 12.0
                elif tm.group(3) is not None:
                    a_scale = abs(float(tm.group(3)))
                    d_scale = abs(float(tm.group(6)))
                    tm_x = float(tm.group(7))
                    tm_y = float(tm.group(8))
                    line_x, line_y = tm_x, tm_y
                    if a_scale > 0.1 and cur_font_size <= 1.5:
                        cur_font_size = d_scale or a_scale
                elif tm.group(9) is not None:
                    dx, dy = float(tm.group(9)), float(tm.group(10))
                    line_x += dx
                    line_y += dy
                    tm_x, tm_y = line_x, line_y
                elif tm.group(11) is not None:
                    dx, dy = float(tm.group(11)), float(tm.group(12))
                    leading = -dy
                    line_x += dx
                    line_y += dy
                    tm_x, tm_y = line_x, line_y
                elif tm.group(13) is not None:
                    leading = float(tm.group(13))
                elif tm.group(14) is not None:
                    line_y -= leading
                    tm_x, tm_y = line_x, line_y
                elif tm.group(15) is not None:
                    arr_body = tm.group(15)
                    f_oid = page_fonts.get(cur_font_key, -1)
                    fname = font_names.get(f_oid, cur_font_key)
                    cursor_x = tm_x
                    sub_pieces: list[Tuple[float, str, float]] = []
                    cur_str_Accum: list[str] = []
                    seg_start_x = cursor_x
                    for sub_m in re.finditer(
                        r"\(((?:\\.|[^\\()])*)\)|<([0-9A-Fa-f]+)>|([+-]?\d+(?:\.\d+)?)",
                        arr_body,
                    ):
                        if sub_m.group(1) is not None or sub_m.group(2) is not None:
                            piece = decode_str_token(sub_m.group(1), sub_m.group(2), cur_font_key)
                            if not cur_str_Accum:
                                seg_start_x = cursor_x
                            cur_str_Accum.append(piece)
                            vis_len = max(1, len(piece.strip()))
                            cursor_x += vis_len * cur_font_size * 0.48
                        else:
                            kern = float(sub_m.group(3))
                            shift_pt = -kern * 0.001 * cur_font_size
                            if shift_pt > cur_font_size * 1.35 and cur_str_Accum:
                                joined = "".join(cur_str_Accum)
                                if joined:
                                    spans.append(
                                        _PdfSpan(
                                            page=page_idx,
                                            x=seg_start_x,
                                            y=tm_y,
                                            text=joined,
                                            font_name=fname,
                                            font_size=cur_font_size,
                                            end_x=cursor_x,
                                        )
                                    )
                                cur_str_Accum = []
                            elif shift_pt > cur_font_size * 0.22:
                                cur_str_Accum.append(" ")
                            cursor_x += shift_pt
                    if cur_str_Accum:
                        joined = "".join(cur_str_Accum)
                        if joined:
                            spans.append(
                                _PdfSpan(
                                    page=page_idx,
                                    x=seg_start_x,
                                    y=tm_y,
                                    text=joined,
                                    font_name=fname,
                                    font_size=cur_font_size,
                                    end_x=cursor_x,
                                )
                            )
                    tm_x = cursor_x
                elif tm.group(16) is not None or tm.group(17) is not None:
                    f_oid = page_fonts.get(cur_font_key, -1)
                    fname = font_names.get(f_oid, cur_font_key)
                    txt = decode_str_token(tm.group(16), tm.group(17), cur_font_key)
                    if txt:
                        w_est = max(1, len(txt.strip())) * cur_font_size * 0.48
                        spans.append(
                            _PdfSpan(
                                page=page_idx,
                                x=tm_x,
                                y=tm_y,
                                text=txt,
                                font_name=fname,
                                font_size=cur_font_size,
                                end_x=tm_x + w_est,
                            )
                        )
                        tm_x += w_est

    return spans, hbars


def _is_math_font(font_name: str) -> bool:
    fn = font_name.upper()
    return any(k in fn for k in ("CMMI", "CMSY", "CMEX", "CMR", "CMBX", "MSBM", "MSAM", "EUFM", "EUSM"))


def _is_vietnamese_text_font(font_name: str) -> bool:
    fn = font_name.upper()
    if _is_math_font(fn):
        return False
    return any(
        k in fn
        for k in (
            "VNNIMBUS",
            "VNR",
            "VNBX",
            "VNTI",
            "VNSL",
            "VNTT",
            "VNSS",
            "PTM",
            "PUT",
            "PPL",
            "PHV",
            "TIMES",
            "ARIAL",
            "CALIBRI",
            "CAMBRIA",
            "CMU",
            "LMROMAN",
        )
    )


def _is_bold_text_font(font_name: str) -> bool:
    fn = font_name.upper()
    return any(k in fn for k in ("BOLD", "VNBX", "CMBX", "NIMBUSB", "-B", "BD"))


def _reconstruct_page_latex_from_2d(
    page_spans: list[_PdfSpan],
    page_bars: list[_PdfHBar],
) -> list[str]:
    """Reconstruct LaTeX paragraphs and equations from 2D positioned spans on a single PDF page."""
    if not page_spans:
        return []

    # Separate header/footer noise (page numbers, top running headers, watermarks)
    ys = [s.y for s in page_spans]
    max_y = max(ys)
    min_y = min(ys)

    filtered_spans: list[_PdfSpan] = []
    for sp in page_spans:
        t_clean = sp.text.strip()
        if not t_clean:
            continue
        # Skip giant background watermark (e.g. font_size > 40 or single 'K' at center)
        if sp.font_size > 40.0:
            continue
        # Skip standalone page numbers at very bottom or top
        if (sp.y > max_y - 18 or sp.y < min_y + 18) and (
            t_clean.isdigit() or "Trang " in t_clean or "Page " in t_clean
        ):
            continue
        filtered_spans.append(sp)

    if not filtered_spans:
        filtered_spans = [s for s in page_spans if s.text.strip()]

    # Identify extensible tall brackets `[` or `{` built from CMEX pieces stacked vertically
    cmex_brackets = [
        sp
        for sp in filtered_spans
        if "CMEX" in sp.font_name.upper() and sp.text.strip() in ("[", "{")
    ]
    non_cmex = [sp for sp in filtered_spans if sp not in cmex_brackets]

    # Group spans into logical horizontal lines (cluster by baseline y within 9pt,
    # while keeping fraction numerators/denominators and radical bars attached to their parent line)
    sorted_by_y = sorted(non_cmex, key=lambda s: (-s.y, s.x))

    # Build primary text baselines from normal-sized spans (font_size >= 9.5)
    clusters: list[list[_PdfSpan]] = []
    cluster_ys: list[float] = []

    for sp in sorted_by_y:
        placed = False
        tol = 7.5 if sp.font_size >= 9.5 else 11.5
        for idx, cy in enumerate(cluster_ys):
            if abs(sp.y - cy) <= tol:
                clusters[idx].append(sp)
                # Update baseline weighted toward normal font size
                if sp.font_size >= 9.5:
                    n_norm = sum(1 for x in clusters[idx] if x.font_size >= 9.5)
                    cluster_ys[idx] = sum(
                        x.y for x in clusters[idx] if x.font_size >= 9.5
                    ) / max(1, n_norm)
                placed = True
                break
        if not placed:
            clusters.append([sp])
            cluster_ys.append(sp.y)

    # Sort clusters top-to-bottom
    paired = sorted(zip(cluster_ys, clusters), key=lambda pair: -pair[0])
    cluster_ys = [p[0] for p in paired]
    clusters = [p[1] for p in paired]

    # Merge fraction numerators & denominators across adjacent clusters when a horizontal fraction bar lies between them
    used_bars: set[int] = set()
    merged_away: set[int] = set()

    for b_idx, bar in enumerate(page_bars):
        # Check if this bar is a radical overline (has a @@RAD@@ immediately to its left)
        is_rad_bar = any(
            "@@RAD" in sp.text and abs(sp.x - bar.x) < 18.0 and abs(sp.y - bar.y) < 30.0
            for sp in filtered_spans
        )
        if is_rad_bar:
            continue

        # Find math spans just above bar (numerator) and just below bar (denominator)
        num_spans = [
            sp
            for sp in non_cmex
            if (bar.x - 3.0 <= sp.x <= bar.x + bar.w - 0.5)
            and (0.5 <= sp.y - bar.y <= 15.0)
            and not _is_vietnamese_text_font(sp.font_name)
        ]
        den_spans = [
            sp
            for sp in non_cmex
            if (bar.x - 3.0 <= sp.x <= bar.x + bar.w - 0.5)
            and (0.5 <= bar.y - sp.y <= 15.0)
            and not _is_vietnamese_text_font(sp.font_name)
        ]
        if num_spans and den_spans:
            used_bars.add(b_idx)
            num_tex = _render_span_sequence_to_latex(num_spans, page_bars, math_context=True)
            den_tex = _render_span_sequence_to_latex(den_spans, page_bars, math_context=True)
            frac_span = _PdfSpan(
                page=bar.page,
                x=bar.x,
                y=bar.y - 2.5,
                text=f"\\dfrac{{{num_tex}}}{{{den_tex}}}",
                font_name="CMMI10",
                font_size=11.9,
                end_x=bar.x + bar.w,
            )
            # Remove num_spans and den_spans from clusters and insert frac_span into the cluster closest to bar.y
            num_den_set = set(id(x) for x in num_spans + den_spans)
            for c_i in range(len(clusters)):
                clusters[c_i] = [x for x in clusters[c_i] if id(x) not in num_den_set]
            # Find best cluster for frac_span
            best_ci = None
            best_dist = 999.0
            for c_i, cy in enumerate(cluster_ys):
                if clusters[c_i] and abs(cy - (bar.y - 2.5)) < best_dist:
                    best_dist = abs(cy - (bar.y - 2.5))
                    best_ci = c_i
            if best_ci is not None and best_dist <= 15.0:
                clusters[best_ci].append(frac_span)
            else:
                clusters.append([frac_span])
                cluster_ys.append(bar.y - 2.5)

    # Re-sort non-empty clusters top-to-bottom
    valid_pairs = [
        (cy, cl) for cy, cl in zip(cluster_ys, clusters) if cl
    ]
    valid_pairs.sort(key=lambda p: -p[0])

    # Detect multi-line CMEX bracket groups (\hoac / \heva) spanning 2+ vertical lines
    cmex_groups: list[tuple[str, float, float, float]] = []
    if cmex_brackets:
        # Cluster cmex_brackets by x coordinate
        for br in sorted(cmex_brackets, key=lambda b: b.x):
            placed_g = False
            for g_i, (g_chr, gx, gy_min, gy_max) in enumerate(cmex_groups):
                if abs(br.x - gx) < 8.0 and g_chr == br.text.strip():
                    cmex_groups[g_i] = (
                        g_chr,
                        gx,
                        min(gy_min, br.y - 14.0),
                        max(gy_max, br.y + 6.0),
                    )
                    placed_g = True
                    break
            if not placed_g:
                cmex_groups.append((br.text.strip(), br.x, br.y - 14.0, br.y + 6.0))

    raw_lines: list[tuple[float, float, str]] = []
    for cy, cl in valid_pairs:
        cl_sorted = sorted(cl, key=lambda s: s.x)
        first_x = cl_sorted[0].x
        line_str = _render_cluster_line_to_latex(cl_sorted, cy, page_bars)
        if line_str.strip():
            raw_lines.append((cy, first_x, line_str.strip()))

    # Combine lines that fall inside a multi-line CMEX `[` or `{` group into \hoac{...} or \heva{...}
    out_paragraphs: list[str] = []
    i = 0
    while i < len(raw_lines):
        cy, fx, l_str = raw_lines[i]
        matched_group = None
        for g_chr, gx, gy_min, gy_max in cmex_groups:
            if gy_min - 8.0 <= cy <= gy_max + 8.0 and fx >= gx - 5.0:
                matched_group = (g_chr, gx, gy_min, gy_max)
                break

        if matched_group is not None:
            g_chr, gx, gy_min, gy_max = matched_group
            group_rows: list[str] = []
            while i < len(raw_lines):
                cy2, fx2, l_str2 = raw_lines[i]
                if gy_min - 8.0 <= cy2 <= gy_max + 8.0 and fx2 >= gx - 5.0:
                    clean_row = l_str2.replace("$", "").strip()
                    if clean_row:
                        group_rows.append(clean_row)
                    i += 1
                else:
                    break
            if len(group_rows) >= 2:
                macro = r"\hoac" if g_chr == "[" else r"\heva"
                joined_rows = r" \\ ".join(group_rows)
                if out_paragraphs and out_paragraphs[-1].endswith(("$", r"\Leftrightarrow$", r"\Rightarrow$")):
                    prev = out_paragraphs[-1]
                    if prev.endswith("$"):
                        out_paragraphs[-1] = prev[:-1] + f" {macro}{{{joined_rows}}}$"
                    else:
                        out_paragraphs.append(f"${macro}{{{joined_rows}}}$")
                else:
                    out_paragraphs.append(f"\\[\n{macro}{{{joined_rows}}}\n\\]")
                continue
            elif len(group_rows) == 1:
                out_paragraphs.append(f"${group_rows[0]}$")
                continue

        # Check if this line is a 4-option multiple choice line (e.g. "A. ... B. ... C. ... D. ...")
        mc_converted = _try_convert_multiple_choice_line(l_str, raw_lines, i)
        if mc_converted is not None:
            mc_tex, consumed = mc_converted
            out_paragraphs.append(mc_tex)
            i += consumed
            continue

        out_paragraphs.append(l_str)
        i += 1

    return out_paragraphs


def _try_convert_multiple_choice_line(
    line_str: str,
    all_lines: list[tuple[float, float, str]],
    idx: int,
) -> tuple[str, int] | None:
    """Detect 1-row (A. B. C. D.) or 2-row (A. B. / C. D.) multiple choice options extracted from PDF."""
    s = re.sub(r"\\textbf\{\s*([A-D]\.)\s*\}", r"\1", line_str).strip()
    # 1-row pattern: A. ... B. ... C. ... D. ...
    m4 = re.match(
        r"^(?:\\textbf\{)?A\.(?:\})?\s*(.+?)\s*(?:\\textbf\{)?B\.(?:\})?\s*(.+?)\s*(?:\\textbf\{)?C\.(?:\})?\s*(.+?)\s*(?:\\textbf\{)?D\.(?:\})?\s*(.+)$",
        s,
    )
    if m4:
        opts = [m4.group(1).strip(), m4.group(2).strip(), m4.group(3).strip(), m4.group(4).strip()]
        return (f"\\choice{{{opts[0]}}}{{{opts[1]}}}{{{opts[2]}}}{{{opts[3]}}}", 1)

    # 2-row pattern: Line idx has A. ... B. ..., Line idx+1 has C. ... D. ...
    if idx + 1 < len(all_lines):
        s2 = re.sub(r"\\textbf\{\s*([A-D]\.)\s*\}", r"\1", all_lines[idx + 1][2]).strip()
        m_ab = re.match(r"^(?:\\textbf\{)?A\.(?:\})?\s*(.+?)\s*(?:\\textbf\{)?B\.(?:\})?\s*(.+)$", s)
        m_cd = re.match(r"^(?:\\textbf\{)?C\.(?:\})?\s*(.+?)\s*(?:\\textbf\{)?D\.(?:\})?\s*(.+)$", s2)
        if m_ab and m_cd:
            opts = [
                m_ab.group(1).strip(),
                m_ab.group(2).strip(),
                m_cd.group(1).strip(),
                m_cd.group(2).strip(),
            ]
            return (f"\\choice{{{opts[0]}}}{{{opts[1]}}}{{{opts[2]}}}{{{opts[3]}}}", 2)

    return None


def _render_span_sequence_to_latex(
    spans: list[_PdfSpan],
    page_bars: list[_PdfHBar],
    *,
    math_context: bool = False,
) -> str:
    """Render an x-sorted list of spans (such as inside a radicand, numerator, or denominator) into LaTeX math."""
    if not spans:
        return ""
    s_sorted = sorted(spans, key=lambda x: x.x)
    base_y = s_sorted[0].y
    for sp in s_sorted:
        if sp.font_size >= 9.5 and "@@RAD" not in sp.text:
            base_y = sp.y
            break

    out: list[str] = []
    i = 0
    n = len(s_sorted)
    while i < n:
        sp = s_sorted[i]
        txt = sp.text
        if "@@RAD_V@@" in txt or "@@RAD_B@@" in txt:
            i += 1
            continue
        if "@@RAD@@" in txt:
            # Find matching overline bar starting near sp.x
            rad_bar = None
            for b in page_bars:
                if abs(b.x - sp.x) < 18.0 and abs(b.y - sp.y) < 32.0:
                    rad_bar = b
                    break
            if rad_bar is not None:
                rad_end_x = rad_bar.x + rad_bar.w + 1.5
                inner_spans: list[_PdfSpan] = []
                j = i + 1
                while j < n and s_sorted[j].x < rad_end_x:
                    inner_spans.append(s_sorted[j])
                    j += 1
                inner_tex = _render_span_sequence_to_latex(inner_spans, page_bars, math_context=True)
                out.append(f"\\sqrt{{{inner_tex}}}")
                i = j
                continue
            else:
                # Fallback if no bar found: take next span as radicand
                if i + 1 < n:
                    inner_tex = _render_span_sequence_to_latex([s_sorted[i + 1]], page_bars, math_context=True)
                    out.append(f"\\sqrt{{{inner_tex}}}")
                    i += 2
                    continue
                i += 1
                continue

        # Check superscript / subscript relative to base_y
        dy = sp.y - base_y
        if sp.font_size < 9.8 and abs(dy) > 1.8 and out:
            script_group = [sp]
            j = i + 1
            while (
                j < n
                and s_sorted[j].font_size < 9.8
                and abs(s_sorted[j].y - sp.y) < 2.2
                and (s_sorted[j].x - script_group[-1].end_x) < 8.0
            ):
                script_group.append(s_sorted[j])
                j += 1
            script_str = "".join(x.text.strip() for x in script_group)
            if dy > 1.8:
                out.append(f"^{{{script_str}}}")
            else:
                out.append(f"_{{{script_str}}}")
            i = j
            continue

        out.append(txt)
        i += 1

    res = "".join(out).strip()
    return sanitize_xml_text(res)


def _render_cluster_line_to_latex(
    cl_sorted: list[_PdfSpan],
    baseline_y: float,
    page_bars: list[_PdfHBar],
) -> str:
    """Convert a horizontal line cluster of 2D PDF spans into mixed Vietnamese text and `$...$` LaTeX math."""
    # First resolve any radicals (`@@RAD@@`) and arrow-with-label (`\xrightarrow`) inside the cluster
    processed: list[_PdfSpan] = []
    i = 0
    n = len(cl_sorted)
    while i < n:
        sp = cl_sorted[i]
        if "@@RAD_V@@" in sp.text or "@@RAD_B@@" in sp.text:
            i += 1
            continue

        if "@@RAD@@" in sp.text:
            rad_bar = None
            for b in page_bars:
                if abs(b.x - sp.x) < 18.0 and abs(b.y - sp.y) < 32.0:
                    rad_bar = b
                    break
            if rad_bar is not None:
                rad_end_x = rad_bar.x + rad_bar.w + 1.5
                inner_spans: list[_PdfSpan] = []
                j = i + 1
                while j < n and cl_sorted[j].x < rad_end_x:
                    inner_spans.append(cl_sorted[j])
                    j += 1
                inner_tex = _render_span_sequence_to_latex(inner_spans, page_bars, math_context=True)
                processed.append(
                    _PdfSpan(
                        page=sp.page,
                        x=sp.x,
                        y=baseline_y,
                        text=f"\\sqrt{{{inner_tex}}}",
                        font_name="CMMI10",
                        font_size=11.9,
                        end_x=rad_end_x,
                    )
                )
                i = j
                continue

        # Check superscript / subscript attachment to previous math span
        dy = sp.y - baseline_y
        if sp.font_size < 9.8 and abs(dy) > 1.8 and processed:
            script_group = [sp]
            j = i + 1
            while (
                j < n
                and cl_sorted[j].font_size < 9.8
                and abs(cl_sorted[j].y - sp.y) < 2.2
                and (cl_sorted[j].x - script_group[-1].end_x) < 8.0
            ):
                script_group.append(cl_sorted[j])
                j += 1
            script_str = "".join(x.text.strip() for x in script_group)
            # Check if followed by an extensible arrow `-\to` (i.e. \xrightarrow{label})
            has_arrow_ahead = any(
                "\\to" in cl_sorted[m_idx].text for m_idx in range(j, min(n, j + 4))
            )
            if (
                dy > 1.8
                and j < n
                and has_arrow_ahead
                and ("\\to" in cl_sorted[j].text or cl_sorted[j].text.strip() in ("-", "−"))
            ):
                # Consume arrow pieces
                k = j
                while k < n and (
                    cl_sorted[k].text.strip() in ("-", "−", r"\to")
                    or "\\to" in cl_sorted[k].text
                ):
                    k += 1
                processed.append(
                    _PdfSpan(
                        page=sp.page,
                        x=sp.x,
                        y=baseline_y,
                        text=f"\\xrightarrow{{{script_str}}}",
                        font_name="CMSY10",
                        font_size=11.9,
                        end_x=cl_sorted[k - 1].end_x if k > j else sp.end_x,
                    )
                )
                i = k
                continue

            prev = processed[-1]
            op_char = "^" if dy > 1.8 else "_"
            prev.text = f"{prev.text.rstrip()}{op_char}{{{script_str}}}"
            prev.font_name = "CMMI10"
            prev.end_x = script_group[-1].end_x
            i = j
            continue

        processed.append(sp)
        i += 1

    # Now segment `processed` into Vietnamese text runs vs `$...$` math runs
    pieces: list[str] = []
    in_math = False
    math_buf: list[str] = []
    prev_end_x: float | None = None

    def flush_math() -> None:
        nonlocal in_math
        if not in_math:
            return
        m_str = "".join(math_buf).strip()
        math_buf.clear()
        in_math = False
        if m_str:
            # Clean redundant spaces around operators
            m_str = re.sub(r"\s+", " ", m_str)
            if m_str in (",", ".", ";", ":", "?", "!"):
                pieces.append(m_str)
            else:
                pieces.append(f"${m_str}$")

    for idx, sp in enumerate(processed):
        txt = sp.text
        if not txt:
            continue

        gap = (sp.x - prev_end_x) if prev_end_x is not None else 0.0
        prev_end_x = max(sp.end_x, sp.x + 2.0)

        is_vn = _is_vietnamese_text_font(sp.font_name)
        is_m = _is_math_font(sp.font_name) or any(
            k in txt for k in ("\\sqrt", "\\dfrac", "\\xrightarrow", "^", "_")
        )

        # Digits/punctuation in CMR10 right after Vietnamese words like "Bài 1.", "Câu 12:", "(1)"
        if (
            is_m
            and "CMR" in sp.font_name.upper()
            and not in_math
            and re.fullmatch(r"[\d.,:;()\[\]+-]+", txt.strip())
            and pieces
            and re.search(r"(?:Bài|Câu|Ví dụ|Đề|Phần|Chương|Cách|Bước|Từ|Thay|Vậy)\s*$", "".join(pieces))
        ):
            is_m = False
            is_vn = True

        if is_vn and not is_m:
            was_math = in_math
            flush_math()
            if (gap > 1.5 or (was_math and txt[:1].isalnum())) and pieces and not pieces[-1].endswith((" ", "\n", "(")):
                pieces.append(" ")
            if _is_bold_text_font(sp.font_name) and txt.strip():
                pieces.append(f"\\textbf{{{txt.strip()}}} ")
            else:
                pieces.append(txt)
        else:
            if not in_math:
                if pieces and not pieces[-1].endswith((" ", "\n", "(")):
                    pieces.append(" ")
                in_math = True
            elif gap > 14.0:
                # Separate distinct equations on the same line
                flush_math()
                pieces.append("   ")
                in_math = True
            elif gap > 3.5 and math_buf and not math_buf[-1].endswith(" "):
                math_buf.append(" ")
            math_buf.append(txt)

    flush_math()
    line_out = "".join(pieces).strip()
    # Merge adjacent math blocks `$a$ $b$` if separated only by comma/semicolon or space when appropriate
    line_out = re.sub(r"\$\s*([=+-])\s*\$", r" \1 ", line_out)
    line_out = re.sub(r"[ \t]{2,}", " ", line_out)
    return sanitize_xml_text(line_out)


def _extract_pdf_plain_probe(pdf_path: Path) -> str:
    """Extract a quick plain-text probe from a PDF for optional matching against existing `.tex` sources."""
    try:
        spans, _ = extract_pdf_spans_pure_python(pdf_path)
        if spans:
            return " ".join(s.text for s in spans[:250])
    except Exception:
        pass
    return ""


def _find_matching_tex_source(pdf_path: Path, search_dirs: list[Path] | None = None) -> Path | None:
    """Check if the uploaded PDF has an exact matching `.tex` source file in the workspace."""
    same_stem = pdf_path.with_suffix(".tex")
    if same_stem.is_file():
        return same_stem

    dirs: list[Path] = []
    if search_dirs:
        dirs.extend(search_dirs)
    dirs.append(pdf_path.parent)
    base_root = Path(__file__).resolve().parent.parent
    dirs.extend([base_root, base_root / "jobs"])

    pdf_probe = _extract_pdf_plain_probe(pdf_path)
    pdf_words = set(re.findall(r"[A-Za-zÀ-ỹ]{4,}", pdf_probe))

    candidates: list[Path] = []
    seen: set[Path] = set()
    for d in dirs:
        if not d.exists():
            continue
        for p in d.rglob("*.tex"):
            rp = p.resolve()
            if rp not in seen and "visual_build" not in rp.parts:
                seen.add(rp)
                candidates.append(rp)

    if not candidates or len(pdf_words) < 8:
        return None

    best_path: Path | None = None
    best_score = 0.0
    for cand in candidates:
        try:
            raw_tex = cand.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        tex_words = set(re.findall(r"[A-Za-zÀ-ỹ]{4,}", raw_tex))
        if not tex_words:
            continue
        overlap = len(pdf_words & tex_words) / max(len(pdf_words), 1)
        if overlap > best_score:
            best_score = overlap
            best_path = cand

    if best_path is not None and best_score >= 0.65:
        return best_path
    return None


def decompile_pdf_to_latex(
    pdf_path: Path,
    *,
    search_dirs: list[Path] | None = None,
    allow_tex_cache: bool = True,
) -> str:
    """Universal 2D Pure-Python Math PDF Decompiler (`decompile_pdf_to_latex`).

    Converts any Math PDF file into structured LaTeX (with `\\dfrac`, `\\sqrt`, `\\hoac`, `\\heva`,
    `\\xrightarrow`, `\\choice`, and Vietnamese Unicode) ready for `build_native_docx_from_latex`.
    """
    pdf_path = pdf_path.resolve()
    if not pdf_path.is_file():
        raise RuntimeError(f"PDF file not found: {pdf_path}")

    if allow_tex_cache:
        matched_tex = _find_matching_tex_source(pdf_path, search_dirs=search_dirs)
        if matched_tex is not None:
            return clean_newlines(matched_tex.read_text(encoding="utf-8", errors="replace"))

    spans, hbars = extract_pdf_spans_pure_python(pdf_path)
    if not spans:
        # Fallback to pypdf text extraction if PDF uses non-standard filters
        try:
            import pypdf

            reader = pypdf.PdfReader(str(pdf_path))
            pages_txt = [sanitize_xml_text(pg.extract_text() or "") for pg in reader.pages]
            body_txt = "\n\n".join(p.strip() for p in pages_txt if p.strip())
            if body_txt:
                return (
                    "\\documentclass[12pt,a4paper]{article}\n"
                    "\\usepackage{amsmath,amssymb}\n"
                    "\\begin{document}\n\n"
                    + body_txt
                    + "\n\n\\end{document}\n"
                )
        except Exception:
            pass

    max_page = max((s.page for s in spans), default=1)
    all_paragraphs: list[str] = []

    for pg in range(1, max_page + 1):
        pg_spans = [s for s in spans if s.page == pg]
        pg_bars = [b for b in hbars if b.page == pg]
        pg_paras = _reconstruct_page_latex_from_2d(pg_spans, pg_bars)
        all_paragraphs.extend(pg_paras)

    body_content = "\n\n".join(p for p in all_paragraphs if p.strip())
    latex_doc = (
        "\\documentclass[12pt,a4paper]{article}\n"
        "\\usepackage[utf8]{vietnam}\n"
        "\\usepackage{amsmath,amssymb}\n"
        "\\begin{document}\n\n"
        f"{body_content}\n\n"
        "\\end{document}\n"
    )
    return clean_newlines(latex_doc)


# ==============================================================================
# UNIFIED LATEX / STRUCTURAL ENVIRONMENT PARSER & DOCX BUILDER
# ==============================================================================
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


def _int_to_roman(num: int) -> str:
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman_num = ""
    i = 0
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
        marker = sanitize_xml_text(m_item.group(2).strip())
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
        "engine": "Native DrawingML + OMML + MathTables + Universal 2D PDF Decompiler Engine v2.4",
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
