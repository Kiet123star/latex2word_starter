"""Unified Core LaTeX & PDF-to-Word Converter Engine (`backend/core_converter.py` - v2.5 Universal).

Single unified conversion engine shared across:
1. CLI (`latex2word.py`)
2. FastAPI Web Server (`backend/server.py` / `run_server.py`)
3. Streamlit Web App (`app.py`)

Features:
- 100% Pure-Python 2D Math PDF Decompiler (`decompile_pdf_to_latex`, zero external PDF library dependencies) with:
  * Full `/ObjStm` compressed object stream unpacking & ordered `/Pages` tree traversal
  * Full 256-slot T5 (VnTeX) Vietnamese encoding & Type1 `/FontFile` cleartext glyph table parser
  * Persistence of active font state (`Tf`) across `BT ... ET` blocks and exact stream length slicing
  * Complete `CMEX10`, `Yhcmex`, `CMSY10`, `CMMI10`, `CMR10`, `MSBM10` math symbol & delimiter tables
  * Vertical absolute-value bar (`vextendsingle` / `|`) reconstruction & deduplication
  * Underscript attachment directly beneath `\max`, `\min`, `\lim`, `\sup`, `\inf` operators
  * Bounded multi-line system reconstruction (`\heva`, `\hoac`) supporting side-by-side systems on the same row
  * Automatic filtering of rotated margin watermarks and running page headers/footers
- Native Word Shapes / DrawingML (<wpg:wgp>, <wps:wsp>, <a:custGeom>) for TikZ & PGFPlots
- Native Editable Math Tables for all math table types (`tkz-tab`, `\choice`, `tabular`, `\immini`)
- Native Office Math (OMML `<m:oMath>`) for all inline & display equations
"""
from __future__ import annotations

import json
import math
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
from docx.oxml import OxmlElement, parse_xml
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
    """Normalize line endings to standard LF and strip illegal XML control characters."""
    s = sanitize_xml_text(s)
    return s.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")


# ==============================================================================
# COMPLETE 256-SLOT T5 (VNTEX) ENCODING TABLE FOR VIETNAMESE TeX FONTS
# Used when VNR12, VNBX12, VNSS12, VNTI12, VnNimbusRomanNo9L, etc. lack /ToUnicode
# ==============================================================================
T5_VNTEX_FALLBACK: dict[int, str] = {
    0: "`", 1: "´", 2: "ˆ", 3: "˜", 4: "¨", 5: "˙", 6: "˚", 7: "ˇ",
    8: "˘", 9: "¯", 10: "˝", 11: "¸", 12: "˛", 13: "‚", 14: "‹", 15: "›",
    16: "“", 17: "”", 18: "„", 19: "«", 20: "»", 21: "–", 22: "—", 23: "",
    24: "₀", 25: "ı", 26: "ȷ", 27: "ff", 28: "fi", 29: "fl", 30: "Đ", 31: "đ",
    32: " ", 33: "!", 34: "\"", 35: "#", 36: "$", 37: "%", 38: "&", 39: "’",
    40: "(", 41: ")", 42: "*", 43: "+", 44: ",", 45: "-", 46: ".", 47: "/",
    48: "0", 49: "1", 50: "2", 51: "3", 52: "4", 53: "5", 54: "6", 55: "7",
    56: "8", 57: "9", 58: ":", 59: ";", 60: "<", 61: "=", 62: ">", 63: "?",
    64: "@", 65: "A", 66: "B", 67: "C", 68: "D", 69: "E", 70: "F", 71: "G",
    72: "H", 73: "I", 74: "J", 75: "K", 76: "L", 77: "M", 78: "N", 79: "O",
    80: "P", 81: "Q", 82: "R", 83: "S", 84: "T", 85: "U", 86: "V", 87: "W",
    88: "X", 89: "Y", 90: "Z", 91: "[", 92: "\\", 93: "]", 94: "^", 95: "_",
    96: "‘", 97: "a", 98: "b", 99: "c", 100: "d", 101: "e", 102: "f", 103: "g",
    104: "h", 105: "i", 106: "j", 107: "k", 108: "l", 109: "m", 110: "n", 111: "o",
    112: "p", 113: "q", 114: "r", 115: "s", 116: "t", 117: "u", 118: "v", 119: "w",
    120: "x", 121: "y", 122: "z", 123: "{", 124: "|", 125: "}", 126: "~", 127: "-",
    128: "À", 129: "Á", 130: "Ã", 131: "Ả", 132: "Ạ", 133: "Â", 134: "Ầ", 135: "Ấ",
    136: "Ẫ", 137: "Ẩ", 138: "Ậ", 139: "Ă", 140: "Ằ", 141: "Ắ", 142: "Ẵ", 143: "Ẳ",
    144: "Ặ", 145: "È", 146: "É", 147: "Ẽ", 148: "Ẻ", 149: "Ẹ", 150: "Ê", 151: "Ề",
    152: "Ế", 153: "Ễ", 154: "Ể", 155: "Ệ", 156: "Ì", 157: "Í", 158: "Ĩ", 159: "Ỉ",
    160: "Ị", 161: "Ò", 162: "Ó", 163: "Õ", 164: "Ỏ", 165: "Ọ", 166: "Ô", 167: "Ồ",
    168: "Ố", 169: "Ỗ", 170: "Ổ", 171: "Ộ", 172: "Ơ", 173: "Ờ", 174: "Ớ", 175: "Ỡ",
    176: "Ở", 177: "Ợ", 178: "Ù", 179: "Ú", 180: "Ũ", 181: "Ủ", 182: "Ụ", 183: "Ư",
    184: "Ừ", 185: "Ứ", 186: "Ữ", 187: "Ử", 188: "Ự", 189: "Ỳ", 190: "Ý", 191: "Ỹ",
    192: "à", 193: "á", 194: "ã", 195: "ả", 196: "ạ", 197: "â", 198: "ầ", 199: "ấ",
    200: "ẫ", 201: "ẩ", 202: "ậ", 203: "ă", 204: "ằ", 205: "ắ", 206: "ẵ", 207: "ẳ",
    208: "ặ", 209: "è", 210: "é", 211: "ẽ", 212: "ẻ", 213: "ẹ", 214: "ê", 215: "ề",
    216: "ế", 217: "ễ", 218: "ể", 219: "ệ", 220: "ì", 221: "í", 222: "ĩ", 223: "ỉ",
    224: "ị", 225: "ò", 226: "ó", 227: "õ", 228: "ỏ", 229: "ọ", 230: "ô", 231: "ồ",
    232: "ố", 233: "ỗ", 234: "ổ", 235: "ộ", 236: "ơ", 237: "ờ", 238: "ớ", 239: "ỡ",
    240: "ở", 241: "ợ", 242: "ù", 243: "ú", 244: "ũ", 245: "ủ", 246: "ụ", 247: "ư",
    248: "ừ", 249: "ứ", 250: "ữ", 251: "ử", 252: "ự", 253: "ỳ", 254: "ý", 255: "ỹ",
}

# Add Yhookabove, yhookabove, Ydotbelow, ydotbelow at standard T5 positions if referenced via Differences
T5_GLYPH_MAP: dict[str, str] = {
    "Dcroat": "Đ", "dcroat": "đ",
    "Agrave": "À", "Aacute": "Á", "Atilde": "Ã", "Ahookabove": "Ả", "Adotbelow": "Ạ",
    "Acircumflex": "Â", "Acircumflexgrave": "Ầ", "Acircumflexacute": "Ấ",
    "Acircumflextilde": "Ẫ", "Acircumflexhookabove": "Ẩ", "Acircumflexdotbelow": "Ậ",
    "Abreve": "Ă", "Abrevegrave": "Ằ", "Abreveacute": "Ắ",
    "Abrevetilde": "Ẵ", "Abrevehookabove": "Ẳ", "Abrevedotbelow": "Ặ",
    "Egrave": "È", "Eacute": "É", "Etilde": "Ẽ", "Ehookabove": "Ẻ", "Edotbelow": "Ẹ",
    "Ecircumflex": "Ê", "Ecircumflexgrave": "Ề", "Ecircumflexacute": "Ế",
    "Ecircumflextilde": "Ễ", "Ecircumflexhookabove": "Ể", "Ecircumflexdotbelow": "Ệ",
    "Igrave": "Ì", "Iacute": "Í", "Itilde": "Ĩ", "Ihookabove": "Ỉ", "Idotbelow": "Ị",
    "Ograve": "Ò", "Oacute": "Ó", "Otilde": "Õ", "Ohookabove": "Ỏ", "Odotbelow": "Ọ",
    "Ocircumflex": "Ô", "Ocircumflexgrave": "Ồ", "Ocircumflexacute": "Ố",
    "Ocircumflextilde": "Ỗ", "Ocircumflexhookabove": "Ổ", "Ocircumflexdotbelow": "Ộ",
    "Ohorn": "Ơ", "Ohorngrave": "Ờ", "Ohornacute": "Ớ",
    "Ohorntilde": "Ỡ", "Ohornhookabove": "Ở", "Ohorndotbelow": "Ợ",
    "Ugrave": "Ù", "Uacute": "Ú", "Utilde": "Ũ", "Uhookabove": "Ủ", "Udotbelow": "Ụ",
    "Uhorn": "Ư", "Uhorngrave": "Ừ", "Uhornacute": "Ứ",
    "Uhorntilde": "Ữ", "Uhornhookabove": "Ử", "Uhorndotbelow": "Ự",
    "Ygrave": "Ỳ", "Yacute": "Ý", "Ytilde": "Ỹ", "Yhookabove": "Ỷ", "Ydotbelow": "Ỵ",
    "agrave": "à", "aacute": "á", "atilde": "ã", "ahookabove": "ả", "adotbelow": "ạ",
    "acircumflex": "â", "acircumflexgrave": "ầ", "acircumflexacute": "ấ",
    "acircumflextilde": "ẫ", "acircumflexhookabove": "ẩ", "acircumflexdotbelow": "ậ",
    "abreve": "ă", "abrevegrave": "ằ", "abreveacute": "ắ",
    "abrevetilde": "ẵ", "abrevehookabove": "ẳ", "abrevedotbelow": "ặ",
    "egrave": "è", "eacute": "é", "etilde": "ẽ", "ehookabove": "ẻ", "edotbelow": "ẹ",
    "ecircumflex": "ê", "ecircumflexgrave": "ề", "ecircumflexacute": "ế",
    "ecircumflextilde": "ễ", "ecircumflexhookabove": "ể", "ecircumflexdotbelow": "ệ",
    "igrave": "ì", "iacute": "í", "itilde": "ĩ", "ihookabove": "ỉ", "idotbelow": "ị",
    "ograve": "ò", "oacute": "ó", "otilde": "õ", "ohookabove": "ỏ", "odotbelow": "ọ",
    "ocircumflex": "ô", "ocircumflexgrave": "ồ", "ocircumflexacute": "ố",
    "ocircumflextilde": "ỗ", "ocircumflexhookabove": "ổ", "ocircumflexdotbelow": "ộ",
    "ohorn": "ơ", "ohorngrave": "ờ", "ohornacute": "ớ",
    "ohorntilde": "ỡ", "ohornhookabove": "ở", "ohorndotbelow": "ợ",
    "ugrave": "ù", "uacute": "ú", "utilde": "ũ", "uhookabove": "ủ", "udotbelow": "ụ",
    "uhorn": "ư", "uhorngrave": "ừ", "uhornacute": "ứ",
    "uhorntilde": "ữ", "uhornhookabove": "ử", "uhorndotbelow": "ự",
    "ygrave": "ỳ", "yacute": "ý", "ytilde": "ỹ", "yhookabove": "ỷ", "ydotbelow": "ỵ",
}

GLYPH_TO_TEXT: dict[str, str] = {
    **T5_GLYPH_MAP,
    # Standard punctuation & digits
    "space": " ", "exclam": "!", "quotedbl": "\"", "numbersign": "#",
    "dollar": "$", "percent": "%", "ampersand": "&", "quotesingle": "'",
    "quoteright": "'", "quoteleft": "`", "parenleft": "(", "parenright": ")",
    "asterisk": "*", "plus": "+", "comma": ",", "hyphen": "-", "minus": "-",
    "period": ".", "slash": "/", "colon": ":", "semicolon": ";",
    "less": "<", "equal": "=", "greater": ">", "question": "?", "at": "@",
    "bracketleft": "[", "backslash": r"\setminus ", "bracketright": "]",
    "asciicircum": "^", "underscore": "_", "braceleft": r"\{", "braceright": r"\}",
    "bar": "|", "brokenbar": "|", "asciitilde": "~",
    "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
    "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
    "endash": "–", "emdash": "—", "bullet": "•", "ellipsis": "…",
    "fi": "fi", "fl": "fl", "ff": "ff", "ffi": "ffi", "ffl": "ffl",
    "dotlessi": "ı", "dotlessj": "ȷ",
    # Greek letters & math symbols
    "alpha": r"\alpha ", "beta": r"\beta ", "gamma": r"\gamma ", "delta": r"\delta ",
    "epsilon": r"\epsilon ", "epsilon1": r"\varepsilon ", "varepsilon": r"\varepsilon ",
    "zeta": r"\zeta ", "eta": r"\eta ", "theta": r"\theta ", "vartheta": r"\vartheta ",
    "iota": r"\iota ", "kappa": r"\kappa ", "lambda": r"\lambda ", "mu": r"\mu ",
    "nu": r"\nu ", "xi": r"\xi ", "pi": r"\pi ", "varpi": r"\varpi ",
    "rho": r"\rho ", "varrho": r"\varrho ", "sigma": r"\sigma ", "varsigma": r"\varsigma ",
    "tau": r"\tau ", "upsilon": r"\upsilon ", "phi": r"\phi ", "phi1": r"\varphi ",
    "varphi": r"\varphi ", "chi": r"\chi ", "psi": r"\psi ", "omega": r"\omega ",
    "Gamma": r"\Gamma ", "Delta": r"\Delta ", "Theta": r"\Theta ", "Lambda": r"\Lambda ",
    "Xi": r"\Xi ", "Pi": r"\Pi ", "Sigma": r"\Sigma ", "Upsilon": r"\Upsilon ",
    "Phi": r"\Phi ", "Psi": r"\Psi ", "Omega": r"\Omega ",
    "infinity": r"\infty ", "infty": r"\infty ",
    "plusminus": r"\pm ", "minusplus": r"\mp ",
    "multiply": r"\times ", "times": r"\times ",
    "periodcentered": r"\cdot ", " mathasterisk": "*", "asteriskmath": "*",
    "divide": r"\div ", "lessequal": r"\le ", "greaterequal": r"\ge ",
    "notequal": r"\neq ", "equivalence": r"\equiv ", "approxequal": r"\approx ",
    "similar": r"\sim ", "similarequal": r"\simeq ", "proportional": r"\propto ",
    "element": r"\in ", "notelement": r"\notin ", "owner": r"\ni ",
    "propersubset": r"\subset ", "reflexsubset": r"\subseteq ",
    "propersuperset": r"\supset ", "reflexsuperset": r"\supseteq ",
    "intersection": r"\cap ", "union": r"\cup ", "emptyset": r"\varnothing ",
    "universal": r"\forall ", "existential": r"\exists ",
    "logicaland": r"\land ", "logicalor": r"\lor ", "logicalnot": r"\neg ",
    "arrowright": r"\to ", "arrowleft": r"\leftarrow ", "arrowboth": r"\leftrightarrow ",
    "arrowdblright": r"\Rightarrow ", "arrowdblleft": r"\Leftarrow ",
    "arrowdblboth": r"\Leftrightarrow ", "mapsto": r"\mapsto ",
    "arrowup": r"\uparrow ", "arrowdown": r"\downarrow ",
    "partialdiff": r"\partial ", "gradient": r"\nabla ", " nabla": r"\nabla ",
    "prime": "'", "minute": "'", "second": "''",
    "radical": r"\sqrt", "radicalbig": r"\sqrt", "radicalBig": r"\sqrt",
    "radicalbigg": r"\sqrt", "radicalBigg": r"\sqrt", "radicalbt": r"\sqrt",
    "radicaltp": "", "radicalvertex": "",
    "summationtext": r"\sum ", "summationdisplay": r"\sum ",
    "producttext": r"\prod ", "productdisplay": r"\prod ",
    "integraltext": r"\int ", "integraldisplay": r"\int ",
    "uniontext": r"\bigcup ", "uniondisplay": r"\bigcup ",
    "intersectiontext": r"\bigcap ", "intersectiondisplay": r"\bigcap ",
    "vextendsingle": "|", "vextenddouble": r"\Vert ", "bardbl": r"\Vert ",
    # Single-line scalable delimiters
    "parenleftbig": "(", "parenrightbig": ")",
    "parenleftBig": "(", "parenrightBig": ")",
    "parenleftbigg": "(", "parenrightbigg": ")",
    "parenleftBigg": "(", "parenrightBigg": ")",
    "bracketleftbig": "[", "bracketrightbig": "]",
    "bracketleftBig": "[", "bracketrightBig": "]",
    "bracketleftbigg": "[", "bracketrightbigg": "]",
    "braceleftbig": r"\{", "bracerightbig": r"\}",
    "braceleftBig": r"\{", "bracerightBig": r"\}",
    "braceleftbigg": r"\{", "bracerightbigg": r"\}",
    # Multi-line system delimiters
    "bracketleftBigg": "@@SYS_BRACKET@@",
    "bracketlefttp": "@@SYS_BRACKET@@",
    "bracketleftbt": "",
    "bracketleftex": "",
    "bracketrightBigg": "]",
    "bracketrighttp": "", "bracketrightbt": "", "bracketrightex": "",
    "braceleftBigg": "@@SYS_BRACE@@",
    "bracelefttp": "@@SYS_BRACE@@",
    "braceleftmid": "", "braceleftbt": "", "braceex": "",
    "bracerightBigg": r"\}",
    "bracerighttp": "", "bracerightmid": "", "bracerightbt": "",
    "parenlefttp": "(", "parenleftbt": "", "parenleftex": "",
    "parenrighttp": ")", "parenrightbt": "", "parenrightex": "",
}

CMSY_FALLBACK: dict[int, str] = {
    0: "-", 1: r"\cdot ", 2: r"\times ", 3: "*", 4: r"\div ", 5: r"\diamond ",
    6: r"\pm ", 7: r"\mp ", 8: r"\oplus ", 9: r"\ominus ", 10: r"\otimes ",
    11: r"\oslash ", 12: r"\odot ", 13: r"\circ ", 14: r"\circ ", 15: r"\bullet ",
    16: r"\asymp ", 17: r"\equiv ", 18: r"\subseteq ", 19: r"\supseteq ",
    20: r"\le ", 21: r"\ge ", 22: r"\preceq ", 23: r"\succeq ",
    24: r"\sim ", 25: r"\approx ", 26: r"\subset ", 27: r"\supset ",
    28: r"\ll ", 29: r"\gg ", 30: r"\prec ", 31: r"\succ ",
    32: r"\leftarrow ", 33: r"\to ", 34: r"\uparrow ", 35: r"\downarrow ",
    36: r"\leftrightarrow ", 37: r"\nearrow ", 38: r"\searrow ", 39: r"\simeq ",
    40: r"\Leftarrow ", 41: r"\Rightarrow ", 42: r"\Uparrow ", 43: r"\Downarrow ",
    44: r"\Leftrightarrow ", 45: r"\nwarrow ", 46: r"\swarrow ", 47: r"\propto ",
    48: "'", 49: r"\infty ", 50: r"\in ", 51: r"\ni ",
    52: r"\triangle ", 53: r"\bigtriangledown ", 54: "/", 55: "",
    56: r"\forall ", 57: r"\exists ", 58: r"\neg ", 59: r"\emptyset ",
    60: r"\Re ", 61: r"\Im ", 62: r"\top ", 63: r"\perp ",
    91: r"\cup ", 92: r"\cap ", 94: r"\land ", 95: r"\lor ",
    102: r"\{", 103: r"\}", 104: r"\langle ", 105: r"\rangle ",
    106: "|", 107: r"\Vert ", 110: r"\setminus ", 112: r"\sqrt",
}

CMEX_FALLBACK: dict[int, str] = {
    0: "(", 1: ")", 2: "[", 3: "]",
    8: r"\{", 9: r"\}", 12: "|", 13: r"\Vert ",
    16: "(", 17: ")", 18: "(", 19: ")",
    20: "[", 21: "]", 26: r"\{", 27: r"\}",
    32: "(", 33: ")", 34: "@@SYS_BRACKET@@", 35: "]",
    40: "@@SYS_BRACE@@", 41: r"\}",
    48: "(", 49: ")", 50: "@@SYS_BRACKET@@", 51: "",
    52: "", 53: "", 54: "", 55: "",
    56: "@@SYS_BRACE@@", 57: "", 58: "", 59: "",
    60: "", 61: "", 62: "", 63: "",
    64: "", 65: "", 66: "",
    80: r"\sum ", 81: r"\prod ", 82: r"\int ", 83: r"\bigcup ", 84: r"\bigcap ",
    88: r"\sum ", 89: r"\prod ", 90: r"\int ", 91: r"\bigcup ", 92: r"\bigcap ",
    98: r"\widehat", 99: r"\widehat", 100: r"\widetilde", 101: r"\widetilde",
    104: "[", 105: "]", 110: r"\{", 111: r"\}",
    112: r"\sqrt", 113: r"\sqrt", 114: r"\sqrt", 115: r"\sqrt", 116: r"\sqrt",
    117: "", 118: "",
    # Yhcmex extended slots used by Vietnamese LaTeX exam packages (ex_test / dethi)
    128: "(", 129: "(", 138: ")", 139: ")",
    150: "@@SYS_BRACKET@@",
    167: r"\{", 168: "@@SYS_BRACE@@", 170: r"\}",
}

CMMI_FALLBACK: dict[int, str] = {
    0: r"\Gamma ", 1: r"\Delta ", 2: r"\Theta ", 3: r"\Lambda ", 4: r"\Xi ",
    5: r"\Pi ", 6: r"\Sigma ", 7: r"\Upsilon ", 8: r"\Phi ", 9: r"\Psi ", 10: r"\Omega ",
    11: r"\alpha ", 12: r"\beta ", 13: r"\gamma ", 14: r"\delta ", 15: r"\epsilon ",
    16: r"\zeta ", 17: r"\eta ", 18: r"\theta ", 19: r"\iota ", 20: r"\kappa ",
    21: r"\lambda ", 22: r"\mu ", 23: r"\nu ", 24: r"\xi ", 25: r"\pi ",
    26: r"\rho ", 27: r"\sigma ", 28: r"\tau ", 29: r"\upsilon ", 30: r"\phi ",
    31: r"\chi ", 32: r"\psi ", 33: r"\omega ", 34: r"\varepsilon ",
    35: r"\vartheta ", 36: r"\varpi ", 37: r"\varrho ", 38: r"\varsigma ", 39: r"\varphi ",
    58: ".", 59: ",", 60: "<", 61: "/", 62: ">", 64: r"\partial ", 96: r"\ell ",
}

CMR_FALLBACK: dict[int, str] = {
    0: r"\Gamma ", 1: r"\Delta ", 2: r"\Theta ", 3: r"\Lambda ", 4: r"\Xi ",
    5: r"\Pi ", 6: r"\Sigma ", 7: r"\Upsilon ", 8: r"\Phi ", 9: r"\Psi ", 10: r"\Omega ",
    11: "ff", 12: "fi", 13: "fl", 14: "ffi", 15: "ffl",
    34: "”", 39: "’", 40: "(", 41: ")", 43: "+", 44: ",", 45: "-", 46: ".", 47: "/",
    58: ":", 59: ";", 61: "=", 91: "[", 93: "]", 96: "“", 123: "–", 124: "—",
}

MSBM_FALLBACK: dict[int, str] = {
    64: r"\varnothing ", 65: r"\mathbb{A}", 66: r"\mathbb{B}", 67: r"\mathbb{C}",
    68: r"\mathbb{D}", 69: r"\mathbb{E}", 70: r"\mathbb{F}", 72: r"\mathbb{H}",
    73: r"\mathbb{I}", 75: r"\mathbb{K}", 78: r"\mathbb{N}", 80: r"\mathbb{P}",
    81: r"\mathbb{Q}", 82: r"\mathbb{R}", 83: r"\mathbb{S}", 90: r"\mathbb{Z}",
}


@dataclass
class PdfSpan2D:
    x: float
    y: float
    end_x: float
    font_size: float
    font_name: str
    text: str
    is_math_font: bool = False
    is_italic: bool = False
    is_bold: bool = False
    underscript: str = ""


@dataclass
class PdfLineSeg2D:
    x1: float
    y1: float
    x2: float
    y2: float


@dataclass
class PdfRect2D:
    x: float
    y: float
    w: float
    h: float


def _mat_mul(m1: list[float], m2: list[float]) -> list[float]:
    """Multiply two 2D affine matrices [a, b, c, d, e, f]."""
    return [
        m1[0] * m2[0] + m1[1] * m2[2],
        m1[0] * m2[1] + m1[1] * m2[3],
        m1[2] * m2[0] + m1[3] * m2[2],
        m1[2] * m2[1] + m1[3] * m2[3],
        m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
        m1[4] * m2[1] + m1[5] * m2[3] + m2[5],
    ]


def _parse_tounicode_cmap(cmap_text: str) -> dict[int, str]:
    """Parse a PDF /ToUnicode CMap stream into a code -> Unicode string dictionary."""
    mapping: dict[int, str] = {}

    def _decode_hex_str(hex_s: str) -> str:
        hs = re.sub(r"\s+", "", hex_s)
        if not hs or len(hs) % 2 != 0:
            return ""
        try:
            raw = bytes.fromhex(hs)
            if len(raw) >= 2 and len(raw) % 2 == 0:
                s = raw.decode("utf-16-be", errors="ignore")
            else:
                s = raw.decode("latin1", errors="ignore")
            # Strip null bytes & control chars that could corrupt XML
            return sanitize_xml_text(s)
        except Exception:
            return ""

    for m_bf in re.finditer(r"beginbfchar(.*?)endbfchar", cmap_text, re.S):
        block = m_bf.group(1)
        for m_pair in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            code = int(m_pair.group(1), 16)
            val = _decode_hex_str(m_pair.group(2))
            if val:
                mapping[code] = val

    for m_rng in re.finditer(r"beginbfrange(.*?)endbfrange", cmap_text, re.S):
        block = m_rng.group(1)
        for line in block.splitlines():
            line = line.strip()
            if not line:
                continue
            m_arr = re.match(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*)\]", line)
            if m_arr:
                start_c = int(m_arr.group(1), 16)
                end_c = int(m_arr.group(2), 16)
                hex_items = re.findall(r"<([0-9A-Fa-f]+)>", m_arr.group(3))
                for idx, h_item in enumerate(hex_items):
                    if start_c + idx <= end_c:
                        val = _decode_hex_str(h_item)
                        if val:
                            mapping[start_c + idx] = val
                continue
            m_seq = re.match(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", line)
            if m_seq:
                start_c = int(m_seq.group(1), 16)
                end_c = int(m_seq.group(2), 16)
                base_u = int(m_seq.group(3), 16)
                for offset in range(end_c - start_c + 1):
                    cp = base_u + offset
                    if 0x20 <= cp <= 0x10FFFF and not (0xD800 <= cp <= 0xDFFF):
                        mapping[start_c + offset] = sanitize_xml_text(chr(cp))

    return mapping


def _parse_type1_fontfile_encoding(raw_font_bytes: bytes) -> dict[int, str]:
    """Parse `dup <code> /<glyph_name> put` from an embedded Type1 `/FontFile` cleartext header."""
    out: dict[int, str] = {}
    if not raw_font_bytes:
        return out
    head_str = raw_font_bytes[:20000].decode("latin1", errors="ignore")
    for m in re.finditer(r"\bdup\s+(\d+)\s+/([A-Za-z0-9_.\-]+)\s+put\b", head_str):
        code = int(m.group(1))
        gname = m.group(2)
        if gname == ".notdef":
            continue
        if gname in GLYPH_TO_TEXT:
            out[code] = GLYPH_TO_TEXT[gname]
        elif len(gname) == 1 and gname.isprintable():
            out[code] = gname
        elif gname.startswith("uni") and len(gname) == 7:
            try:
                cp = int(gname[3:], 16)
                if cp >= 0x20:
                    out[code] = sanitize_xml_text(chr(cp))
            except Exception:
                pass
    return out


def _build_font_decoder(font_obj: Any) -> tuple[str, dict[int, str], bool, bool, bool]:
    """Build `(base_font_name, char_map, is_math_font, is_italic, is_bold)` for a PDF Font dictionary."""
    base_font = str(font_obj.get("/BaseFont", "")) if hasattr(font_obj, "get") else ""
    clean_name = re.sub(r"^/[A-Z]{6}\+", "/", base_font).lstrip("/")
    up_name = clean_name.upper()

    is_math = any(
        k in up_name
        for k in ("CMMI", "CMSY", "CMEX", "YHCMEX", "MSAM", "MSBM", "EUFM", "RSFS", "LINE10", "LCIRCLE")
    )
    is_italic = ("ITALIC" in up_name) or ("OBLIQUE" in up_name) or up_name.startswith("VNTI") or up_name.startswith("CMTI") or up_name.startswith("CMMI")
    is_bold = ("BOLD" in up_name) or up_name.startswith("VNBX") or up_name.startswith("CMBX")

    char_map: dict[int, str] = {}

    # 1. Pre-populate font-family fallback table first
    if "YHCMEX" in up_name or "CMEX" in up_name:
        char_map.update(CMEX_FALLBACK)
    elif "CMSY" in up_name:
        char_map.update(CMSY_FALLBACK)
    elif "CMMI" in up_name:
        for i in range(65, 91):
            char_map[i] = chr(i)
        for i in range(97, 123):
            char_map[i] = chr(i)
        for i in range(48, 58):
            char_map[i] = chr(i)
        char_map.update(CMMI_FALLBACK)
    elif "MSBM" in up_name or "MSAM" in up_name:
        char_map.update(MSBM_FALLBACK)
    elif up_name.startswith(("VN", "UTM")) or "VNNIMBUS" in up_name:
        char_map.update(T5_VNTEX_FALLBACK)
    elif up_name.startswith(("CMR", "CMBX", "CMTI", "CMSS", "CMTT")):
        for i in range(32, 127):
            char_map[i] = chr(i)
        char_map.update(CMR_FALLBACK)
    else:
        for i in range(32, 127):
            char_map[i] = chr(i)

    # 2. Check embedded Type1 /FontDescriptor -> /FontFile cleartext encoding table
    try:
        if "/FontDescriptor" in font_obj:
            fd = font_obj["/FontDescriptor"].get_object()
            for ff_key in ("/FontFile", "/FontFile2", "/FontFile3"):
                if ff_key in fd:
                    ff_data = fd[ff_key].get_object().get_data()
                    t1_map = _parse_type1_fontfile_encoding(ff_data)
                    if t1_map:
                        char_map.update(t1_map)
    except Exception:
        pass

    # 3. Check /Encoding -> /Differences array (overrides base font slots)
    try:
        if "/Encoding" in font_obj:
            enc = font_obj["/Encoding"].get_object()
            if hasattr(enc, "get") and "/Differences" in enc:
                diffs = enc["/Differences"]
                cur_code = 0
                for item in diffs:
                    if isinstance(item, int):
                        cur_code = int(item)
                    else:
                        gname = str(item).lstrip("/")
                        if gname in GLYPH_TO_TEXT:
                            char_map[cur_code] = GLYPH_TO_TEXT[gname]
                        elif len(gname) == 1 and gname.isprintable():
                            char_map[cur_code] = gname
                        cur_code += 1
    except Exception:
        pass

    # 4. Check /ToUnicode CMap (only for non-CMEX/CMSY fonts, or where ToUnicode maps to valid printable chars)
    if not ("CMEX" in up_name or "YHCMEX" in up_name or "CMSY" in up_name):
        try:
            if "/ToUnicode" in font_obj:
                tu_data = font_obj["/ToUnicode"].get_object().get_data().decode("latin1", errors="ignore")
                tu_map = _parse_tounicode_cmap(tu_data)
                for k, v in tu_map.items():
                    if v and all(ord(c) >= 0x20 for c in v):
                        # Normalize common math/TeX ligatures from ToUnicode
                        if v == "\u2212":
                            v = "-"
                        char_map[k] = v
        except Exception:
            pass

    return clean_name, char_map, is_math, is_italic, is_bold


def _decode_pdf_bytes_with_font(raw_bytes: bytes, char_map: dict[int, str]) -> str:
    """Decode raw PDF string bytes using the font's character map."""
    out: list[str] = []
    for b in raw_bytes:
        if b in char_map:
            out.append(char_map[b])
        elif 32 <= b <= 126:
            out.append(chr(b))
    return sanitize_xml_text("".join(out))


def _unescape_pdf_literal_bytes(s: str) -> bytes:
    """Convert contents of a PDF literal string `(...)` into raw bytes, handling octal `\\ddd` and escapes."""
    out = bytearray()
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "\\" and i + 1 < n:
            nxt = s[i + 1]
            if nxt in "01234567":
                oct_digits = nxt
                i += 2
                for _ in range(2):
                    if i < n and s[i] in "01234567":
                        oct_digits += s[i]
                        i += 1
                    else:
                        break
                out.append(int(oct_digits, 8) & 0xFF)
                continue
            elif nxt == "n":
                out.append(10)
            elif nxt == "r":
                out.append(13)
            elif nxt == "t":
                out.append(9)
            elif nxt == "b":
                out.append(8)
            elif nxt == "f":
                out.append(12)
            elif nxt in "\r\n":
                i += 2
                if i < n and nxt == "\r" and s[i - 1] == "\n":
                    i += 1
                continue
            else:
                out.append(ord(nxt) & 0xFF)
            i += 2
        else:
            out.append(ord(ch) & 0xFF)
            i += 1
    return bytes(out)


def _tokenize_pdf_tj_array(arr_str: str) -> list[tuple[str, Any]]:
    """Parse inside `[...]` of a `TJ` operator into `[('str', bytes), ('kern', float), ...]`."""
    items: list[tuple[str, Any]] = []
    i = 0
    n = len(arr_str)
    while i < n:
        while i < n and arr_str[i] in " \t\r\n":
            i += 1
        if i >= n:
            break
        ch = arr_str[i]
        if ch == "(":
            i += 1
            depth = 1
            buf: list[str] = []
            while i < n and depth > 0:
                if arr_str[i] == "\\" and i + 1 < n:
                    buf.append(arr_str[i : i + 2])
                    i += 2
                    continue
                if arr_str[i] == "(":
                    depth += 1
                    buf.append("(")
                elif arr_str[i] == ")":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                    buf.append(")")
                else:
                    buf.append(arr_str[i])
                i += 1
            items.append(("str", _unescape_pdf_literal_bytes("".join(buf))))
        elif ch == "<":
            end_h = arr_str.find(">", i + 1)
            if end_h == -1:
                break
            hs = re.sub(r"\s+", "", arr_str[i + 1 : end_h])
            if len(hs) % 2 == 1:
                hs += "0"
            try:
                items.append(("str", bytes.fromhex(hs)))
            except Exception:
                pass
            i = end_h + 1
        else:
            m_num = re.match(r"^[+-]?(?:\d+\.?\d*|\.\d+)", arr_str[i:])
            if m_num:
                items.append(("kern", float(m_num.group(0))))
                i += m_num.end()
            else:
                i += 1
    return items


def _parse_page_content_stream(
    stream_text: str,
    fonts_info: dict[str, tuple[str, dict[int, str], bool, bool, bool]],
) -> tuple[list[PdfSpan2D], list[PdfLineSeg2D], list[PdfRect2D]]:
    """Execute a 2D state-machine over a PDF page content stream, tracking CTM, BT/ET, Tf, Tm, Td, Tj, TJ, m, l, re."""
    spans: list[PdfSpan2D] = []
    lines: list[PdfLineSeg2D] = []
    rects: list[PdfRect2D] = []

    ctm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    ctm_stack: list[list[float]] = []

    in_text = False
    tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    tlm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    leading = 0.0

    # CRITICAL FIX: Persist cur_font_key and cur_font_size across BT ... ET blocks!
    cur_font_key = ""
    cur_font_size = 10.0

    path_pts: list[tuple[float, float]] = []
    cur_subpath: list[tuple[float, float]] = []

    # Tokenize content stream preserving strings (...), hex <...>, arrays [...], and operators
    token_re = re.compile(
        r"(\((?:\\.|[^\\()]|\((?:\\.|[^\\()])*\))*\)|<[0-9A-Fa-f\s]*>|\[(?:\((?:\\.|[^\\()]|\((?:\\.|[^\\()])*\))*\)|<[0-9A-Fa-f\s]*>|[^\[\]()<>])*\]|/[A-Za-z0-9_.\-+]+|[+-]?(?:\d+\.?\d*|\.\d+)|[A-Za-z*'\"01]+)",
        re.S,
    )
    operands: list[str] = []

    def _emit_text_bytes(raw_b: bytes) -> None:
        nonlocal tm
        if not raw_b:
            return
        f_name, c_map, is_math, is_ital, is_bold = fonts_info.get(
            cur_font_key, ("CMR10", {i: chr(i) for i in range(32, 127)}, False, False, False)
        )
        decoded = _decode_pdf_bytes_with_font(raw_b, c_map)
        if not decoded:
            return

        # Compute page coordinates via M = tm x ctm
        m_full = _mat_mul(tm, ctm)
        # Discard rotated vertical margin watermarks (e.g., 90-degree rotated text on page edges)
        if abs(m_full[1]) > 0.25 or abs(m_full[2]) > 0.25:
            return

        px = m_full[4]
        py = m_full[5]
        eff_fs = abs(cur_font_size * math.hypot(m_full[0], m_full[1]))
        if eff_fs < 1.0:
            eff_fs = abs(cur_font_size)

        if ("cmex" in f_name.lower() or "yhcmex" in f_name.lower()) and decoded.strip() in ("[", "{"):
            decoded = "@@SYS_BRACKET@@" if decoded.strip() == "[" else "@@SYS_BRACE@@"
        char_w = eff_fs * (0.44 if not is_math else 0.48)
        adv_user = len(raw_b) * cur_font_size * 0.44
        end_px = px + len(raw_b) * char_w

        spans.append(
            PdfSpan2D(
                x=px,
                y=py,
                end_x=end_px,
                font_size=eff_fs,
                font_name=f_name,
                text=decoded,
                is_math_font=is_math,
                is_italic=is_ital,
                is_bold=is_bold,
            )
        )
        tm[4] += adv_user * tm[0]

    for m_tok in token_re.finditer(stream_text):
        tok = m_tok.group(0)
        first = tok[0]
        if first in "(<[/+-0123456789.":
            operands.append(tok)
            continue

        op = tok
        if op == "q":
            ctm_stack.append(list(ctm))
            operands.clear()
        elif op == "Q":
            if ctm_stack:
                ctm = ctm_stack.pop()
            operands.clear()
        elif op == "cm":
            if len(operands) >= 6:
                try:
                    vals = [float(x) for x in operands[-6:]]
                    ctm = _mat_mul(vals, ctm)
                except Exception:
                    pass
            operands.clear()
        elif op == "BT":
            in_text = True
            tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
            tlm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
            operands.clear()
        elif op == "ET":
            in_text = False
            operands.clear()
        elif op == "Tf":
            if len(operands) >= 2:
                cur_font_key = operands[-2]
                try:
                    cur_font_size = float(operands[-1])
                except Exception:
                    pass
            operands.clear()
        elif op == "TL":
            if operands:
                try:
                    leading = float(operands[-1])
                except Exception:
                    pass
            operands.clear()
        elif op in ("Td", "TD"):
            if len(operands) >= 2:
                try:
                    tx = float(operands[-2])
                    ty = float(operands[-1])
                    if op == "TD":
                        leading = -ty
                    tlm = _mat_mul([1.0, 0.0, 0.0, 1.0, tx, ty], tlm)
                    tm = list(tlm)
                except Exception:
                    pass
            operands.clear()
        elif op == "Tm":
            if len(operands) >= 6:
                try:
                    tm = [float(x) for x in operands[-6:]]
                    tlm = list(tm)
                except Exception:
                    pass
            operands.clear()
        elif op == "T*":
            tlm = _mat_mul([1.0, 0.0, 0.0, 1.0, 0.0, -leading], tlm)
            tm = list(tlm)
            operands.clear()
        elif op in ("Tj", "'", '"'):
            if op in ("'", '"'):
                tlm = _mat_mul([1.0, 0.0, 0.0, 1.0, 0.0, -leading], tlm)
                tm = list(tlm)
            if operands:
                s_op = operands[-1]
                if s_op.startswith("(") and s_op.endswith(")"):
                    _emit_text_bytes(_unescape_pdf_literal_bytes(s_op[1:-1]))
                elif s_op.startswith("<") and s_op.endswith(">"):
                    hs = re.sub(r"\s+", "", s_op[1:-1])
                    if len(hs) % 2 == 1:
                        hs += "0"
                    try:
                        _emit_text_bytes(bytes.fromhex(hs))
                    except Exception:
                        pass
            operands.clear()
        elif op == "TJ":
            if operands:
                arr_op = operands[-1]
                if arr_op.startswith("[") and arr_op.endswith("]"):
                    tj_items = _tokenize_pdf_tj_array(arr_op[1:-1])
                    buf_b = bytearray()
                    intra_kern = 0.0
                    for kind, val in tj_items:
                        if kind == "str":
                            buf_b.extend(val)
                        elif kind == "kern":
                            if val <= -140:
                                if buf_b:
                                    _emit_text_bytes(bytes(buf_b))
                                    buf_b.clear()
                                tm[4] -= ((val + intra_kern) / 1000.0) * cur_font_size * tm[0]
                                intra_kern = 0.0
                            else:
                                intra_kern += val
                    if buf_b:
                        _emit_text_bytes(bytes(buf_b))
                    if intra_kern != 0.0:
                        tm[4] -= (intra_kern / 1000.0) * cur_font_size * tm[0]
            operands.clear()
        elif not in_text:
            if op == "m" and len(operands) >= 2:
                try:
                    ux, uy = float(operands[-2]), float(operands[-1])
                    px = ux * ctm[0] + uy * ctm[2] + ctm[4]
                    py = ux * ctm[1] + uy * ctm[3] + ctm[5]
                    cur_subpath = [(px, py)]
                except Exception:
                    pass
                operands.clear()
            elif op == "l" and len(operands) >= 2:
                try:
                    ux, uy = float(operands[-2]), float(operands[-1])
                    px = ux * ctm[0] + uy * ctm[2] + ctm[4]
                    py = ux * ctm[1] + uy * ctm[3] + ctm[5]
                    if cur_subpath:
                        p0 = cur_subpath[-1]
                        lines.append(PdfLineSeg2D(p0[0], p0[1], px, py))
                    cur_subpath.append((px, py))
                except Exception:
                    pass
                operands.clear()
            elif op == "re" and len(operands) >= 4:
                try:
                    rx, ry, rw, rh = [float(x) for x in operands[-4:]]
                    px1 = rx * ctm[0] + ry * ctm[2] + ctm[4]
                    py1 = rx * ctm[1] + ry * ctm[3] + ctm[5]
                    px2 = (rx + rw) * ctm[0] + (ry + rh) * ctm[2] + ctm[4]
                    py2 = (rx + rw) * ctm[1] + (ry + rh) * ctm[3] + ctm[5]
                    bx, by = min(px1, px2), min(py1, py2)
                    bw, bh = abs(px2 - px1), abs(py2 - py1)
                    rects.append(PdfRect2D(bx, by, bw, bh))
                    if bh < 1.8 and bw > 4.0:
                        lines.append(PdfLineSeg2D(bx, by + bh / 2.0, bx + bw, by + bh / 2.0))
                    elif bw < 1.8 and bh > 4.0:
                        lines.append(PdfLineSeg2D(bx + bw / 2.0, by, bx + bw / 2.0, by + bh))
                except Exception:
                    pass
                operands.clear()
            elif op in ("S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"):
                cur_subpath = []
                operands.clear()
            else:
                operands.clear()
        else:
            operands.clear()

    return spans, lines, rects


def _extract_balanced_pdf_delim(s: str, open_ch: str, close_ch: str, start_idx: int = 0) -> str:
    """Extract a balanced `<< ... >>` or `[ ... ]` substring from a PDF object string."""
    idx = s.find(open_ch, start_idx)
    if idx == -1:
        return ""
    depth = 0
    i = idx
    n = len(s)
    step = len(open_ch)
    while i < n:
        if s.startswith(open_ch, i):
            depth += 1
            i += step
        elif s.startswith(close_ch, i):
            depth -= 1
            i += step
            if depth == 0:
                return s[idx:i]
        else:
            i += 1
    return s[idx:]


def extract_pdf_pages_2d(
    pdf_path: Path,
) -> list[tuple[list[PdfSpan2D], list[PdfLineSeg2D], list[PdfRect2D]]]:
    """Extract ordered pages of `(spans, lines, rects)` from any PDF file using 100% pure Python (`zlib` + `re`)."""
    pdf_bytes = Path(pdf_path).read_bytes()
    raw_objs: dict[int, tuple[str, bytes | None]] = {}

    for m in re.finditer(rb"(\d+)\s+(\d+)\s+obj\b(.*?)endobj", pdf_bytes, re.S):
        onum = int(m.group(1))
        body = m.group(3)
        m_stream = re.search(rb"<<(.*?)>>\s*stream", body, re.S)
        if m_stream:
            dict_bytes = b"<<" + m_stream.group(1) + b">>"
            stream_start = m_stream.end()
            if stream_start < len(body) and body[stream_start : stream_start + 2] == b"\r\n":
                stream_start += 2
            elif stream_start < len(body) and body[stream_start : stream_start + 1] in (b"\n", b"\r"):
                stream_start += 1
            stream_end = body.rfind(b"endstream")
            if stream_end == -1:
                stream_end = len(body)
            raw_stream = body[stream_start:stream_end]
            raw_objs[onum] = (dict_bytes.decode("latin1", errors="ignore"), raw_stream)
        else:
            raw_objs[onum] = (body.decode("latin1", errors="ignore").strip(), None)

    def _get_int_val(onum: int) -> int | None:
        s = raw_objs.get(onum, ("", None))[0].strip()
        m_i = re.match(r"^[+-]?\d+$", s)
        return int(m_i.group(0)) if m_i else None

    stream_cache: dict[int, bytes] = {}

    def _get_stream_data(onum: int) -> bytes:
        if onum in stream_cache:
            return stream_cache[onum]
        if onum not in raw_objs:
            return b""
        d_str, r_bytes = raw_objs[onum]
        if r_bytes is None:
            return b""
        length_val = None
        m_len_ref = re.search(r"/Length\s+(\d+)\s+\d+\s+R", d_str)
        if m_len_ref:
            length_val = _get_int_val(int(m_len_ref.group(1)))
        else:
            m_len_dir = re.search(r"/Length\s+(\d+)\b", d_str)
            if m_len_dir:
                length_val = int(m_len_dir.group(1))

        candidates: list[bytes] = []
        if length_val is not None and 0 < length_val <= len(r_bytes):
            candidates.append(r_bytes[:length_val])
        candidates.append(r_bytes)
        if r_bytes.endswith(b"\r\n"):
            candidates.append(r_bytes[:-2])
        elif r_bytes.endswith((b"\n", b"\r")):
            candidates.append(r_bytes[:-1])
        candidates.append(r_bytes.rstrip(b"\r\n "))

        if "/FlateDecode" in d_str or "/Fl" in d_str:
            for cand in candidates:
                for wbits in (zlib.MAX_WBITS, -zlib.MAX_WBITS):
                    try:
                        dec = zlib.decompress(cand, wbits)
                        stream_cache[onum] = dec
                        return dec
                    except Exception:
                        try:
                            dobj = zlib.decompressobj(wbits)
                            dec = dobj.decompress(cand)
                            if dec:
                                stream_cache[onum] = dec
                                return dec
                        except Exception:
                            pass
        res = candidates[0] if candidates else r_bytes
        stream_cache[onum] = res
        return res

    # Unpack PDF 1.5+ /ObjStm compressed object streams
    objstm_nums = [
        onum
        for onum, (d_str, r_bytes) in list(raw_objs.items())
        if r_bytes is not None and "/ObjStm" in d_str
    ]
    for onum in objstm_nums:
        d_str, _ = raw_objs[onum]
        m_n = re.search(r"/N\s+(\d+)", d_str)
        m_first = re.search(r"/First\s+(\d+)", d_str)
        if not m_n or not m_first:
            continue
        n_objs = int(m_n.group(1))
        first_off = int(m_first.group(1))
        dec = _get_stream_data(onum)
        if not dec or first_off > len(dec):
            continue
        header_str = dec[:first_off].decode("latin1", errors="ignore")
        nums = [int(x) for x in re.findall(r"\d+", header_str)]
        pairs: list[tuple[int, int]] = []
        for i in range(0, min(len(nums) - 1, n_objs * 2), 2):
            pairs.append((nums[i], nums[i + 1]))
        for idx, (sub_onum, rel_off) in enumerate(pairs):
            start_pos = first_off + rel_off
            end_pos = (first_off + pairs[idx + 1][1]) if (idx + 1 < len(pairs)) else len(dec)
            sub_str = dec[start_pos:end_pos].decode("latin1", errors="ignore").strip()
            if sub_onum not in raw_objs:
                raw_objs[sub_onum] = (sub_str, None)

    def _get_obj_str(onum: int) -> str:
        return raw_objs.get(onum, ("", None))[0]

    def _resolve_subdict_or_ref(parent_str: str, key: str) -> str:
        idx = parent_str.find(key)
        if idx == -1:
            return ""
        rest = parent_str[idx + len(key) :].lstrip()
        m_ref = re.match(r"^(\d+)\s+\d+\s+R\b", rest)
        if m_ref:
            return _get_obj_str(int(m_ref.group(1)))
        if rest.startswith("<<"):
            return _extract_balanced_pdf_delim(rest, "<<", ">>", 0)
        if rest.startswith("["):
            return _extract_balanced_pdf_delim(rest, "[", "]", 0)
        return ""

    font_decoder_cache: dict[int, tuple[str, dict[int, str], bool, bool, bool]] = {}

    def _build_font_decoder_pure(
        f_onum: int, f_inline_str: str = ""
    ) -> tuple[str, dict[int, str], bool, bool, bool]:
        if f_onum > 0 and f_onum in font_decoder_cache:
            return font_decoder_cache[f_onum]
        f_str = _get_obj_str(f_onum) if f_onum > 0 else f_inline_str
        m_bf = re.search(r"/BaseFont\s*/([A-Za-z0-9_.\-+]+)", f_str)
        base_font = m_bf.group(1) if m_bf else ""
        clean_name = re.sub(r"^[A-Z]{6}\+", "", base_font).lstrip("/")
        up_name = clean_name.upper()

        is_math = any(
            k in up_name
            for k in ("CMMI", "CMSY", "CMEX", "YHCMEX", "MSAM", "MSBM", "EUFM", "RSFS", "LINE10", "LCIRCLE")
        )
        is_italic = (
            ("ITALIC" in up_name)
            or ("OBLIQUE" in up_name)
            or up_name.startswith(("VNTI", "CMTI", "CMMI"))
        )
        is_bold = ("BOLD" in up_name) or up_name.startswith(("VNBX", "CMBX"))

        char_map: dict[int, str] = {}
        if "YHCMEX" in up_name or "CMEX" in up_name:
            char_map.update(CMEX_FALLBACK)
        elif "CMSY" in up_name:
            char_map.update(CMSY_FALLBACK)
        elif "CMMI" in up_name:
            for i in range(65, 91):
                char_map[i] = chr(i)
            for i in range(97, 123):
                char_map[i] = chr(i)
            for i in range(48, 58):
                char_map[i] = chr(i)
            char_map.update(CMMI_FALLBACK)
        elif "MSBM" in up_name or "MSAM" in up_name:
            char_map.update(MSBM_FALLBACK)
        elif up_name.startswith(("VN", "UTM")) or "VNNIMBUS" in up_name:
            char_map.update(T5_VNTEX_FALLBACK)
        elif up_name.startswith(("CMR", "CMBX", "CMTI", "CMSS", "CMTT")):
            for i in range(32, 127):
                char_map[i] = chr(i)
            char_map.update(CMR_FALLBACK)
        else:
            for i in range(32, 127):
                char_map[i] = chr(i)

        # 2. Check embedded Type1 /FontDescriptor -> /FontFile cleartext encoding table
        fd_str = _resolve_subdict_or_ref(f_str, "/FontDescriptor")
        if fd_str:
            for ff_key in ("/FontFile", "/FontFile2", "/FontFile3"):
                m_ff = re.search(ff_key + r"\b\s*(\d+)\s+\d+\s+R", fd_str)
                if m_ff:
                    ff_data = _get_stream_data(int(m_ff.group(1)))
                    t1_map = _parse_type1_fontfile_encoding(ff_data)
                    if t1_map:
                        char_map.update(t1_map)

        # 3. Check /Encoding -> /Differences array
        enc_str = _resolve_subdict_or_ref(f_str, "/Encoding")
        if enc_str and "/Differences" in enc_str:
            diff_arr = _resolve_subdict_or_ref(enc_str, "/Differences")
            if diff_arr.startswith("[") and diff_arr.endswith("]"):
                inner = diff_arr[1:-1]
                cur_code = 0
                for tok in re.findall(r"/[A-Za-z0-9_.\-+]+|[+-]?\d+", inner):
                    if tok.startswith("/"):
                        gname = tok[1:]
                        if gname in GLYPH_TO_TEXT:
                            char_map[cur_code] = GLYPH_TO_TEXT[gname]
                        elif len(gname) == 1 and gname.isprintable():
                            char_map[cur_code] = gname
                        cur_code += 1
                    else:
                        cur_code = int(tok)

        # 4. Check /ToUnicode CMap
        if not ("CMEX" in up_name or "YHCMEX" in up_name or "CMSY" in up_name):
            m_tu = re.search(r"/ToUnicode\s+(\d+)\s+\d+\s+R", f_str)
            if m_tu:
                tu_data = _get_stream_data(int(m_tu.group(1))).decode("latin1", errors="ignore")
                tu_map = _parse_tounicode_cmap(tu_data)
                for k, v in tu_map.items():
                    if v and all(ord(c) >= 0x20 for c in v):
                        if v == "\u2212":
                            v = "-"
                        char_map[k] = v

        res = (clean_name, char_map, is_math, is_italic, is_bold)
        if f_onum > 0:
            font_decoder_cache[f_onum] = res
        return res

    # Traverse /Catalog -> /Pages -> /Kids tree in exact page order
    ordered_pages: list[tuple[int, str]] = []
    root_pages_onum: int | None = None
    for onum, (d_str, _) in raw_objs.items():
        if re.search(r"/Type\s*/Catalog\b", d_str):
            m_p = re.search(r"/Pages\s+(\d+)\s+\d+\s+R", d_str)
            if m_p:
                root_pages_onum = int(m_p.group(1))
                break
    if root_pages_onum is None:
        m_root = re.search(rb"/Root\s+(\d+)\s+\d+\s+R", pdf_bytes)
        if m_root:
            cat_str = _get_obj_str(int(m_root.group(1)))
            m_p = re.search(r"/Pages\s+(\d+)\s+\d+\s+R", cat_str)
            if m_p:
                root_pages_onum = int(m_p.group(1))

    visited_nodes: set[int] = set()

    def _walk_pages_node(node_onum: int, inherited_res: str = "") -> None:
        if node_onum in visited_nodes or node_onum not in raw_objs:
            return
        visited_nodes.add(node_onum)
        n_str = _get_obj_str(node_onum)
        own_res = _resolve_subdict_or_ref(n_str, "/Resources")
        eff_res = own_res if own_res else inherited_res
        if re.search(r"/Type\s*/Pages\b", n_str) or "/Kids" in n_str:
            kids_arr = _resolve_subdict_or_ref(n_str, "/Kids")
            for m_k in re.finditer(r"(\d+)\s+\d+\s+R", kids_arr):
                _walk_pages_node(int(m_k.group(1)), eff_res)
        elif re.search(r"/Type\s*/Page\b", n_str):
            ordered_pages.append((node_onum, eff_res))

    if root_pages_onum is not None:
        _walk_pages_node(root_pages_onum, "")
    if not ordered_pages:
        for onum in sorted(raw_objs.keys()):
            d_str = _get_obj_str(onum)
            if re.search(r"/Type\s*/Page\b", d_str):
                ordered_pages.append((onum, _resolve_subdict_or_ref(d_str, "/Resources")))

    pages_out: list[tuple[list[PdfSpan2D], list[PdfLineSeg2D], list[PdfRect2D]]] = []
    for p_onum, eff_res in ordered_pages:
        p_str = _get_obj_str(p_onum)
        fonts_info: dict[str, tuple[str, dict[int, str], bool, bool, bool]] = {}
        if eff_res:
            font_dict_str = _resolve_subdict_or_ref(eff_res, "/Font")
            if font_dict_str:
                for m_f in re.finditer(r"(/[A-Za-z0-9_.\-+]+)\s+(\d+)\s+\d+\s+R", font_dict_str):
                    f_key = m_f.group(1)
                    f_onum = int(m_f.group(2))
                    fonts_info[f_key] = _build_font_decoder_pure(f_onum)

        stream_chunks: list[str] = []
        idx_c = p_str.find("/Contents")
        if idx_c != -1:
            rest_c = p_str[idx_c + len("/Contents") :].lstrip()
            if rest_c.startswith("["):
                arr_c = _extract_balanced_pdf_delim(rest_c, "[", "]", 0)
                for m_cr in re.finditer(r"(\d+)\s+\d+\s+R", arr_c):
                    c_data = _get_stream_data(int(m_cr.group(1)))
                    if c_data:
                        stream_chunks.append(c_data.decode("latin1", errors="ignore"))
            else:
                m_cr = re.match(r"^(\d+)\s+\d+\s+R", rest_c)
                if m_cr:
                    c_onum = int(m_cr.group(1))
                    c_data = _get_stream_data(c_onum)
                    if c_data:
                        stream_chunks.append(c_data.decode("latin1", errors="ignore"))
                    else:
                        # Indirect reference to an array of content stream references
                        c_arr_str = _get_obj_str(c_onum).strip()
                        if c_arr_str.startswith("["):
                            for m_sub in re.finditer(r"(\d+)\s+\d+\s+R", c_arr_str):
                                sub_data = _get_stream_data(int(m_sub.group(1)))
                                if sub_data:
                                    stream_chunks.append(sub_data.decode("latin1", errors="ignore"))

        full_stream = "\n".join(stream_chunks)
        spans, lines, rects = _parse_page_content_stream(full_stream, fonts_info)
        pages_out.append((spans, lines, rects))

    return pages_out


# ==============================================================================
# 2D SPATIAL MATH & VIETNAMESE RECONSTRUCTION ENGINE (PDF -> CLEAN LATEX)
# ==============================================================================
VN_ACCENTED_CHARS = set(
    "àáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ"
    "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ"
)


def _deduplicate_and_filter_page_spans(spans: list[PdfSpan2D]) -> list[PdfSpan2D]:
    """Remove running headers/footers, duplicate overlaid TikZ nodes, and stacked vertical bar segments."""
    if not spans:
        return []

    ys = [sp.y for sp in spans]
    min_y, max_y = min(ys), max(ys)

    filtered: list[PdfSpan2D] = []
    seen_exact: set[tuple[int, int, str]] = set()
    vert_bars: list[tuple[float, float]] = []

    for sp in spans:
        txt = sp.text
        if not txt or not txt.strip():
            continue

        # Filter top running header (e.g., "42. MAX, MIN CỦA HÀM TRỊ TUYỆT ĐỐI..." or "TRỊ TUYỆT ĐỐI 2 TPOC...")
        if sp.y > 768 and (max_y - min_y) > 120:
            continue
        if sp.y > 745 and any(k in sp.text for k in ("PHÁT TRIỂN ĐỀ MINH HỌA", "TPOC LUYỆN THI", "TRỊ TUYỆT ĐỐI 2")):
            continue
        # Filter bottom running footer (e.g., "Geogebra Pro", "Trang 526")
        if sp.y < 52 and (max_y - min_y) > 120:
            continue
        if "Geogebra Pro" in txt or "TPOC LUYỆN THI" in txt:
            continue

        # Deduplicate stacked vertical bar extension glyphs (`vextendsingle` / `|`) at the same X coordinate
        if txt.strip() == "|":
            if any(abs(sp.x - bx) < 1.8 and abs(sp.y - by) < 22.0 for bx, by in vert_bars):
                continue
            vert_bars.append((sp.x, sp.y))

        # Deduplicate identical text drawn twice at the exact same (x, y) (e.g. TikZ BBT nodes)
        key = (int(round(sp.x * 2)), int(round(sp.y * 2)), txt.strip())
        if key in seen_exact:
            continue
        seen_exact.add(key)
        filtered.append(sp)

    # Merge adjacent characters on the exact same baseline and font when kerning split a word/number
    filtered.sort(key=lambda s: (-round(s.y, 1), s.x))
    merged: list[PdfSpan2D] = []
    for sp in filtered:
        if not merged:
            merged.append(sp)
            continue
        prev = merged[-1]
        same_y = abs(sp.y - prev.y) < 0.8
        same_font = (sp.font_name == prev.font_name) and abs(sp.font_size - prev.font_size) < 0.6
        gap = sp.x - prev.end_x
        if (
            same_y
            and same_font
            and -2.2 <= gap <= 0.25
            and not prev.text.startswith("@@")
            and not sp.text.startswith("@@")
            and prev.text.strip() != "|"
            and sp.text.strip() != "|"
        ):
            prev.text += sp.text
            prev.end_x = max(prev.end_x, sp.end_x)
        else:
            merged.append(sp)

    return merged


def _attach_underscripts_to_operators(spans: list[PdfSpan2D]) -> list[PdfSpan2D]:
    """Attach underscripts like `[0;3]`, `[a;b]`, `[-1;2]`, `x \in [0;2]` directly below `max`, `min`, `lim`."""
    op_words = ("max", "min", "lim", "sup", "inf")
    op_indices: list[int] = []
    for idx, sp in enumerate(spans):
        t_clean = sp.text.strip()
        if t_clean in op_words or any(t_clean.endswith(w) for w in op_words):
            op_indices.append(idx)

    if not op_indices:
        return spans

    consumed: set[int] = set()
    for op_idx in op_indices:
        op_sp = spans[op_idx]
        # Find where the operator word starts horizontally
        t_clean = op_sp.text.strip()
        op_x0 = op_sp.x
        op_x1 = op_sp.end_x
        for w in op_words:
            pos = op_sp.text.rfind(w)
            if pos != -1 and len(op_sp.text) > len(w):
                ratio = pos / max(len(op_sp.text), 1)
                op_x0 = op_sp.x + (op_sp.end_x - op_sp.x) * ratio

        sub_cands: list[tuple[int, PdfSpan2D]] = []
        for j, cand in enumerate(spans):
            if j == op_idx or j in consumed:
                continue
            dy = op_sp.y - cand.y
            if 2.2 <= dy <= 11.5 and cand.font_size < op_sp.font_size * 0.92:
                if (op_x0 - 12.0) <= cand.x <= (op_x1 + 16.0):
                    sub_cands.append((j, cand))

        if sub_cands:
            sub_cands.sort(key=lambda item: item[1].x)
            sub_str = "".join(c.text.strip() for _, c in sub_cands)
            op_sp.underscript = sub_str
            for j, _ in sub_cands:
                consumed.add(j)

    return [sp for idx, sp in enumerate(spans) if idx not in consumed]


def _detect_and_extract_answer_grid(
    spans: list[PdfSpan2D],
) -> tuple[str | None, list[PdfSpan2D]]:
    """Detect a 10-column answer key table (`BẢNG ĐÁP ÁN`: `1. B  2. C ...`) and render as a clean LaTeX tabular."""
    ans_re = re.compile(r"^(\d{1,2})\.\s*([A-D])$")
    ans_items: list[tuple[int, str, int]] = []
    for idx, sp in enumerate(spans):
        m = ans_re.match(sp.text.strip())
        if m:
            ans_items.append((int(m.group(1)), m.group(2), idx))

    # Also check pairs of adjacent spans where one is `1.` and the next is `B`
    if len(ans_items) < 8:
        used_idx: set[int] = set()
        ans_items.clear()
        for idx, sp in enumerate(spans):
            if idx in used_idx:
                continue
            t = sp.text.strip()
            m_both = ans_re.match(t)
            if m_both:
                ans_items.append((int(m_both.group(1)), m_both.group(2), idx))
                used_idx.add(idx)
                continue
            m_num = re.match(r"^(\d{1,2})\.$", t)
            if m_num:
                # Look for single letter A-D to the right on the same Y
                for j, sp2 in enumerate(spans):
                    if j != idx and j not in used_idx and abs(sp2.y - sp.y) < 3.0 and 0 < (sp2.x - sp.x) < 42.0:
                        if sp2.text.strip() in ("A", "B", "C", "D"):
                            ans_items.append((int(m_num.group(1)), sp2.text.strip(), idx))
                            used_idx.add(idx)
                            used_idx.add(j)
                            break

    if len(ans_items) < 10:
        return None, spans

    consumed_indices: set[int] = set()
    min_y = min(spans[i].y for _, _, i in ans_items) - 8.0
    max_y = max(spans[i].y for _, _, i in ans_items) + 8.0
    for idx, sp in enumerate(spans):
        if min_y <= sp.y <= max_y:
            consumed_indices.add(idx)

    ans_items.sort(key=lambda x: x[0])
    cells = [f"\\textbf{{{num}.}} {let}" for num, let, _ in ans_items]
    cols = 10
    rows_tex: list[str] = []
    for r in range(0, len(cells), cols):
        chunk = cells[r : r + cols]
        while len(chunk) < cols:
            chunk.append("")
        rows_tex.append(" & ".join(chunk) + r" \\ \hline")

    col_spec = "|" + "|".join(["c"] * cols) + "|"
    table_tex = (
        "\\begin{tabular}{" + col_spec + "}\n\\hline\n"
        + "\n".join(rows_tex)
        + "\n\\end{tabular}"
    )
    remaining = [sp for idx, sp in enumerate(spans) if idx not in consumed_indices]
    return table_tex, remaining


def _detect_and_extract_bbt_tables(
    spans: list[PdfSpan2D],
    lines: list[PdfLineSeg2D],
) -> tuple[list[tuple[float, str]], list[PdfSpan2D]]:
    """Detect Variation Tables (`Bảng biến thiên` with `x`, `g'(x)`/`y'`/`f'(x)`, `g(x)`/`y`/`f(x)`) and convert to `tkz-tab`."""
    horiz_lines = [
        ln for ln in lines if abs(ln.y1 - ln.y2) < 1.5 and abs(ln.x2 - ln.x1) > 100.0
    ]
    if len(horiz_lines) < 3:
        return [], spans

    horiz_lines.sort(key=lambda l: -l.y1)
    # Cluster horizontal lines that belong to the same table box
    groups: list[list[PdfLineSeg2D]] = []
    for ln in horiz_lines:
        if not groups or abs(groups[-1][-1].y1 - ln.y1) > 95.0:
            groups.append([ln])
        else:
            groups[-1].append(ln)

    extracted_tables: list[tuple[float, str]] = []
    consumed_ids: set[int] = set()

    for grp in groups:
        if len(grp) < 3:
            continue
        top_y = max(l.y1 for l in grp) + 4.0
        bot_y = min(l.y1 for l in grp) - 4.0
        left_x = min(min(l.x1, l.x2) for l in grp) - 6.0
        right_x = max(max(l.x1, l.x2) for l in grp) + 6.0

        box_spans = [
            (idx, sp)
            for idx, sp in enumerate(spans)
            if idx not in consumed_ids and left_x <= sp.x <= right_x and bot_y <= sp.y <= top_y
        ]
        if len(box_spans) < 5:
            continue

        # Sort distinct Y levels of the horizontal lines: top, row1_sep, row2_sep, bottom
        y_levels = sorted({round(l.y1, 1) for l in grp}, reverse=True)
        if len(y_levels) < 3:
            continue

        # Find separator vertical line between header column and values
        sep_x = left_x + 42.0
        for ln in lines:
            if abs(ln.x1 - ln.x2) < 1.5 and abs(ln.y2 - ln.y1) > 30.0:
                if left_x + 15.0 <= ln.x1 <= left_x + 85.0 and bot_y <= min(ln.y1, ln.y2) and max(ln.y1, ln.y2) <= top_y:
                    sep_x = ln.x1
                    break

        r0_top, r0_bot = y_levels[0] + 4.0, y_levels[1] - 2.0
        r1_top, r1_bot = y_levels[1] + 2.0, (y_levels[2] - 2.0 if len(y_levels) >= 4 else bot_y)
        r2_top, r2_bot = (y_levels[2] + 2.0 if len(y_levels) >= 4 else r1_bot), bot_y

        row0_val_spans = [sp for _, sp in box_spans if sp.x > sep_x and r0_bot <= sp.y <= r0_top]
        row1_val_spans = [sp for _, sp in box_spans if sp.x > sep_x and r1_bot <= sp.y <= r1_top]
        row2_val_spans = [sp for _, sp in box_spans if sp.x > sep_x and r2_bot <= sp.y < r1_bot]

        # Cluster row0 X-values horizontally
        row0_val_spans.sort(key=lambda s: s.x)
        x_clusters: list[list[PdfSpan2D]] = []
        for sp in row0_val_spans:
            if not x_clusters or (sp.x - x_clusters[-1][-1].end_x) > 14.0:
                x_clusters.append([sp])
            else:
                x_clusters[-1].append(sp)

        if len(x_clusters) < 2:
            continue

        x_vals: list[str] = []
        x_centers: list[float] = []
        for cl in x_clusters:
            x_centers.append(sum(s.x for s in cl) / len(cl))
            txt_cl = _render_math_fragment_from_spans(cl).strip()
            x_vals.append(f"${txt_cl}$" if txt_cl else "$0$")

        # Header labels
        h_r1 = "$g'(x)$"
        h_r2 = "$g(x)$"
        for _, sp in box_spans:
            if sp.x <= sep_x:
                if "f" in sp.text:
                    h_r1, h_r2 = "$f'(x)$", "$f(x)$"
                elif "y" in sp.text:
                    h_r1, h_r2 = "$y'$", "$y$"

        # Build tkzTabLine tokens (2*N - 1 slots)
        num_pts = len(x_vals)
        line_slots = [""] * (2 * num_pts - 1)
        for sp in row1_val_spans:
            tok = sp.text.strip().replace("−", "-")
            if tok not in ("+", "-", "0"):
                continue
            # Find closest point or interval
            best_slot = 0
            best_dist = 1e9
            for p_i in range(num_pts):
                d = abs(sp.x - x_centers[p_i])
                if d < best_dist:
                    best_dist = d
                    best_slot = 2 * p_i
                if p_i + 1 < num_pts:
                    mid_x = 0.5 * (x_centers[p_i] + x_centers[p_i + 1])
                    dm = abs(sp.x - mid_x)
                    if dm < best_dist:
                        best_dist = dm
                        best_slot = 2 * p_i + 1
            line_slots[best_slot] = tok

        # Cluster row2 extrema values around each x_center
        var_entries: list[str] = []
        mid_y_r2 = 0.5 * (r2_top + r2_bot)
        for p_i, xc in enumerate(x_centers):
            near = [
                sp for sp in row2_val_spans
                if abs(sp.x - xc) < 28.0
            ]
            if near:
                avg_y = sum(s.y for s in near) / len(near)
                pos_prefix = "+/" if avg_y >= mid_y_r2 else "-/"
                val_ltx = _render_math_fragment_from_spans(near).strip()
                var_entries.append(f"{pos_prefix} ${val_ltx}$")
            else:
                pos_prefix = "+/" if (p_i % 2 == 1) else "-/"
                var_entries.append(f"{pos_prefix} $a$")

        tikz_bbt = (
            "\\begin{tikzpicture}\n"
            f"\\tkzTabInit[lgt=1.3,espcl=1.8]{{$x$ / 0.9, {h_r1} / 0.8, {h_r2} / 1.8}}{{{', '.join(x_vals)}}}\n"
            f"\\tkzTabLine{{{', '.join(line_slots)}}}\n"
            f"\\tkzTabVar{{{', '.join(var_entries)}}}\n"
            "\\end{tikzpicture}"
        )
        extracted_tables.append(((top_y + bot_y) / 2.0, tikz_bbt))
        for idx, _ in box_spans:
            consumed_ids.add(idx)

    remaining = [sp for idx, sp in enumerate(spans) if idx not in consumed_ids]
    return extracted_tables, remaining


def _render_math_fragment_from_spans(
    spans: list[PdfSpan2D],
    lines: list[PdfLineSeg2D] | None = None,
) -> str:
    """Convert a 2D cluster of math spans and fraction/radical horizontal bars into clean LaTeX math."""
    if not spans:
        return ""

    # Sort primarily left-to-right
    s_list = sorted(spans, key=lambda s: (round(s.x, 1), -round(s.y, 1)))

    # 1. Check if there is a horizontal fraction bar or long arrow bar spanning across part of `s_list`
    if lines:
        min_x = min(s.x for s in s_list) - 3.0
        max_x = max(s.end_x for s in s_list) + 3.0
        min_y = min(s.y for s in s_list) - 6.0
        max_y = max(s.y for s in s_list) + 12.0

        cand_bars = [
            ln
            for ln in lines
            if abs(ln.y1 - ln.y2) < 1.2
            and 6.0 <= abs(ln.x2 - ln.x1) <= (max_x - min_x + 15.0)
            and min_x <= min(ln.x1, ln.x2)
            and max(ln.x1, ln.x2) <= max_x
            and min_y <= ln.y1 <= max_y
        ]
        # Sort bars left to right
        cand_bars.sort(key=lambda l: min(l.x1, l.x2))
        for bar in cand_bars:
            bx1, bx2 = min(bar.x1, bar.x2), max(bar.x1, bar.x2)
            by = bar.y1
            above = [
                s for s in s_list
                if (s.end_x > bx1 - 2.0 and s.x < bx2 + 2.0) and (0.5 <= (s.y - by) <= 18.0)
            ]
            below = [
                s for s in s_list
                if (s.end_x > bx1 - 2.0 and s.x < bx2 + 2.0) and (1.5 <= (by - s.y) <= 18.0)
            ]
            if above and below:
                above_ids = {id(s) for s in above}
                below_ids = {id(s) for s in below}
                left_s = [s for s in s_list if id(s) not in above_ids and id(s) not in below_ids and s.x < bx1]
                right_s = [s for s in s_list if id(s) not in above_ids and id(s) not in below_ids and s.x >= bx1]
                remaining_lines = [l for l in lines if l is not bar]

                left_ltx = _render_math_fragment_from_spans(left_s, remaining_lines)
                num_ltx = _render_math_fragment_from_spans(above, remaining_lines)
                den_ltx = _render_math_fragment_from_spans(below, remaining_lines)
                right_ltx = _render_math_fragment_from_spans(right_s, remaining_lines)

                # Check if this bar is actually a long arrow `\xrightarrow[den]{num}` (when right_s starts with \to or -)
                if any(r"\to" in s.text or "→" in s.text for s in above + below + right_s[:2]):
                    pass
                return f"{left_ltx}\\frac{{{num_ltx}}}{{{den_ltx}}}{right_ltx}"

    # 2. Check for `\sqrt` radical symbol (`\sqrt` in s.text) and wrap its radicand in `{...}`
    for idx, sp in enumerate(s_list):
        if r"\sqrt" in sp.text:
            # Find horizontal overbar starting near sp.end_x
            rad_end_x = sp.end_x + 28.0
            if lines:
                for ln in lines:
                    if abs(ln.y1 - ln.y2) < 1.2 and abs( min(ln.x1, ln.x2) - sp.x ) < 16.0 and ln.y1 > sp.y:
                        rad_end_x = max(ln.x1, ln.x2) + 2.0
                        break
            inside = [
                s for j, s in enumerate(s_list)
                if j > idx and s.x < rad_end_x and abs(s.y - sp.y) < 8.0
            ]
            if inside:
                in_ids = {id(s) for s in inside}
                before = s_list[:idx]
                after = [s for j, s in enumerate(s_list) if j > idx and id(s) not in in_ids]
                prefix = sp.text.replace(r"\sqrt", "")
                b_ltx = _render_math_fragment_from_spans(before, lines)
                in_ltx = _render_math_fragment_from_spans(inside, lines)
                a_ltx = _render_math_fragment_from_spans(after, lines)
                return f"{b_ltx}{prefix}\\sqrt{{{in_ltx}}}{a_ltx}"

    # 3. Linear pass with superscript (`dy > 1.8`) and subscript (`dy < -1.8`) detection
    # Find dominant baseline Y
    main_spans = [s for s in s_list if s.font_size >= 8.5] or s_list
    base_y = sorted(s.y for s in main_spans)[len(main_spans) // 2]
    base_fs = max((s.font_size for s in main_spans), default=10.0)

    out_tokens: list[str] = []
    i = 0
    n = len(s_list)
    while i < n:
        sp = s_list[i]
        raw_t = sp.text.replace("@@SYS_BRACKET@@", "[").replace("@@SYS_BRACE@@", r"\{")

        # Handle operator with attached underscript (e.g. `max` with `[0;3]`)
        if sp.underscript:
            u_clean = sp.underscript
            for op_w in ("max", "min", "lim", "sup", "inf"):
                if op_w in raw_t:
                    raw_t = raw_t.replace(op_w, f"\\{op_w}_{{{u_clean}}} ", 1)
                    break

        dy = sp.y - base_y
        is_script_size = sp.font_size < base_fs * 0.86

        if out_tokens and (abs(dy) > 1.9 or is_script_size) and abs(dy) > 1.2:
            # Group consecutive script tokens at the same script Y
            script_group = [sp]
            j = i + 1
            while j < n:
                nxt = s_list[j]
                if abs(nxt.y - sp.y) < 1.8 and (nxt.x - script_group[-1].end_x) < 8.0:
                    script_group.append(nxt)
                    j += 1
                else:
                    break
            scr_str = "".join(
                s.text.replace("@@SYS_BRACKET@@", "").replace("@@SYS_BRACE@@", "") for s in script_group
            ).strip()
            if scr_str:
                if dy > 0:
                    if scr_str in ("'", "′"):
                        out_tokens.append("'")
                    else:
                        out_tokens.append(f"^{{{scr_str}}}")
                else:
                    out_tokens.append(f"_{{{scr_str}}}")
            i = j
            continue

        # Fix negated `\in` + `/` -> `\notin`
        if raw_t.strip() == "/" and out_tokens and out_tokens[-1].strip().endswith(r"\in"):
            out_tokens[-1] = re.sub(r"\\in\s*$", r"\\notin ", out_tokens[-1])
            i += 1
            continue
        if raw_t.strip() == r"\in" and i + 1 < n and s_list[i + 1].text.strip() == "/" and abs(s_list[i + 1].x - sp.x) < 8.0:
            out_tokens.append(r"\notin ")
            i += 2
            continue
        # Fix negated `=` + `/` -> `\neq`
        if raw_t.strip() == "/" and out_tokens and out_tokens[-1].strip().endswith("="):
            out_tokens[-1] = out_tokens[-1].rstrip()[:-1] + r"\neq "
            i += 1
            continue
        if raw_t.strip() == "=" and i + 1 < n and s_list[i + 1].text.strip() == "/" and abs(s_list[i + 1].x - sp.x) < 8.0:
            out_tokens.append(r"\neq ")
            i += 2
            continue

        out_tokens.append(raw_t)
        i += 1

    res = "".join(out_tokens)
    for op_w in ("max", "min", "lim", "sup", "inf", "sin", "cos", "tan", "cot", "ln", "log"):
        res = re.sub(rf"(?<!\\)\b{op_w}\b", rf"\\{op_w} ", res)
    res = res.replace("∈/", r"\notin ").replace("/∈", r"\notin ").replace("=/=", r"\neq ")
    return sanitize_xml_text(res.strip())


def _format_mixed_vietnamese_and_math_line(
    line_spans: list[PdfSpan2D],
    page_lines: list[PdfLineSeg2D],
) -> str:
    """Convert a horizontal line of spans into mixed Vietnamese text and `$ ... $` inline math."""
    if not line_spans:
        return ""

    line_spans.sort(key=lambda s: s.x)

    # Check if there is a long arrow `----->` (`\xrightarrow[below]{above}`) on this line
    # e.g., `|K+m| <= h|k+m| --(|K+m|>=|k+m| / K+m cùng dấu k+m)--> m \in S_1`
    out_parts: list[tuple[str, list[PdfSpan2D]]] = []

    def _is_span_math(sp: PdfSpan2D) -> bool:
        t = sp.text.strip()
        if not t:
            return False
        if any(c in VN_ACCENTED_CHARS for c in t):
            return False
        if sp.is_math_font:
            return True
        if sp.underscript:
            return True
        if any(k in t for k in ("\\", "^", "_", "{", "}", "|", "=", "<", ">", "≤", "≥", "≠", "∈", "∉", "⊂", "∪", "∩", "⇒", "⇔", "→", "∀", "∃", "∞", "±")):
            return True
        if re.match(r"^[0-9+\-*/()\[\];,.]+$", t):
            return True
        if sp.is_italic and len(t) <= 3 and t.isalpha():
            return True
        if t in ("max", "min", "lim", "sin", "cos", "tan", "ln", "log", "f", "g", "u", "y", "x", "m", "a", "b", "c", "k", "K", "M", "S", "T", "P", "A", "B", "C", "D"):
            return True
        return False

    # Segment `line_spans` into alternating ('text', [...]) and ('math', [...]) runs
    for sp in line_spans:
        kind = "math" if _is_span_math(sp) else "text"
        # Single capital letter A, B, C, D at start of option or after "phương án" is text if followed by period/space
        if not out_parts or out_parts[-1][0] != kind:
            out_parts.append((kind, [sp]))
        else:
            out_parts[-1][1].append(sp)

    # Merge `[text]` bridges that are just spaces/punctuation between two `math` segments
    merged_parts: list[tuple[str, list[PdfSpan2D]]] = []
    for kind, grp in out_parts:
        if (
            kind == "text"
            and merged_parts
            and merged_parts[-1][0] == "math"
            and all(s.text.strip() in ("", ",", ";", ".", ":", "(", ")", "[", "]") for s in grp)
        ):
            merged_parts[-1][1].extend(grp)
        else:
            merged_parts.append((kind, grp))

    final_str_parts: list[str] = []
    prev_end_x: float | None = None

    for kind, grp in merged_parts:
        if not grp:
            continue
        if prev_end_x is not None and (grp[0].x - prev_end_x) > 0.70:
            final_str_parts.append(" ")
        prev_end_x = grp[-1].end_x

        if kind == "text":
            t_pieces: list[str] = []
            p_end: float | None = None
            for s in grp:
                if p_end is not None and (s.x - p_end) > 0.65 and not (t_pieces and t_pieces[-1].endswith(" ")):
                    t_pieces.append(" ")
                t_pieces.append(s.text.replace("@@SYS_BRACKET@@", "[").replace("@@SYS_BRACE@@", "{"))
                p_end = s.end_x
            txt_joined = "".join(t_pieces)
            if all(s.is_bold for s in grp if s.text.strip()) and len(txt_joined.strip()) > 1:
                final_str_parts.append(f"\\textbf{{{txt_joined.strip()}}}")
            else:
                final_str_parts.append(txt_joined)
        else:
            m_ltx = _render_math_fragment_from_spans(grp, page_lines).strip()
            if not m_ltx:
                continue
            # If the math fragment is just a standalone punctuation mark like `.` or `,` or `(1)`, keep it plain
            if m_ltx in (".", ",", ":", ";", "?", "!"):
                final_str_parts.append(m_ltx)
            else:
                final_str_parts.append(f"${m_ltx}$")

    line_out = "".join(final_str_parts).strip()
    # Clean up adjacent `$a$$b$` -> `$a b$`
    line_out = re.sub(r"\$\s*\$", " ", line_out)
    # Clean up long dashed arrows `− − − − − →` into `\xrightarrow{...}`
    line_out = re.sub(r"(?:-\s*){3,}\\to", r"\\longrightarrow ", line_out)
    return sanitize_xml_text(line_out)


def _extract_multiline_systems_and_lines(
    spans: list[PdfSpan2D],
    page_lines: list[PdfLineSeg2D],
) -> list[str]:
    """Reconstruct a full page of 2D spans into ordered LaTeX paragraphs, display systems (`\heva`, `\hoac`), and `\choice` macros."""
    if not spans:
        return []

    # 1. Identify multi-line system delimiters (`@@SYS_BRACKET@@` or `@@SYS_BRACE@@`)
    sys_delims = [
        sp for sp in spans if ("@@SYS_BRACKET@@" in sp.text or "@@SYS_BRACE@@" in sp.text)
    ]

    # Cluster spans into visual rows by Y coordinate, EXCEPT spans that belong inside a multi-line system
    # First, pair each system delimiter with its vertical span and horizontal X bound
    sys_boxes: list[tuple[float, float, float, float, str, PdfSpan2D]] = []
    sys_delims.sort(key=lambda s: (-round(s.y, 1), s.x))

    used_delim_ids: set[int] = set()
    for d_sp in sys_delims:
        if id(d_sp) in used_delim_ids:
            continue
        used_delim_ids.add(id(d_sp))
        cmd = r"\hoac" if "@@SYS_BRACKET@@" in d_sp.text else r"\heva"

        # Vertical extent of the tall delimiter in TeX CMEX/Yhcmex
        # Note: in CMEX/Yhcmex, the delimiter baseline `d_sp.y` is either near the top (`tp`) or middle/bottom
        # Let's check spans immediately to the right (`d_sp.x + 2 < sp.x < d_sp.x + 90`) within `d_sp.y - 38 .. d_sp.y + 38`
        y_top = d_sp.y + max(18.0, d_sp.font_size * 1.4)
        y_bot = d_sp.y - max(28.0, d_sp.font_size * 2.4)

        # Horizontal right boundary: stop before the next system delimiter or `\Leftrightarrow` / `\Rightarrow` / `\lor` on the same row band!
        x_left = d_sp.x + 1.0
        x_right = 560.0
        for other in spans:
            if other is d_sp:
                continue
            if other.x > d_sp.x + 12.0 and (y_bot - 4.0) <= other.y <= (y_top + 4.0):
                if (
                    "@@SYS_" in other.text
                    or r"\Leftrightarrow" in other.text
                    or r"\Rightarrow" in other.text
                    or r"\lor" in other.text
                    or other.text.strip() in ("⇔", "⇒", "∨", ". Vậy", "Vậy")
                ):
                    x_right = min(x_right, other.x - 1.5)

        sys_boxes.append((x_left, x_right, y_bot, y_top, cmd, d_sp))

    # Group spans into horizontal rows (tolerance ~ 4.5pt for main baseline, while keeping fraction numerators/denominators attached to their row)
    # To keep fraction numerators/denominators with their row, first find horizontal fraction bars
    frac_bars = [
        ln for ln in page_lines if abs(ln.y1 - ln.y2) < 1.2 and 5.0 <= abs(ln.x2 - ln.x1) <= 180.0
    ]

    def _effective_row_y(sp: PdfSpan2D) -> float:
        if "@@SYS_" in sp.text:
            return sp.y - 11.5
        for fb in frac_bars:
            bx1, bx2 = min(fb.x1, fb.x2) - 2.0, max(fb.x1, fb.x2) + 2.0
            if bx1 <= sp.x <= bx2 and abs(sp.y - fb.y1) <= 15.0:
                return fb.y1 - 2.5
        return sp.y

    spans_with_ry = [(sp, _effective_row_y(sp)) for sp in spans]
    spans_with_ry.sort(key=lambda item: (-round(item[1], 1), item[0].x))

    rows: list[tuple[float, list[PdfSpan2D]]] = []
    for sp, ry in spans_with_ry:
        if not rows:
            rows.append((ry, [sp]))
            continue
        prev_ry, prev_list = rows[-1]
        if abs(ry - prev_ry) <= 3.6:
            prev_list.append(sp)
        else:
            rows.append((ry, [sp]))

    # Now merge rows that contain multi-line systems (`@@SYS_BRACKET@@` / `@@SYS_BRACE@@`)
    out_lines: list[str] = []
    r_idx = 0
    while r_idx < len(rows):
        ry, r_spans = rows[r_idx]
        has_sys = any("@@SYS_" in s.text for s in r_spans)
        if not has_sys and r_idx + 1 < len(rows):
            nxt_ry, nxt_spans = rows[r_idx + 1]
            nxt_delims = [s.x for s in nxt_spans if "@@SYS_" in s.text]
            if nxt_delims and (ry - nxt_ry) <= 15.0 and min(s.x for s in r_spans) >= min(nxt_delims) - 5.0:
                has_sys = True

        if has_sys:
            # Collect consecutive rows that belong to this multi-line system band
            band_rows = [r_spans]
            j = r_idx + 1
            first_sys_xs = [s.x for br in band_rows for s in br if "@@SYS_" in s.text]
            if not first_sys_xs and j < len(rows):
                first_sys_xs = [s.x for s in rows[j][1] if "@@SYS_" in s.text]
            min_sys_x = min(first_sys_xs) if first_sys_xs else 0.0
            y_start = ry
            while j < len(rows) and (rows[j - 1][0] - rows[j][0]) <= 19.0 and (y_start - rows[j][0]) <= 36.0 and len(band_rows) < 5:
                row_j_sorted = sorted(rows[j][1], key=lambda s: s.x)
                first_txt = "".join(s.text for s in row_j_sorted[:2]).strip()
                if re.match(r"^(?:Vậy\s+S|Chọn|Câu\s+\d|Trường\s+hợp|Lấy|Suy\s+ra|Đặt|Xét|Ta\s+có|Nếu|Do\s+đó|Kiểm\s+tra|Cách\s+\d|Phương\s+pháp|Dạng\s+\d|BT\d)", first_txt):
                    break
                if re.match(r"^TH\d", first_txt) and (y_start - rows[j][0]) > 14.0:
                    break
                has_j_sys = any("@@SYS_" in s.text for s in row_j_sorted)
                if not has_j_sys and min_sys_x > 0 and row_j_sorted[0].x < min_sys_x - 8.0 and (y_start - rows[j][0]) > 14.0:
                    break
                band_rows.append(rows[j][1])
                j += 1

            if len(band_rows) >= 2:
                all_band_spans = [s for br in band_rows for s in br]
                band_delims = sorted(
                    [s for s in all_band_spans if "@@SYS_" in s.text],
                    key=lambda s: s.x,
                )
                unique_delims: list[PdfSpan2D] = []
                for bd in band_delims:
                    if not unique_delims or abs(bd.x - unique_delims[-1].x) > 6.0:
                        unique_delims.append(bd)

                if unique_delims:
                    first_delim_x = unique_delims[0].x
                    prefix_spans = [
                        s for s in all_band_spans
                        if s.x < first_delim_x - 2.0 and "@@SYS_" not in s.text
                    ]
                    prefix_str = _format_mixed_vietnamese_and_math_line(prefix_spans, page_lines) if prefix_spans else ""

                    sys_strs: list[str] = []
                    last_sys_right_x = 560.0
                    for d_i, d_sp in enumerate(unique_delims):
                        x0 = d_sp.x + 1.0
                        x1 = unique_delims[d_i + 1].x - 2.0 if d_i + 1 < len(unique_delims) else 560.0
                        cmd = r"\hoac" if "@@SYS_BRACKET@@" in d_sp.text else r"\heva"

                        sep_spans = [
                            s for s in all_band_spans
                            if (x0 + 18.0) < s.x < x1
                            and (r"\Leftrightarrow" in s.text or r"\Rightarrow" in s.text or r"\lor" in s.text or s.text.strip() in ("⇔", "⇒", "∨", ". Vậy"))
                        ]
                        mid_op = ""
                        if sep_spans:
                            sep_x = min(s.x for s in sep_spans)
                            if d_i + 1 < len(unique_delims) and sep_x < x1:
                                x1 = sep_x - 1.5
                                op_tok = sep_spans[0].text.strip()
                                mid_op = f" {op_tok} "
                            elif d_i + 1 == len(unique_delims) and sep_x < x1:
                                x1 = sep_x - 1.5
                                last_sys_right_x = sep_x - 1.0

                        branch_lines: list[str] = []
                        for br in band_rows:
                            br_in = [
                                s for s in br
                                if x0 <= s.x <= x1 and "@@SYS_" not in s.text
                            ]
                            if br_in:
                                b_ltx = _render_math_fragment_from_spans(br_in, page_lines).strip()
                                if b_ltx and not re.match(r"^(?:\\Leftrightarrow|\\Rightarrow|\\lor|\\land|[\s.,;:])+$", b_ltx):
                                    branch_lines.append(b_ltx)
                        if branch_lines:
                            sys_strs.append(f"{cmd}{{" + r" \\ ".join(branch_lines) + "}" + mid_op)

                    suffix_spans = [
                        s for s in all_band_spans
                        if s.x >= last_sys_right_x and "@@SYS_" not in s.text
                    ]
                    suffix_str = _format_mixed_vietnamese_and_math_line(suffix_spans, page_lines) if suffix_spans else ""

                    if sys_strs:
                        combined_sys = "$" + " ".join(sys_strs) + "$"
                        full_line = f"{prefix_str} {combined_sys} {suffix_str}".strip()
                        full_line = re.sub(r"\$\s*\$", " ", full_line)
                        out_lines.append(full_line)
                        r_idx = j
                        continue

        # Detect 4-option multiple choice row (A ... B ... C ... D ...) directly from span X positions
        r_sorted = sorted(r_spans, key=lambda s: s.x)
        choice_lbl_indices: list[int] = []
        expected_lbls = ["A", "B", "C", "D"]
        for s_i, s_obj in enumerate(r_sorted):
            t_lbl = s_obj.text.strip().rstrip(".")
            if len(choice_lbl_indices) < 4 and t_lbl == expected_lbls[len(choice_lbl_indices)]:
                if not choice_lbl_indices or (s_obj.x - r_sorted[choice_lbl_indices[-1]].x) >= 45.0:
                    choice_lbl_indices.append(s_i)
        if len(choice_lbl_indices) == 4 and choice_lbl_indices[0] == 0:
            opts_4: list[str] = []
            for c_k in range(4):
                start_i = choice_lbl_indices[c_k]
                end_i = choice_lbl_indices[c_k + 1] if c_k + 1 < 4 else len(r_sorted)
                opt_spans = list(r_sorted[start_i:end_i])
                first_txt = opt_spans[0].text.strip()
                if first_txt in ("A", "B", "C", "D", "A.", "B.", "C.", "D."):
                    opt_spans = opt_spans[1:]
                opt_str = _format_mixed_vietnamese_and_math_line(opt_spans, page_lines).strip() if opt_spans else ""
                opts_4.append(opt_str)
            if all(opts_4):
                out_lines.append(f"\\choice{{{opts_4[0]}}}{{{opts_4[1]}}}{{{opts_4[2]}}}{{{opts_4[3]}}}")
                r_idx += 1
                continue

        # Standard single-row formatting
        line_str = _format_mixed_vietnamese_and_math_line(r_spans, page_lines)
        if line_str:
            out_lines.append(line_str)
        r_idx += 1

    # Post-process multiple-choice option lines (`A ... B ... C ... D ...`) into `\choice{...}{...}{...}{...}`
    final_blocks: list[str] = []
    idx = 0
    while idx < len(out_lines):
        ln = out_lines[idx].strip()
        # Match 4 options on one line: `A -16. B 16. C -12. D -2.` or `A. ... B. ... C. ... D. ...`
        m_4opt = re.match(
            r"^(?:\\textbf\{)?A\.?(?:\})?\s+(.+?)\s+(?:\\textbf\{)?B\.?(?:\})?\s+(.+?)\s+(?:\\textbf\{)?C\.?(?:\})?\s+(.+?)\s+(?:\\textbf\{)?D\.?(?:\})?\s+(.+)$",
            ln,
        )
        if m_4opt:
            o1, o2, o3, o4 = [x.strip() for x in m_4opt.groups()]
            final_blocks.append(f"\\choice{{{o1}}}{{{o2}}}{{{o3}}}{{{o4}}}")
            idx += 1
            continue

        # Match 4 consecutive lines starting with `A.`, `B.`, `C.`, `D.`
        if idx + 3 < len(out_lines):
            l_a, l_b, l_c, l_d = [out_lines[idx + k].strip() for k in range(4)]
            ma = re.match(r"^(?:\\textbf\{)?A\.?(?:\})?\s*(.+)$", l_a)
            mb = re.match(r"^(?:\\textbf\{)?B\.?(?:\})?\s*(.+)$", l_b)
            mc = re.match(r"^(?:\\textbf\{)?C\.?(?:\})?\s*(.+)$", l_c)
            md = re.match(r"^(?:\\textbf\{)?D\.?(?:\})?\s*(.+)$", l_d)
            if ma and mb and mc and md:
                final_blocks.append(
                    f"\\choice{{{ma.group(1).strip()}}}{{{mb.group(1).strip()}}}{{{mc.group(1).strip()}}}{{{md.group(1).strip()}}}"
                )
                idx += 4
                continue

        final_blocks.append(ln)
        idx += 1

    return final_blocks


def decompile_pdf_to_latex(
    pdf_path: Path,
    *,
    search_dirs: list[Path] | None = None,
) -> str:
    """Decompile any Vietnamese Mathematics PDF into structured LaTeX for `convert_latex_file_unified`."""
    pdf_path = pdf_path.resolve()

    # 1. Check if a companion .tex source with the exact same stem exists in the same folder
    companion_tex = pdf_path.with_suffix(".tex")
    if companion_tex.exists() and companion_tex.stat().st_size > 50:
        return clean_newlines(companion_tex.read_text(encoding="utf-8", errors="replace"))

    # 2. Extract 2D spans, lines, and rectangles from all pages
    pages_2d = extract_pdf_pages_2d(pdf_path)

    doc_blocks: list[str] = [
        r"\documentclass[12pt,a4paper]{article}",
        r"\usepackage{amsmath,amssymb}",
        r"\usepackage{tikz,pgfplots,tkz-tab}",
        r"\begin{document}",
    ]

    for page_idx, (raw_spans, raw_lines, raw_rects) in enumerate(pages_2d):
        spans = _deduplicate_and_filter_page_spans(raw_spans)
        if not spans:
            continue

        # Attach underscripts directly below max / min / lim
        spans = _attach_underscripts_to_operators(spans)

        # Check for Answer Key Grid (`BẢNG ĐÁP ÁN`)
        ans_table_tex, spans = _detect_and_extract_answer_grid(spans)

        # Check for Variation Tables (`Bảng biến thiên`)
        bbt_tables, spans = _detect_and_extract_bbt_tables(spans, raw_lines)

        if bbt_tables:
            # Split spans above and below each BBT table Y coordinate so BBT appears in exact vertical position
            bbt_tables.sort(key=lambda item: -item[0])
            rem_spans = spans
            for bbt_y, bbt_tex in bbt_tables:
                above_s = [s for s in rem_spans if s.y > bbt_y]
                rem_spans = [s for s in rem_spans if s.y <= bbt_y]
                doc_blocks.extend(_extract_multiline_systems_and_lines(above_s, raw_lines))
                doc_blocks.append(bbt_tex)
            if rem_spans:
                doc_blocks.extend(_extract_multiline_systems_and_lines(rem_spans, raw_lines))
        else:
            doc_blocks.extend(_extract_multiline_systems_and_lines(spans, raw_lines))

        if ans_table_tex:
            doc_blocks.append(ans_table_tex)

    doc_blocks.append(r"\end{document}")
    return clean_newlines("\n\n".join(doc_blocks))


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
