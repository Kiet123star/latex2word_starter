"""Phase B handlers for `latex_preprocessor.py` (block structures, inline text, math, tables, lists, exam macros).

Tách riêng cho dễ đọc; các hàm ở đây được gắn vào lớp `LatexPreprocessor` qua `install_parts()`.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

from . import latex_preprocessor as L
from .latex_preprocessor import (
    BOX_ENVS, CHOICE_LIST_ENVS, CIRCLED, DING_MAP, ENGINE_MATH_CMDS, ENGINE_MATH_SYMBOL_NAMES, FRAK_LOWER,
    FRAK_UPPER, BB_DIGITS, BB_LOWER, BB_UPPER, LIST_ENVS, LIT_CHARS, LIT_REVERSE, MATH_DISPLAY_ENVS,
    MATH_OPERATOR_NAMES, MATH_OPERATOR_SPECIAL, MATH_SYMBOLS_ALL, MATH_SYMBOLS_EXTRA, MK_DOTFILL,
    MK_HFILL, MK_NUMPAGES, MK_PAGE, MK_RULEFILL, MK_TAB, MK_TABLE, NOT_MAP,
    NUM_FORMATS, SCRIPT_LOWER, SCRIPT_UPPER, SI_UNITS, SIZE_SWITCH_PT, TABLE_ENVS, TEXT_ACCENTS,
    TEXT_SYMBOLS, THEOREM_BUILTINS_EN, THEOREM_BUILTINS_VI, TRANSPARENT_ENVS, _REF_C, _REF_O,
    _VERB_INLINE, _Ctx, _ICtx, _Pending, _map_alpha, _roman, _alph, _today_str, char_idx, dim_to_twips,
    find_env_end, idx_char, lit, read_arg, read_bracket, read_braced, read_cmd, read_group, read_opt,
    read_paren, read_star, skip_spaces_only, skip_ws, split_top_level,
)

PLACEHOLDER_IMG_SPAN = {"i": 1, "c": "64748B"}

CHOICE_LIKE_RE = re.compile(r"(?:pa|PA|choice|Choice|chon|Chon|opt|Opt|ans|Ans|phuongan|tracnghiem|luachon)")
QUESTION_LIKE = {"cau": "Câu", "Cau": "Câu", "bai": "Bài", "Bai": "Bài", "vidu": "Ví dụ", "baitap": "Bài tập",
                 "cauhoi": "Câu hỏi", "caun": "Câu", "question@": "Câu"}
INLINE_LIST_ENVS = {"inparaenum", "inparaitem", "inlinelist"}


# =====================================================================================
# INLINE RENDERER
# =====================================================================================
def _inline(self, s: str, ictx: _ICtx) -> str:
    out: list[str] = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == "\\":
            name, j = read_cmd(s, i)
            r = self._inline_cmd(name, s, i, j, ictx)
            text, k = r
            out.append(text)
            i = k
            continue
        if c == "{":
            inner, j = read_group(s, i)
            out.append(self._inline(inner, _ICtx(**ictx.__dict__)))
            i = j
            continue
        if c == "}":
            i += 1
            continue
        if c == "$":
            if s.startswith("$$", i):
                end = s.find("$$", i + 2)
                if end == -1:
                    end = n
                out.append(self._inline_math(s[i + 2 : end], ictx))
                i = end + 2
                continue
            end = self._find_inline_math_end(s, i + 1)
            out.append(self._inline_math(s[i + 1 : end], ictx))
            i = end + 1
            continue
        if c == "~":
            out.append("\u00a0")
            i += 1
            continue
        if c == "-":
            if s.startswith("---", i):
                out.append("—")
                i += 3
            elif s.startswith("--", i):
                out.append("–")
                i += 2
            else:
                out.append("-")
                i += 1
            continue
        if c == "`":
            if s.startswith("``", i):
                out.append("“")
                i += 2
            else:
                out.append("‘")
                i += 1
            continue
        if c == "'":
            if s.startswith("''", i):
                out.append("”")
                i += 2
            else:
                out.append("’")
                i += 1
            continue
        if c in "\n\t\r":
            out.append(" ")
            i += 1
            continue
        if c in "&%#":
            out.append(c if (ictx.plain or c == "#") else LIT_CHARS[c])
            i += 1
            continue
        if c == _VERB_INLINE and i + 1 < n:
            body = self.verbatims[char_idx(s[i + 1])][1]
            if ictx.plain:
                out.append(body)
            else:
                out.append(self.span({"f": "Courier New"}, lit(body).replace(" ", "\u00a0")))
            i += 2
            continue
        if c in "!?" and s.startswith("`", i + 1):
            out.append("¡" if c == "!" else "¿")
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _inline_math(self, m: str, ictx: _ICtx) -> str:
    norm = self._math(m)
    if not norm.strip():
        return ""
    if ictx.plain:
        return self._math_to_plain(norm)
    return "$" + norm + "$"


def _eat_space_after_word(s: str, k: int) -> int:
    """TeX swallows spaces after a control word; also eat an empty {} group."""
    n = len(s)
    j = k
    while j < n and s[j] in " \t":
        j += 1
    if j < n and s[j] == "\n" and not s[j + 1 : j + 2] == "\n":
        j += 1
        while j < n and s[j] in " \t":
            j += 1
    if s.startswith("{}", j):
        return j + 2
    return j


def _wrap_fmt(self, ictx: _ICtx, kind: str, text: str) -> str:
    if not text:
        return ""
    if ictx.plain:
        return text
    if kind == "b":
        if ictx.bold_via_span:
            return self.span({"b": 1}, text)
        return "\\textbf{" + text + "}"
    if kind == "i":
        return "\\textit{" + text + "}"
    if kind == "up":
        return self.span({"i": 0}, text)
    if kind == "u":
        return "\\underline{" + text + "}"
    if kind == "nob":
        return self.span({"b": 0}, text)
    return text


FMT_CMDS = {
    "textbf": "b", "textit": "i", "textsl": "i", "emph": "emph", "underline": "u", "uline": "u", "ul": "u",
    "underbar": "u", "textup": "up", "textrm": "", "textnormal": "norm", "textmd": "nob", "textsf": "sf",
    "texttt": "tt", "textsc": "sc", "sout": "s", "st": "s", "xout": "s", "uuline": "uu", "uwave": "uw",
    "dashuline": "ud", "dotuline": "udot", "hl": "hl", "highlight": "hl", "textsuperscript": "sup",
    "textsubscript": "sub", "up": "sup", "caps": "caps", "textls": "", "so": "", "mbox": "", "hbox": "",
    "text": "", "textmath": "", "nolinkurl": "tt", "path": "tt", "alert": "alert", "structure": "",
    "textcolorbox": "", "invisible": "", "smash": "", "Centerstack": "", "Longstack": "",
}
DECLARATIONS = {
    "bfseries": "b", "bf": "b", "itshape": "i", "it": "i", "slshape": "i", "sl": "i", "em": "emph",
    "upshape": "up", "rmfamily": "", "rm": "", "normalfont": "norm", "mdseries": "nob", "sffamily": "sf",
    "sf": "sf", "ttfamily": "tt", "tt": "tt", "scshape": "sc", "sc": "sc", "boldmath": "", "unboldmath": "",
}


def _apply_fmt(self, kind: str, text: str, ictx: _ICtx) -> str:
    if kind in ("b", "i", "u", "up", "nob"):
        return self._wrap_fmt(ictx, kind, text)
    if kind == "emph":
        return self._wrap_fmt(ictx, "up" if ictx.italic else "i", text)
    if ictx.plain or not text:
        return text
    spec = {
        "sf": {"f": "Arial"}, "tt": {"f": "Courier New"}, "sc": {"sc": 1}, "s": {"s": 1}, "uu": {"u": "double"},
        "uw": {"u": "wave"}, "ud": {"u": "dash"}, "udot": {"u": "dotted"}, "hl": {"hl": "yellow"},
        "sup": {"va": "sup"}, "sub": {"va": "sub"}, "caps": {"caps": 1}, "alert": {"c": "FF0000"},
        "norm": {"b": 0, "i": 0},
    }.get(kind)
    return self.span(spec, text) if spec else text


def _inline_cmd(self, name: str, s: str, i: int, j: int, ictx: _ICtx) -> tuple[str, int]:
    n = len(s)
    plain = ictx.plain
    # ---------- control symbols ----------
    if not name:
        return "", j
    if not (name[0].isalpha() or name[0] == "@"):
        if name == "\\":
            k = j
            if k < n and s[k] == "*":
                k += 1
            kk = skip_spaces_only(s, k)
            if kk < n and s[kk] == "[":
                _, k = read_bracket(s, kk)
            return (" " if plain else " \\\\ "), k
        if name in ("{", "}", "$", "%", "&", "_", "#", "^"):
            if name in ("_", "#"):
                return name, j
            if name == "^" :
                a, k = read_arg(s, j)
                if a is not None:
                    return self._accent("^", a, ictx), k
                return "^", j
            return (name if plain else LIT_CHARS[name]), j
        if name in (" ", "\n"):
            return " ", j
        if name == ",":
            return "\u2009", j
        if name == ";":
            return "\u2005", j
        if name == ":":
            return "\u205f", j
        if name in ("!", "/", "@", "|"):
            return ("‖" if name == "|" else ""), j
        if name == "-":
            return "\u00ad", j
        if name in TEXT_ACCENTS:
            a, k = read_arg(s, j)
            if a is None:
                return "", j
            return self._accent(name, a, ictx), k
        if name == "(":
            end = s.find("\\)", j)
            if end == -1:
                end = n
            return self._inline_math(s[j:end], ictx), min(n, end + 2)
        if name == "[":
            end = self._find_display_end(s, j)
            return self._inline_math(s[j:end], ictx), min(n, end + 2)
        if name == "]" or name == ")":
            return "", j
        return "", j

    # ---------- accents with letter names ----------
    if name in TEXT_ACCENTS and len(name) <= 2 or name in ("textcommabelow", "textdotbelow", "texthookabove"):
        a, k = read_arg(s, j)
        if a is not None and name in TEXT_ACCENTS:
            return self._accent(name, a, ictx), k

    # ---------- symbols ----------
    if name in TEXT_SYMBOLS:
        val = TEXT_SYMBOLS[name]
        if plain:
            val = "".join(LIT_REVERSE.get(ch, ch) for ch in val)
        return val, _eat_space_after_word(s, j)
    if name == "today":
        return _today_str(self.lang), _eat_space_after_word(s, j)
    if name in ("hfill", "hfil", "fill"):
        return ("  " if plain else MK_HFILL), _eat_space_after_word(s, j)
    if name == "dotfill":
        return ("…" if plain else MK_DOTFILL), _eat_space_after_word(s, j)
    if name == "hrulefill":
        return ("____" if plain else MK_RULEFILL), _eat_space_after_word(s, j)
    if name in ("quad", "qquad", "enspace", "enskip", "thinspace"):
        return TEXT_SYMBOLS[name], _eat_space_after_word(s, j)

    # ---------- formatting ----------
    if name in FMT_CMDS:
        st, k = read_star(s, j)
        o, k = read_opt(s, k, allow_newline=False) if name in ("hl",) else (None, k)
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        kind = FMT_CMDS[name]
        sub = _ICtx(**ictx.__dict__)
        if kind == "b":
            sub.bold = True
        if kind in ("i",):
            sub.italic = True
        if kind == "emph":
            sub.italic = not ictx.italic
        if kind == "up":
            sub.italic = False
        inner = self._inline(a, sub)
        return self._apply_fmt(kind, inner, ictx), k
    if name in DECLARATIONS:
        kind = DECLARATIONS[name]
        rest = s[j:]
        sub = _ICtx(**ictx.__dict__)
        if kind == "b":
            sub.bold = True
        if kind == "i":
            sub.italic = True
        if kind == "emph":
            sub.italic = not ictx.italic
        inner = self._inline(rest.lstrip(" "), sub)
        if not kind:
            return inner, n
        return self._apply_fmt(kind, inner, ictx), n
    if name == "fontsize":
        a, k = read_arg(s, j)
        b, k = read_arg(s, k)
        inner = self._inline(s[k:].lstrip(" "), ictx)
        try:
            pt = float(re.sub(r"[^\d.]", "", a or "") or 0)
        except ValueError:
            pt = 0
        if plain or not inner.strip() or pt <= 0:
            return inner, n
        return self.span({"sz": pt}, inner), n
    if name in SIZE_SWITCH_PT:
        inner = self._inline(s[j:].lstrip(" "), ictx)
        if plain or not inner.strip():
            return inner, n
        return self.span({"sz": SIZE_SWITCH_PT[name]}, inner), n
    if name == "color":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_braced(s, k)
        inner = self._inline(s[k:], ictx)
        hx = self.color_hex(a or "", o)
        if plain or not hx:
            return inner, n
        return self.span({"c": hx}, inner), n
    if name in ("textcolor", "colorbox", "fcolorbox", "highlightbox", "fcolorbox*"):
        o, k = read_opt(s, j, allow_newline=False)
        c1, k = read_braced(s, k)
        c2 = None
        if name.startswith("fcolorbox"):
            c2, k = read_braced(s, k)
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        inner = self._inline(a, ictx)
        if plain:
            return inner, k
        if name == "textcolor":
            hx = self.color_hex(c1 or "", o)
            return (self.span({"c": hx}, inner) if hx else inner), k
        if name == "colorbox":
            hx = self.color_hex(c1 or "", o)
            return (self.span({"shd": hx}, inner) if hx else inner), k
        hx = self.color_hex(c2 or "", o)
        bc = self.color_hex(c1 or "", o)
        spec: dict[str, Any] = {"bdr": bc or "000000"}
        if hx:
            spec["shd"] = hx
        return self.span(spec, inner), k
    if name in ("fbox", "framebox", "boxed", "frame", "tcbox", "shadowbox", "doublebox", "ovalbox", "Ovalbox", "cfbox", "khung"):
        k = j
        if name == "cfbox":
            _, k = read_braced(s, k)
        while True:
            o, k2 = read_opt(s, k, allow_newline=False)
            if o is None:
                break
            k = k2
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        inner = self._inline(a, ictx)
        return (inner if plain else self.span({"bdr": "000000"}, inner)), k
    if name in ("makebox", "parbox", "raisebox", "scalebox", "resizebox", "rotatebox", "reflectbox",
                "adjustbox", "smash", "clap", "llap", "rlap", "mathclap", "mathllap", "mathrlap",
                "centerline", "rightline", "leftline", "shortstack", "makecell", "thead", "Centerstack",
                "stackunder", "stackon", "savebox", "usebox", "sbox", "vcenter", "vbox", "vtop", "colorbox*",
                "textcircled@"):
        k = j
        st, k = read_star(s, k)
        lead = {"parbox": 1, "raisebox": 1, "scalebox": 1, "resizebox": 2, "rotatebox": 1, "adjustbox": 1, "sbox": 1,
                "savebox": 1}.get(name, 0)
        while True:
            o, k2 = read_opt(s, k, allow_newline=False)
            if o is None:
                break
            k = k2
        for _ in range(lead):
            _, k = read_arg(s, k)
            while True:
                o, k2 = read_opt(s, k, allow_newline=False)
                if o is None:
                    break
                k = k2
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        if name in ("shortstack", "makecell", "thead", "Centerstack"):
            return self._inline(a, ictx), k
        return self._inline(a, ictx), k
    if name in ("phantom", "hphantom"):
        a, k = read_arg(s, j)
        txt = self._plain(a or "")
        return "\u00a0" * max(1, len(txt)), k
    if name in ("vphantom", "label", "index", "glossary", "nocite", "addcontentsline", "markboth", "markright",
                "hypertarget", "pdfbookmark", "bookmark", "cleaninternal@", "linelabel", "zlabel", "hyphenation"):
        if name == "label":
            a, k = read_braced(s, j)
            if a is not None:
                self.label_values[a.strip()] = self.current_label
                return "", k
            return "", j
        nargs = {"addcontentsline": 3, "markboth": 2, "hypertarget": 2, "pdfbookmark": 2}.get(name, 1)
        k = j
        for _ in range(nargs):
            _, k2 = read_opt(s, k, allow_newline=False)
            k = k2
            a, k = read_arg(s, k)
        if name == "hypertarget":
            # keep the visible text (2nd arg)
            return "", k
        return "", k
    if name in ("hspace", "hskip", "kern", "hspace*"):
        st, k = read_star(s, j)
        if name in ("hskip", "kern"):
            m = re.match(r"\s*=?\s*(-?[\d.]+\s*(?:cm|mm|pt|em|ex|in|bp)?)", s[k:])
            arg = m.group(1) if m else ""
            k = k + (m.end() if m else 0)
        else:
            a, k = read_arg(s, k)
            arg = a or ""
        tw = dim_to_twips(arg) or 0
        if tw <= 0:
            return "", k
        spaces = max(1, round(tw / 120))
        if "fill" in arg:
            return MK_HFILL, k
        return "\u00a0" * min(spaces, 60), k
    if name in ("vspace", "vskip", "smallskip", "medskip", "bigskip", "addvspace", "vfill", "par", "noindent",
                "indent", "centering", "raggedright", "raggedleft", "newpage", "clearpage", "pagebreak",
                "linebreak@", "nopagebreak", "samepage", "pause", "vfil", "displaybreak", "allowdisplaybreaks"):
        k = j
        if name in ("vspace", "addvspace", "vspace*"):
            _, k = read_star(s, k)
            _, k = read_arg(s, k)
        elif name == "vskip":
            m = re.match(r"\s*[-\d.]+\s*(?:cm|mm|pt|em|ex|in)?(?:\s*plus\s*[-\d.]+\s*\w+)?(?:\s*minus\s*[-\d.]+\s*\w+)?", s[k:])
            k += m.end() if m else 0
        elif name in ("pagebreak", "nopagebreak", "linebreak@"):
            _, k = read_opt(s, k, allow_newline=False)
        return (" " if name == "par" else ""), _eat_space_after_word(s, k)
    if name in ("newline", "linebreak", "tabularnewline", "cr", "crcr"):
        k = j
        o, k2 = read_opt(s, k, allow_newline=False)
        if o is not None:
            k = k2
        return (" " if plain else " \\\\ "), _eat_space_after_word(s, k)

    # ---------- references / citations / links ----------
    if name in ("ref", "eqref", "pageref", "autoref", "cref", "Cref", "vref", "nameref", "cpageref", "labelcref",
                "autopageref", "ref*", "subref"):
        st, k = read_star(s, j)
        a, k = read_braced(s, k)
        if a is None:
            return "", j
        key = a.strip().split(",")[0].strip()
        if name == "pageref":
            if key.lower() in ("lastpage", "lastpage@"):
                return ("?" if plain else MK_NUMPAGES), k
            return "…", k
        kind = {"eqref": "e", "autoref": "a", "cref": "a", "Cref": "a", "vref": "a"}.get(name, "r")
        return _REF_O + kind + key + _REF_C, k
    if name == "hyperref":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        return self._inline(a or "", ictx), k
    if name in ("cite", "citep", "citet", "parencite", "textcite", "autocite", "footcite", "citealp", "citeauthor",
                "citeyear", "Cite", "Citep", "Citet"):
        st, k = read_star(s, j)
        o1, k = read_opt(s, k, allow_newline=False)
        o2, k2 = read_opt(s, k, allow_newline=False)
        if o2 is not None:
            k = k2
        a, k = read_braced(s, k)
        if a is None:
            return "", j
        note = o2 if o2 is not None else o1
        inner = _REF_O + "c" + a + _REF_C
        if note:
            inner += ", " + self._inline(note, ictx)
        return "[" + inner + "]", k
    if name in ("url", "nolinkurl", "email"):
        a, k = read_braced(s, j)
        txt = (a or "").replace("\\%", "%").replace("\\#", "#").replace("\\_", "_").replace("\\&", "&")
        return (txt if plain else self.span({"f": "Courier New"}, lit(txt))), k
    if name == "href":
        a, k = read_braced(s, j)
        b, k = read_arg(s, k)
        return self._inline(b or "", ictx), k
    if name in ("footnote", "footnotetext", "thanks", "endnote", "tablefootnote", "sidenote", "marginpar",
                "marginnote", "todo"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        if plain or name in ("marginpar", "marginnote", "todo"):
            return "", k
        body = self._inline(a, _ICtx())
        self.res.footnotes.append(body.strip())
        self.step("footnote")
        self._stat("footnotes")
        return L.MK_FOOTNOTE + idx_char(len(self.res.footnotes) - 1), k
    if name == "footnotemark":
        o, k = read_opt(s, j, allow_newline=False)
        return "", k

    # ---------- counters ----------
    if name in NUM_FORMATS:
        a, k = read_braced(s, j)
        if a is None:
            return "", j
        cname = a.strip().rstrip("*")
        return NUM_FORMATS[name](self.counters.get(cname, 0)), k
    if name.startswith("the") and name[3:] in self.counters:
        return self.the(name[3:]), _eat_space_after_word(s, j)
    if name == "thepage":
        return ("1" if plain else MK_PAGE), _eat_space_after_word(s, j)
    if name == "value":
        a, k = read_braced(s, j)
        return str(self.counters.get((a or "").strip(), 0)), k
    if name in self._CONFIG_CMDS:
        k = self._config_cmd(name, s, j, None)
        return "", (k if k is not None else j)

    # ---------- special text ----------
    if name == "textcircled":
        a, k = read_arg(s, j)
        t = self._plain(a or "")
        if t.isdigit() and 0 <= int(t) <= 20:
            return CIRCLED[int(t)], k
        if len(t) == 1 and t.isalpha():
            base = 0x24B6 if t.isupper() else 0x24D0
            return chr(base + ord(t.lower()) - 97), k
        return "(" + t + ")", k
    if name in ("ding", "dingnum"):
        a, k = read_braced(s, j)
        return DING_MAP.get((a or "").strip(), "•"), k
    if name in ("ce", "cee", "chemfig@"):
        a, k = read_braced(s, j)
        if a is None:
            return "", j
        m = self._chem_to_math(a)
        return (self._math_to_plain(m) if plain else "$" + m + "$"), k
    if name in ("SI", "qty", "SIrange", "qtyrange", "si", "unit", "num", "numrange", "ang", "SIlist", "numlist",
                "qtylist", "tablenum"):
        out, k = self._siunitx(name, s, j)
        if plain:
            return self._math_to_plain(out), k
        return "$" + out + "$", k
    if name == "ensuremath":
        a, k = read_arg(s, j)
        return self._inline_math(a or "", ictx), k
    if name in ("heva", "hoac"):
        a, k = read_arg(s, j)
        return self._inline_math("\\" + name + "{" + (a or "") + "}", ictx), k
    if name == "includegraphics":
        st, k = read_star(s, j)
        o, k = read_opt(s, k, allow_newline=False)
        a, k = read_braced(s, k)
        nm = (a or "").strip()
        self.res.warnings.append(f"Ảnh '{nm}' nằm trong dòng chữ/ô bảng — chèn chỗ giữ chỗ.")
        return ("" if plain else self.span(PLACEHOLDER_IMG_SPAN, f"[Hình: {lit(nm)}]")), k
    if name == "tikz":
        k = j
        o, k = read_opt(s, k, allow_newline=False)
        kk = skip_ws(s, k)
        if kk < n and s[kk] == "{":
            _, k = read_group(s, kk)
        else:
            e = s.find(";", k)
            k = n if e == -1 else e + 1
        self.res.warnings.append("Bỏ qua hình \\tikz nhỏ nằm trong dòng chữ.")
        return "", k
    if name == "item":
        o, k = read_opt(s, j, allow_newline=False)
        lbl = self._inline(o, ictx) if o is not None else "•"
        return " " + lbl + " ", _eat_space_after_word(s, k)
    if name in ("cau", "Cau", "bai", "Bai") and name not in self.expander.macros:
        title = "Câu" if name.lower() == "cau" else "Bài"
        num, k = self._question_number(s, j, "ex")
        txt = f"{title} {num}."
        return (txt + " " if plain else self._wrap_fmt(ictx, "b", txt) + " "), k
    if name in ("True", "Truee"):
        return "", _eat_space_after_word(s, j)
    if name in ("choice", "choiceTF", "choiceTFt", "boncot", "haicot", "motcot", "bonpa", "choiceA", "choiceB",
                "choiceC", "choiceD"):
        opts, k = self._read_n_braced(s, j, 8 if name.startswith("choiceTF") else 4)
        labels = ["A.", "B.", "C.", "D."] if not name.startswith("choiceTF") else ["a)", "b)", "c)", "d)"]
        parts = []
        for idx, o in enumerate(opts[:4] if not name.startswith("choiceTF") else opts[::2] if len(opts) == 8 else opts):
            o2 = o.replace("\\True", "")
            parts.append(self._wrap_fmt(ictx, "b", labels[idx % 4]) + " " + self._inline(o2, ictx))
        return "\u2003".join(parts), k
    if name in ("loigiai", "giai", "huongdan", "shortans", "dapso"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        lbl = {"loigiai": "Lời giải.", "giai": "Giải.", "huongdan": "Hướng dẫn.", "shortans": "Đáp án:", "dapso": "Đáp số:"}[name]
        if name in ("loigiai", "giai", "huongdan") and not self.show_solutions:
            return "", k
        return self._wrap_fmt(ictx, "b", lbl) + " " + self._inline(a or "", ictx), k
    if name == "begin":
        return self._inline_env(s, j, ictx)
    if name == "end":
        a, k = read_braced(s, j)
        return "", k
    if name in ("pause", "onslide", "only", "uncover", "visible", "invisible", "alt", "temporal", "action"):
        k = j
        if k < n and s[k] == "<":
            e = s.find(">", k)
            k = e + 1 if e != -1 else k
        if name == "pause":
            return "", k
        a, k2 = read_arg(s, k)
        if name == "alt":
            b, k2 = read_arg(s, k2)
        return (self._inline(a or "", ictx) if name != "invisible" else ""), k2
    if name == "multicolumn":
        a, k = read_arg(s, j)
        b, k = read_arg(s, k)
        c, k = read_arg(s, k)
        return self._inline(c or "", ictx), k
    if name == "multirow":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        o, k = read_opt(s, k, allow_newline=False)
        b, k = read_arg(s, k)
        o, k = read_opt(s, k, allow_newline=False)
        c, k = read_arg(s, k)
        return self._inline(c or "", ictx), k
    if name in ("diagbox", "backslashbox", "slashbox"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        b, k = read_arg(s, k)
        return self._inline(a or "", ictx) + " / " + self._inline(b or "", ictx), k
    if name in ("hline", "cline", "toprule", "midrule", "bottomrule", "cmidrule", "hhline", "Xhline",
                "rowcolor", "cellcolor", "arrayrulecolor", "endhead", "endfirsthead", "endfoot", "endlastfoot",
                "specialrule", "addlinespace", "morecmidrules", "hdashline", "cdashline", "rule"):
        k = j
        if name in ("cline", "cmidrule", "hhline", "Xhline", "rowcolor", "cellcolor", "arrayrulecolor", "specialrule",
                    "cdashline", "rule"):
            if name == "cmidrule":
                kk = skip_spaces_only(s, k)
                if kk < n and s[kk] == "(":
                    _, k = read_paren(s, kk)
            o, k2 = read_opt(s, k, allow_newline=False)
            if o is not None:
                k = k2
            cnt = {"specialrule": 3, "rule": 2}.get(name, 1)
            for _ in range(cnt):
                a, k = read_arg(s, k)
        elif name == "addlinespace":
            o, k2 = read_opt(s, k, allow_newline=False)
            k = k2 if o is not None else k
        if name == "rule":
            m_w = re.search(r"\\rule\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}", s[i:k])
            tw = dim_to_twips(m_w.group(1), self.textwidth) if m_w else None
            if plain:
                return "____", k
            if tw and tw < self.textwidth * 0.9:
                return "_" * max(2, int(tw / 120)), k
            return MK_RULEFILL, k
        return "", k
    if name in ("tableofcontents", "listoffigures", "listoftables", "maketitle", "printbibliography",
                "appendix", "frontmatter", "mainmatter", "backmatter", "onecolumn", "twocolumn", "columnbreak",
                "titlepage@"):
        return "", _eat_space_after_word(s, j)

    # ---------- math-only commands used in text (LaTeX would auto-error; we wrap them as math) ----------
    if name in MATH_SYMBOLS_ALL or name in ENGINE_MATH_CMDS or name in MATH_OPERATOR_NAMES or name in (
        "frac", "dfrac", "sqrt", "vec", "overrightarrow", "overline", "widehat", "mathbb", "mathcal", "mathrm",
        "mathbf", "binom", "lim", "sum", "int", "prod"
    ):
        if name not in ("text", "textbf", "textit", "textrm", "underline", "begin", "end", "left", "right"):
            k = self._skip_inline_unit(s, i, j, name)
            # also eat following sub/superscripts
            while k < n and s[k] in "_^":
                a, k = read_arg(s, k + 1)
            return self._inline_math(s[i:k], ictx), k

    # ---------- user environments leftovers / unknown command fallback ----------
    if name in self.expander.macros:
        # macro left unexpanded (missing args / expansion limit) — never print "\\name"
        m = self.expander.macros[name]
        if m.nargs == 0 and ("\\" + name) not in m.body and getattr(self, "_inline_depth", 0) < 60:
            self._inline_depth = getattr(self, "_inline_depth", 0) + 1
            try:
                return self._inline(m.body, ictx), _eat_space_after_word(s, j)
            finally:
                self._inline_depth -= 1
        return "", _eat_space_after_word(s, j)
    self._unknown(name)
    k = j
    st, k = read_star(s, k)
    raw_args: list[str] = []
    while True:
        kk = skip_spaces_only(s, k)
        if kk < n and s[kk] == "[":
            _, k = read_bracket(s, kk)
            continue
        if kk < n and s[kk] == "{":
            a, k = read_group(s, kk)
            raw_args.append(a)
            continue
        break
    pieces = []
    for idx, a in enumerate(raw_args):
        if re.fullmatch(r"\s*-?[\d.]*\s*(cm|mm|pt|em|ex|in|bp)?\s*(\\(textwidth|linewidth|columnwidth))?\s*", a):
            continue
        if len(raw_args) > 1 and idx == 0 and a.strip().lower() in XCOLOR_LIKE:
            continue
        t = self._inline(a, ictx)
        if t.strip():
            pieces.append(t)
    if not pieces:
        return "", (_eat_space_after_word(s, k) if not raw_args else k)
    return " ".join(pieces), k


XCOLOR_LIKE = {"red", "blue", "green", "black", "white", "cyan", "magenta", "yellow", "gray", "grey", "orange",
               "purple", "brown", "violet", "pink", "teal", "olive", "lime", "darkgray", "lightgray", "navy"}


def _accent(self, name: str, arg: str, ictx: _ICtx) -> str:
    comb = TEXT_ACCENTS.get(name, "")
    base = self._plain(arg) if arg.strip() else " "
    if not base:
        base = " "
    if base in ("ı",):
        base = "i"
    if base in ("ȷ",):
        base = "j"
    res = base[0] + comb + base[1:]
    return unicodedata.normalize("NFC", res)


def _plain(self, s: str) -> str:
    out = self._inline(s, _ICtx(plain=True))
    out = "".join(LIT_REVERSE.get(ch, ch) for ch in out)
    out = re.sub("[\ue000-\ue00f]", "", out)
    return re.sub(r"\s+", " ", out).strip()


def _inline_env(self, s: str, j: int, ictx: _ICtx) -> tuple[str, int]:
    a, k = read_braced(s, j)
    if a is None:
        return "", j
    env = a.strip()
    ce, after = find_env_end(s, k, env)
    content = s[k:ce]
    base = env.rstrip("*")
    if env in MATH_DISPLAY_ENVS or base in ("math", "displaymath"):
        body = re.sub(r"^\s*\{[^{}]*\}", "", content) if base in ("alignat", "xalignat") else content
        body = re.sub(r"\\(nonumber|notag)\b|\\label\{[^{}]*\}", "", body)
        if "\\\\" in body:
            body = "\\begin{aligned}" + body + "\\end{aligned}"
        return self._inline_math(body, ictx), after
    if env in ("cases", "aligned", "array", "matrix", "pmatrix", "bmatrix", "vmatrix", "gathered", "split"):
        return self._inline_math("\\begin{" + env + "}" + content + "\\end{" + env + "}", ictx), after
    if env in TABLE_ENVS:
        rows = [r for r in self._split_rows(re.sub(r"^\s*(\[[^\]]*\])?\s*\{[^{}]*(\{[^{}]*\}[^{}]*)*\}", "", content, count=1))]
        parts = []
        for r in rows:
            r = re.sub(r"\\(hline|toprule|midrule|bottomrule)\b|\\cline\{[^}]*\}", "", r).strip()
            if r:
                parts.append(" ".join(self._inline(c, ictx) for c in self._split_cells(r)))
        sep = " " if ictx.plain else " \\\\ "
        return sep.join(p for p in parts if p.strip()), after
    if env in LIST_ENVS or env in INLINE_LIST_ENVS:
        chunks = split_top_level(content, ("item",))
        parts = []
        cnt = 0
        for cmd, chunk in chunks:
            if cmd != "item":
                continue
            cnt += 1
            o, kk = read_opt(chunk, 0)
            if o is not None:
                lbl = self._inline(o, ictx)
                chunk = chunk[kk:]
            elif env.startswith("enum") or env.startswith("inparaenum"):
                lbl = f"({cnt})" if env.startswith("inpara") else f"{cnt}."
            else:
                lbl = "•"
            parts.append(lbl + " " + self._inline(chunk.strip(), ictx))
        return " ".join(parts), after
    if env in ("tikzpicture", "axis", "pspicture", "tikzcd"):
        self.res.warnings.append("Hình vẽ nằm trong ô bảng/dòng chữ — bỏ qua (Word không đặt được hình trong vị trí này).")
        return ("" if ictx.plain else self.span(PLACEHOLDER_IMG_SPAN, "[Hình vẽ]")), after
    # transparent
    content2 = re.sub(r"^\s*(\[[^\]]*\])?(\{[^{}]*\})?", "", content, count=1) if env in ("minipage", "adjustbox", "tcolorbox") else content
    return self._inline(content2, ictx), after


def _read_n_braced(self, s: str, j: int, nmax: int) -> tuple[list[str], int]:
    args: list[str] = []
    k = j
    while len(args) < nmax:
        kk = skip_ws(s, k)
        if kk < len(s) and s[kk] == "{":
            a, k = read_group(s, kk)
            args.append(a)
        else:
            break
    return args, k


def _question_number(self, s: str, j: int, counter: str) -> tuple[str, int]:
    n = len(s)
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    if o is not None:
        k = k2
    kk = skip_spaces_only(s, k)
    if kk < n and s[kk] == "{":
        a, k3 = read_group(s, kk)
        num = self._plain(a)
        if num:
            if num.isdigit():
                self.counters[counter] = int(num)
            self.current_label = (num, counter)
            return num, k3
        return self.refstep(counter, counter), k3
    m = re.match(r"(\d+)", s[k:])
    if m and k == j:
        num = m.group(1)
        self.counters[counter] = int(num)
        self.current_label = (num, counter)
        return num, k + m.end()
    return self.refstep(counter, counter), _eat_space_after_word(s, k)


# =====================================================================================
# MATH NORMALIZER
# =====================================================================================
TEXT_IN_MATH = {"text", "textrm", "textnormal", "mbox", "hbox", "textup", "textmd", "textsf", "texttt",
                "emph", "textit", "textbf", "textsl", "textsc", "intertext", "shortintertext", "mathnormal@"}
MATH_DROP_NOARG = {"limits", "nolimits", "displaylimits", "nonumber", "notag", "allowbreak", "nobreak",
                   "displaybreak", "mathstrut", "strut", "relax", "protect", "leavevmode", "displaystyle",
                   "textstyle", "scriptstyle", "scriptscriptstyle", "boldmath", "unboldmath", "big", "Big",
                   "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr", "biggl", "biggr", "Biggl", "Biggr", "bigm",
                   "Bigm", "biggm", "Biggm", "middle", "hline", "hdashline", "toprule", "midrule", "bottomrule",
                   "centering", "raggedright", "noindent", "negthinspace", "negmedspace", "negthickspace", "!",
                   "newline", "nolinebreak", "hfill", "hfil", "ignorespaces", "unskip", "xspace", "arraybackslash"}
MATH_KEEP_ARG_ONLY = {"boxed", "fbox", "cancel", "bcancel", "xcancel", "smash", "mathclap", "mathllap",
                      "mathrlap", "clap", "llap", "rlap", "mathop", "mathrel", "mathbin", "mathord", "mathpunct",
                      "mathinner", "ensuremath", "underbracket@", "lefteqn", "phantomsection", "hphantom@",
                      "mathit", "mathnormal", "mathversion@", "operatornamewithlimits", "displaylimits@",
                      "uncover", "only", "visible", "alert", "colorbox@", "highlight", "hl"}
ACCENT_OVERSET = {"check": "ˇ", "breve": "˘", "acute": "´", "grave": "`", "mathring": "˚", "dddot": "⃛",
                  "ddddot": "⃜", "widecheck": "ˇ", "overleftarrow": "←", "overleftrightarrow": "↔",
                  "wideparen": "⌢", "overrightharpoon": "⇀", "overleftharpoon": "↼", "Overrightarrow": "⇒",
                  "overlinesegment": "—", "widebar": "¯", "wideparen@": "⌢"}
ACCENT_UNDERSET = {"underrightarrow": "→", "underleftarrow": "←", "underleftrightarrow": "↔", "utilde": "~",
                   "undertilde": "~", "underparen": "‿", "underlinesegment": "—"}
X_ARROWS = {"xRightarrow": "⟹", "xLeftarrow": "⟸", "xleftrightarrow": "⟷", "xLeftrightarrow": "⟺",
            "xmapsto": "⟼", "xhookrightarrow": "↪", "xhookleftarrow": "↩", "xtwoheadrightarrow": "↠",
            "xtwoheadleftarrow": "↞", "xlongequal": "＝", "xrightleftharpoons": "⇌", "xleftrightharpoons": "⇋",
            "xlongrightarrow": "⟶", "xlongleftarrow": "⟵", "xrightharpoonup": "⇀", "xrightharpoondown": "⇁"}
MATRIX_ENVS = {"matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix", "matrix*", "pmatrix*",
               "bmatrix*", "Bmatrix*", "vmatrix*", "Vmatrix*"}


def _math(self, s: str) -> str:
    s = s or ""
    for lb in re.findall(r"\\label\s*\{([^{}]*)\}", s):
        self.label_values[lb.strip()] = self.current_label
    s = re.sub(r"\\label\s*\{[^{}]*\}", "", s)
    out = self._math_walk(s)
    out = re.sub(r"[ \t\n]+", " ", out).strip()
    return out


def _math_plain_arg(self, a: str) -> str:
    """Raw text for engine commands that print their argument verbatim (\\text, \\mathrm ...)."""
    t = a or ""
    if "$" in t:
        t = re.sub(r"\$([^$]*)\$", lambda m: self._math_to_plain(self._math_walk(m.group(1))), t)
    t = self._plain(t)
    t = t.replace("{", "").replace("}", "").replace("\\", "")
    return t


def _math_simple_symbolic(self, a: str) -> str | None:
    """If arg is only letters/digits/symbol commands → Unicode string, else None."""
    out = []
    i = 0
    n = len(a)
    while i < n:
        c = a[i]
        if c == "\\":
            name, j = read_cmd(a, i)
            if name in MATH_SYMBOLS_ALL and not MATH_SYMBOLS_ALL[name].startswith("\\"):
                out.append(MATH_SYMBOLS_ALL[name])
                i = j
                continue
            if name in (",", ";", " ", ":"):
                out.append(" ")
                i = j
                continue
            return None
        if c in "{}":
            i += 1
            continue
        if c in "_^":
            return None
        out.append(c)
        i += 1
    return "".join(out)


def _math_walk(self, s: str) -> str:
    out: list[str] = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == "\\":
            name, j = read_cmd(s, i)
            r = self._math_cmd(name, s, i, j)
            if r is not None:
                txt, k = r
                out.append(txt)
                i = k
                continue
            out.append(s[i:j])
            if j < n and name.isalpha() and (s[j].isalnum() or s[j] in "\\"):
                out.append(" ")
            i = j
            continue
        if c == "{":
            inner, j = read_group(s, i)
            out.append("{" + self._math_walk(inner) + "}")
            i = j
            continue
        if c == "$":
            i += 1
            continue
        if c == "~":
            out.append("\\ ")
            i += 1
            continue
        if c in "\n\t":
            out.append(" ")
            i += 1
            continue
        if c in LIT_REVERSE or ("\ue000" <= c <= "\uf8ff"):
            i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _math_cmd(self, name: str, s: str, i: int, j: int) -> tuple[str, int] | None:
    n = len(s)
    if not name:
        return "", j
    # control symbols
    if not (name[0].isalpha() or name[0] == "@"):
        if name in ("{", "}", "|", ",", ";", "!", " ", "%", "#", "_", "$", "\\", "quad"):
            if name == "\\":
                k = j
                if k < n and s[k] == "*":
                    k += 1
                kk = skip_spaces_only(s, k)
                if kk < n and s[kk] == "[":
                    _, k = read_bracket(s, kk)
                return " \\\\ ", k
            return "\\" + name, j
        if name == "&":
            return LIT_CHARS["&"], j
        if name in (":", ">"):
            return "\\,", j
        if name == "/":
            return "", j
        if name in ("'", "`", "^", '"', "~", "=", "."):
            a, k = read_arg(s, j)
            return (self._math_walk(a) if a else ""), k
        return "", j
    # text-like
    if name in TEXT_IN_MATH:
        st, k = read_star(s, j)
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        if "$" in a:
            parts = re.split(r"(\$[^$]*\$)", a)
            res = []
            for p in parts:
                if p.startswith("$") and p.endswith("$") and len(p) >= 2:
                    res.append(self._math_walk(p[1:-1]))
                elif p:
                    t = self._math_plain_arg(p)
                    if t:
                        res.append("\\text{" + t + "}")
            return " ".join(res), k
        t = self._math_plain_arg(a)
        if not t.strip():
            return "\\ ", k
        lead = " " if a[:1].isspace() else ""
        trail = " " if a[-1:].isspace() else ""
        t = lead + t.strip() + trail
        cmd = "\\textbf" if name == "textbf" else ("\\textit" if name in ("textit", "emph", "textsl") else "\\text")
        return cmd + "{" + t + "}", k
    if name in ("mathbb", "Bbb", "mathds", "mathbbm", "varmathbb", "mathbbb"):
        a, k = read_arg(s, j)
        t = self._math_plain_arg(a or "")
        return _map_alpha(t, BB_UPPER, BB_LOWER, BB_DIGITS), k
    if name in ("mathcal", "mathscr", "mathcalligra", "pazocal", "mathpzc", "EuScript", "mathscr@"):
        a, k = read_arg(s, j)
        t = self._math_plain_arg(a or "")
        return _map_alpha(t, SCRIPT_UPPER, SCRIPT_LOWER), k
    if name in ("mathfrak", "EuFrak", "frak"):
        a, k = read_arg(s, j)
        t = self._math_plain_arg(a or "")
        return _map_alpha(t, FRAK_UPPER, FRAK_LOWER), k
    if name in ("mathbf", "boldsymbol", "bm", "pmb", "mathbfit", "bmath", "mathbold", "vb", "symbf", "bfseries"):
        a, k = read_arg(s, j)
        if a is None:
            return "", j
        simple = self._math_simple_symbolic(a)
        if simple is not None and simple.strip():
            return "\\mathbf{" + simple + "}", k
        return "{" + self._math_walk(a) + "}", k
    if name in ("mathrm", "mathsf", "mathtt", "mathup", "operatorname", "textup@", "mathsfit", "symrm", "symup",
                "rm@"):
        st, k = read_star(s, j)
        a, k = read_arg(s, k)
        if a is None:
            return "", j
        simple = self._math_simple_symbolic(a)
        cmd = "\\operatorname" if name == "operatorname" else "\\mathrm"
        if simple is not None and simple.strip():
            return cmd + "{" + simple + "}", k
        return "{" + self._math_walk(a) + "}", k
    if name in MATH_OPERATOR_NAMES:
        return "\\operatorname{" + name + "}", j
    if name in MATH_OPERATOR_SPECIAL:
        return "\\operatorname{" + MATH_OPERATOR_SPECIAL[name] + "}", j
    if name in ("left", "right"):
        k = skip_ws(s, j)
        if k < n and s[k] == "\\":
            dn, k2 = read_cmd(s, k)
            d = MATH_SYMBOLS_EXTRA.get(dn)
            if d is None:
                if dn in ("{", "}", "|"):
                    d = "\\" + dn
                elif dn in ("Vert",):
                    d = "‖"
                elif dn == "vert":
                    d = "|"
                elif dn in ("uparrow", "downarrow", "Uparrow", "Downarrow", "updownarrow"):
                    d = MATH_SYMBOLS_ALL.get(dn, "")
                else:
                    d = "\\" + dn
            return "\\" + name + d + " ", k2
        if k < n:
            ch = s[k]
            if ch == "<":
                ch = "⟨"
            elif ch == ">":
                ch = "⟩"
            return "\\" + name + ch + " ", k + 1
        return "", k
    if name in MATH_DROP_NOARG:
        return " ", j
    if name in ("quad", "qquad"):
        return "\\" + name + " ", j
    if name in ("enspace", "thinspace", "medspace", "thickspace", "enskip", "space", "nobreakspace"):
        return "\\," if name in ("thinspace", "medspace") else "\\ ", j
    if name in ("hspace", "hskip", "kern", "mkern", "mspace", "hphantom", "phantom", "mskip"):
        st, k = read_star(s, j)
        if name in ("hskip", "kern", "mkern", "mskip"):
            m = re.match(r"\s*=?\s*-?[\d.]+\s*(cm|mm|pt|em|ex|in|mu)?", s[k:])
            k += m.end() if m else 0
            return "\\ ", k
        a, k = read_arg(s, k)
        if name in ("phantom", "hphantom"):
            return "\\ ", k
        tw = dim_to_twips((a or "").replace("mu", "pt")) or 0
        return ("\\quad " if tw >= 200 else "\\ "), k
    if name in ("vphantom", "vspace", "color", "label", "tag@", "nocite", "index"):
        if name == "color":
            o, k = read_opt(s, j, allow_newline=False)
            a, k = read_braced(s, k)
            return "", k
        st, k = read_star(s, j)
        a, k = read_arg(s, k)
        return "", k
    if name in ("textcolor", "colorbox"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_braced(s, k)
        b, k = read_arg(s, k)
        return "{" + self._math_walk(b or "") + "}", k
    if name == "fcolorbox":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_braced(s, k)
        a, k = read_braced(s, k)
        b, k = read_arg(s, k)
        return "{" + self._math_walk(b or "") + "}", k
    if name in MATH_KEEP_ARG_ONLY:
        st, k = read_star(s, j)
        o, k2 = read_opt(s, k, allow_newline=False)
        if o is not None:
            k = k2
        a, k = read_arg(s, k)
        return "{" + self._math_walk(a or "") + "}", k
    if name == "cancelto":
        a, k = read_arg(s, j)
        b, k = read_arg(s, k)
        return "{" + self._math_walk(b or "") + "}", k
    if name == "stackrel":
        a, k = read_arg(s, j)
        b, k = read_arg(s, k)
        return "\\overset{" + self._math_walk(a or "") + "}{" + self._math_walk(b or "") + "}", k
    if name in ("overbrace", "underbrace", "overbracket", "underbracket"):
        a, k = read_arg(s, j)
        body = self._math_walk(a or "")
        over = name.startswith("over")
        ch = "⏞" if name == "overbrace" else ("⏟" if name == "underbrace" else ("⎴" if over else "⎵"))
        core = ("\\overset{" if over else "\\underset{") + ch + "}{" + body + "}"
        kk = skip_ws(s, k)
        if kk < n and s[kk] == ("^" if over else "_"):
            lab, k = read_arg(s, kk + 1)
            core = ("\\overset{" if over else "\\underset{") + self._math_walk(lab or "") + "}{" + core + "}"
        return core, k
    if name in ACCENT_OVERSET:
        a, k = read_arg(s, j)
        return "\\overset{" + ACCENT_OVERSET[name] + "}{" + self._math_walk(a or "") + "}", k
    if name in ACCENT_UNDERSET:
        a, k = read_arg(s, j)
        return "\\underset{" + ACCENT_UNDERSET[name] + "}{" + self._math_walk(a or "") + "}", k
    if name in X_ARROWS:
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        core = "\\overset{" + self._math_walk(a or "") + "}{" + X_ARROWS[name] + "}"
        if o:
            core = "\\underset{" + self._math_walk(o) + "}{" + core + "}"
        return core, k
    if name in ("xrightarrow", "xleftarrow"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        res = "\\" + name
        if o is not None:
            res += "[" + self._math_walk(o) + "]"
        return res + "{" + self._math_walk(a or "") + "}", k
    if name in ("substack", "subarray"):
        if name == "subarray":
            _, j = read_arg(s, j)
        a, k = read_arg(s, j)
        return "\\begin{matrix}" + self._math_walk(a or "") + "\\end{matrix}", k
    if name == "pmod":
        a, k = read_arg(s, j)
        return "\\quad (\\operatorname{mod}\\ " + self._math_walk(a or "") + ")", k
    if name == "pod":
        a, k = read_arg(s, j)
        return "\\quad (" + self._math_walk(a or "") + ")", k
    if name == "bmod":
        return "\\ \\operatorname{mod}\\ ", j
    if name == "mod":
        return "\\quad \\operatorname{mod}\\ ", j
    if name == "not":
        k = skip_ws(s, j)
        if k < n and s[k] == "\\":
            dn, k2 = read_cmd(s, k)
            if dn in NOT_MAP:
                return NOT_MAP[dn], k2
            base = MATH_SYMBOLS_ALL.get(dn, dn)
            return base + "\u0338", k2
        if k < n:
            if s[k] in NOT_MAP:
                return NOT_MAP[s[k]], k + 1
            return s[k] + "\u0338", k + 1
        return "", k
    if name in ("nicefrac", "sfrac", "flatfrac", "slantfrac", "xfrac@"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        b, k = read_arg(s, k)
        return "{" + self._math_walk(a or "") + "}/{" + self._math_walk(b or "") + "}", k
    if name == "genfrac":
        args = []
        k = j
        for _ in range(6):
            a, k = read_arg(s, k)
            args.append(a or "")
        return "\\frac{" + self._math_walk(args[4]) + "}{" + self._math_walk(args[5]) + "}", k
    if name == "cfrac":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        b, k = read_arg(s, k)
        return "\\cfrac{" + self._math_walk(a or "") + "}{" + self._math_walk(b or "") + "}", k
    if name == "root":
        m = re.match(r"\s*(.+?)\\of\s*", s[j:])
        if m:
            deg = m.group(1)
            a, k = read_arg(s, j + m.end())
            return "\\sqrt[" + self._math_walk(deg) + "]{" + self._math_walk(a or "") + "}", k
    if name == "sqrt":
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        res = "\\sqrt"
        if o is not None:
            res += "[" + self._math_walk(o) + "]"
        return res + "{" + self._math_walk(a or "") + "}", k
    if name in ("ce", "cee"):
        a, k = read_arg(s, j)
        return self._chem_to_math(a or ""), k
    if name in ("SI", "qty", "si", "unit", "num", "ang", "SIrange", "numrange", "qtyrange"):
        return self._siunitx(name, s, j)
    if name == "begin":
        a, k = read_braced(s, j)
        if a is None:
            return "", j
        env = a.strip()
        ce, after = find_env_end(s, k, env)
        body = s[k:ce]
        return self._math_env(env, body), after
    if name == "end":
        a, k = read_braced(s, j)
        return "", k
    if name in ("tag",):
        st, k = read_star(s, j)
        a, k = read_arg(s, k)
        return "\\tag{" + self._plain(a or "") + "}", k
    if name in ("heva", "hoac"):
        o, k = read_opt(s, j, allow_newline=False)
        a, k = read_arg(s, k)
        return "\\" + name + "{" + self._math_walk(a or "") + "}", k
    if name in ("dots", "cdots", "ldots", "vdots", "ddots"):
        return "\\" + name + " ", j
    if name in MATH_SYMBOLS_EXTRA and name not in ENGINE_MATH_SYMBOL_NAMES:
        v = MATH_SYMBOLS_EXTRA[name]
        return v + (" " if v.startswith("\\") else ""), j
    if name in ENGINE_MATH_SYMBOL_NAMES or name in ENGINE_MATH_CMDS:
        return None  # keep as is (handled by omml_math)
    if name in ("vv", "overrightarrow@"):
        a, k = read_arg(s, j)
        return "\\overrightarrow{" + self._math_walk(a or "") + "}", k
    if name in ("overarc", "arc", "widearc", "overparen"):
        a, k = read_arg(s, j)
        return "\\overset{⌢}{" + self._math_walk(a or "") + "}", k
    if name in ("intertext", "shortintertext"):
        a, k = read_arg(s, j)
        return "\\text{" + self._math_plain_arg(a or "") + "}", k
    if name.startswith("the") and name[3:] in self.counters:
        return self.the(name[3:]), j
    if name in ("hdots",):
        return "…", j
    if name in TEXT_SYMBOLS and TEXT_SYMBOLS[name] and not TEXT_SYMBOLS[name][0] in "\ue000\ue010\ue011":
        v = TEXT_SYMBOLS[name]
        return (v if v.strip() else " "), j
    # unknown → keep args, or print name upright
    self._unknown(name)
    k = j
    st, k = read_star(s, k)
    pieces = []
    while True:
        kk = skip_spaces_only(s, k)
        if kk < n and s[kk] == "[":
            _, k = read_bracket(s, kk)
            continue
        if kk < n and s[kk] == "{":
            a, k = read_group(s, kk)
            pieces.append("{" + self._math_walk(a) + "}")
            continue
        break
    if pieces:
        return "".join(pieces), k
    return "\\operatorname{" + name + "}", k


def _math_env(self, env: str, body: str) -> str:
    base = env.rstrip("*")
    body2 = body
    if base in ("matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix"):
        body2 = re.sub(r"^\s*\[[^\]]*\]", "", body2)
        return "\\begin{" + base + "}" + self._math_walk(self._strip_hlines(body2)) + "\\end{" + base + "}"
    if base in ("smallmatrix", "psmallmatrix", "bsmallmatrix", "Bsmallmatrix", "vsmallmatrix", "Vsmallmatrix"):
        nb = {"smallmatrix": "matrix", "psmallmatrix": "pmatrix", "bsmallmatrix": "bmatrix", "Bsmallmatrix": "Bmatrix",
              "vsmallmatrix": "vmatrix", "Vsmallmatrix": "Vmatrix"}[base]
        return "\\begin{" + nb + "}" + self._math_walk(self._strip_hlines(body2)) + "\\end{" + nb + "}"
    if base == "array":
        body2 = re.sub(r"^\s*\[[^\]]*\]", "", body2)
        m = re.match(r"^\s*\{", body2)
        spec = ""
        if m:
            spec, e = read_group(body2, m.end() - 1)
            body2 = body2[e:]
        spec_simple = re.sub(r"[^lcr]", "", re.sub(r"\{[^{}]*\}", "", spec)) or "c"
        return "\\begin{array}{" + spec_simple + "}" + self._math_walk(self._strip_hlines(body2)) + "\\end{array}"
    if base in ("cases", "dcases", "numcases", "subnumcases", "cases*", "dcases*"):
        if base in ("numcases", "subnumcases"):
            _, e = read_arg(body2, 0)
            body2 = body2[e:]
        return "\\begin{cases}" + self._math_walk(body2) + "\\end{cases}"
    if base == "rcases":
        return "\\left.\\begin{aligned}" + self._math_walk(body2) + "\\end{aligned}\\right\\}"
    if base in ("aligned", "alignedat", "gathered", "split", "multlined", "align", "gather", "flalign",
                "alignat", "eqnarray", "equation", "multline", "lgathered", "rgathered", "dmath", "darray"):
        body2 = re.sub(r"^\s*\[[^\]]*\]", "", body2)
        if base in ("alignedat", "alignat"):
            body2 = re.sub(r"^\s*\{[^{}]*\}", "", body2)
        body2 = re.sub(r"\\(nonumber|notag)\b|\\tag\*?\{[^{}]*\}", "", body2)
        if base == "eqnarray":
            body2 = body2.replace("&=&", "&=")
        return "\\begin{aligned}" + self._math_walk(body2) + "\\end{aligned}"
    if base in ("subarray",):
        body2 = re.sub(r"^\s*\{[^{}]*\}", "", body2)
        return "\\begin{matrix}" + self._math_walk(body2) + "\\end{matrix}"
    # unknown env inside math → content
    return self._math_walk(body2)


def _strip_hlines(self, body: str) -> str:
    return re.sub(r"\\(hline|hdashline|toprule|midrule|bottomrule)\b|\\cline\s*\{[^}]*\}|\\cdashline\s*\{[^}]*\}", "", body)


def _math_to_plain(self, m: str) -> str:
    t = m
    t = re.sub(r"\\(?:text|mathrm|operatorname|mathbf|textbf|textit)\{([^{}]*)\}", r"\1", t)
    t = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", r"\1/\2", t)
    for k2, v in sorted(MATH_SYMBOLS_ALL.items(), key=lambda x: -len(x[0])):
        if not v.startswith("\\"):
            t = t.replace("\\" + k2, v)
    t = re.sub(r"\^\{?2\}?", "²", t)
    t = re.sub(r"\^\{?3\}?", "³", t)
    t = re.sub(r"\\[A-Za-z]+", "", t)
    t = t.replace("{", "").replace("}", "").replace("\\", "")
    return t.strip()


# ---------------- chemistry (mhchem) ----------------
def _chem_to_math(self, src: str) -> str:
    s = src.strip()
    arrows = [("<=>>", "⇌"), ("<<=>", "⇌"), ("<=>", "⇌"), ("<->", "↔"), ("->", "→"), ("<-", "←")]
    out: list[str] = []
    i = 0
    n = len(s)
    prev_elem = False
    while i < n:
        c = s[i]
        matched = False
        for a, sym in arrows:
            if s.startswith(a, i):
                i += len(a)
                lab = ""
                if i < n and s[i] == "[":
                    lab, i = read_bracket(s, i)
                    if i < n and s[i] == "[":
                        _, i = read_bracket(s, i)
                if lab:
                    out.append(" \\xrightarrow{" + self._chem_to_math(lab) + "} ")
                else:
                    out.append(" " + sym + " ")
                prev_elem = False
                matched = True
                break
        if matched:
            continue
        if c.isupper():
            j = i + 1
            while j < n and s[j].islower():
                j += 1
            out.append("\\mathrm{" + s[i:j] + "}")
            i = j
            prev_elem = True
            continue
        if c.isdigit():
            j = i
            while j < n and (s[j].isdigit() or s[j] in "./"):
                j += 1
            num = s[i:j]
            if prev_elem:
                out.append("_{" + num + "}")
            else:
                out.append(num)
            i = j
            continue
        if c == "^":
            if i + 1 < n and s[i + 1] == "{":
                a, j = read_group(s, i + 1)
            else:
                m = re.match(r"\^(\d*[+-]|\d+)", s[i:])
                a = m.group(1) if m else ""
                j = i + (m.end() if m else 1)
            a = a.replace("-", "−")
            out.append("^{" + a + "}")
            i = j
            prev_elem = False
            continue
        if c == "_" and i + 1 < n:
            if s[i + 1] == "{":
                a, j = read_group(s, i + 1)
            else:
                a, j = s[i + 1], i + 2
            out.append("_{" + a + "}")
            i = j
            continue
        if c in "+-" and prev_elem and (i + 1 >= n or s[i + 1] in " )]"):
            out.append("^{" + ("+" if c == "+" else "−") + "}")
            i += 1
            continue
        if c == "(" and re.match(r"\((s|l|g|aq|r|k|dd|đ|dd)\)", s[i:]):
            m = re.match(r"\(([a-z]+)\)", s[i:])
            out.append("_{(\\mathrm{" + m.group(1) + "})}")
            i += m.end()
            continue
        if c in "()[]":
            out.append(c)
            prev_elem = c in ")]"
            i += 1
            continue
        if c in ".*" and prev_elem:
            out.append("⋅")
            i += 1
            prev_elem = False
            continue
        if c == "$":
            e = s.find("$", i + 1)
            if e == -1:
                e = n
            out.append(self._math_walk(s[i + 1 : e]))
            i = e + 1
            continue
        if c == "\\":
            name, j = read_cmd(s, i)
            out.append("\\" + name + " ")
            i = j
            continue
        if c == "^" or c == "v" and (i + 1 >= n or s[i + 1] == " ") and (i == 0 or s[i - 1] == " "):
            out.append("↓" if c == "v" else "↑")
            i += 1
            continue
        if c == " ":
            out.append(" ")
            prev_elem = False
            i += 1
            continue
        if c == "+":
            out.append(" + ")
            prev_elem = False
            i += 1
            continue
        if c == "=":
            out.append(" = ")
            prev_elem = False
            i += 1
            continue
        out.append(c)
        prev_elem = False
        i += 1
    return "".join(out).strip()


# ---------------- siunitx ----------------
def _si_unit(self, u: str) -> str:
    u = u.strip()
    if not u:
        return ""
    if "\\" not in u:
        t = u.replace(".", "⋅").replace("~", " ")
        t = re.sub(r"\^\{?(-?\d+)\}?", lambda m: "^{" + m.group(1).replace("-", "−") + "}", t)
        t = re.sub(r"([A-Za-zµΩ°]+)", r"\\mathrm{\1}", t)
        return t
    parts = []
    tokens = re.findall(r"\\([A-Za-z]+)|([^\\\s]+)", u)
    pending_per = False
    pending_prefix = ""
    for name, other in tokens:
        if name:
            if name == "per":
                pending_per = True
                continue
            if name in ("square", "squared"):
                if name == "squared" and parts:
                    parts[-1] += "^{2}"
                else:
                    pending_prefix += "^2"
                continue
            if name in ("cubic", "cubed"):
                if name == "cubed" and parts:
                    parts[-1] += "^{3}"
                else:
                    pending_prefix += "^3"
                continue
            if name in ("centi", "milli", "micro", "nano", "pico", "kilo", "mega", "giga", "tera", "deci", "deca",
                        "hecto", "femto"):
                pending_prefix = SI_UNITS[name] + pending_prefix
                continue
            sym = SI_UNITS.get(name, name)
            exp = ""
            if pending_prefix.endswith("^2"):
                exp = "^{2}"
                pending_prefix = pending_prefix[:-2]
            if pending_prefix.endswith("^3"):
                exp = "^{3}"
                pending_prefix = pending_prefix[:-2]
            piece = "\\mathrm{" + pending_prefix + sym + "}" + exp
            pending_prefix = ""
            if pending_per:
                piece = "/" + piece
                pending_per = False
            parts.append(piece)
        elif other:
            parts.append("\\mathrm{" + other + "}")
    return "⋅".join(parts).replace("⋅/", "/")


def _si_num(self, v: str) -> str:
    v = v.strip()
    m = re.match(r"^([-+]?[\d.,]+)\s*[eE]\s*([-+]?\d+)$", v)
    if m:
        return m.group(1) + "\\times 10^{" + m.group(2) + "}"
    v = v.replace("x", "\\times ").replace("+-", "\\pm ")
    return v


def _siunitx(self, name: str, s: str, j: int) -> tuple[str, int]:
    o, k = read_opt(s, j, allow_newline=False)
    if name in ("SI", "qty"):
        a, k = read_braced(s, k)
        o2, k2 = read_opt(s, k, allow_newline=False)
        if o2 is not None:
            k = k2
        b, k = read_braced(s, k)
        unit = self._si_unit(b or "")
        sep = "" if (b or "").strip() in ("\\percent", "\\degree", "%", "\\arcminute", "\\arcsecond") else "\\ "
        return self._si_num(a or "") + sep + unit, k
    if name in ("si", "unit"):
        b, k = read_braced(s, k)
        return self._si_unit(b or ""), k
    if name in ("num", "tablenum"):
        a, k = read_braced(s, k)
        return self._si_num(a or ""), k
    if name == "ang":
        a, k = read_braced(s, k)
        parts = (a or "").split(";")
        syms = ["°", "′", "″"]
        return "".join(p.strip() + syms[i2] for i2, p in enumerate(parts) if p.strip() and i2 < 3), k
    if name in ("SIrange", "qtyrange", "numrange"):
        a, k = read_braced(s, k)
        b, k = read_braced(s, k)
        unit = ""
        if name != "numrange":
            c, k = read_braced(s, k)
            unit = "\\ " + self._si_unit(c or "")
        return self._si_num(a or "") + unit + "\\text{ – }" + self._si_num(b or "") + unit, k
    a, k = read_braced(s, k)
    return self._si_num(a or ""), k


# =====================================================================================
# BLOCK HANDLERS
# =====================================================================================
def _set_pending(self, ctx: _Ctx, text: str, spec: dict[str, Any] | None = None) -> None:
    if ctx.pending is not None and not ctx.pending.used:
        ctx.pending.text += text
        ctx.pending.spec.update(spec or {})
        return
    ctx.pending = _Pending(text=text, spec=dict(spec or {}), used=False)


def _h_par(self, s, i, j, ctx, buf):
    self._flush(buf, ctx)
    return j


def _h_skip(self, s, i, j, ctx, buf):
    return _eat_space_after_word(s, j)


MINIPAGE_GAP_RE = re.compile(r"(?:\s|\\hfill\b|\\hspace\*?\s*\{[^}]*\}|\\quad\b|\\qquad\b|~|\\vline\b|\\noindent\b|\\centering\b)*")


def _h_begin(self, s, i, j, ctx, buf):
    k = skip_spaces_only(s, j)
    if k >= len(s) or s[k] != "{":
        return None
    env, k2 = read_group(s, k)
    env = env.strip()
    ce, after = find_env_end(s, k2, env)
    content = s[k2:ce]
    if env == "minipage" and not getattr(self, "_config_only", False):
        group = [content]
        pos = after
        while True:
            m = MINIPAGE_GAP_RE.match(s, pos)
            q = m.end() if m else pos
            if s.startswith("\\begin{minipage}", q):
                ce2, after2 = find_env_end(s, q + len("\\begin{minipage}"), "minipage")
                group.append(s[q + len("\\begin{minipage}") : ce2])
                pos = after2
                continue
            break
        if len(group) >= 2:
            self._flush(buf, ctx)
            if self._minipage_row(group, ctx):
                return pos
            for g in group:
                self._env("minipage", g, ctx, buf)
            return pos
    self._env(env, content, ctx, buf)
    return after


def _minipage_row(self, group: list[str], ctx: _Ctx) -> bool:
    parts = []
    for g in group:
        (o1, o2, o3, w), body = _strip_env_args(g, "ooom")
        parts.append((dim_to_twips(w or "", self.textwidth), body))
    tikz_idx = [n for n, (_, b) in enumerate(parts) if "\\begin{tikzpicture}" in b]
    blocky = re.compile(r"\\begin\{(?!tikzpicture|center|minipage)[a-zA-Z*]+\}|\\\[|\$\$|\\choice|\\includegraphics|\\section|\\item")
    if len(parts) == 2 and len(tikz_idx) == 1:
        t = tikz_idx[0]
        text_body = parts[1 - t][1]
        fig_body = parts[t][1]
        if "\\begin{tikzpicture}" not in text_body:
            synthetic = "\\immini{" + text_body + "}{" + fig_body + "}"
            r = self._BLOCK_HANDLERS["immini"](self, synthetic, 0, len("\\immini"), ctx, [])
            return r is not None
    if tikz_idx or any(blocky.search(b) for _, b in parts):
        return False
    specs = []
    cells = []
    for w, b in parts:
        b2 = b.strip()
        center = b2.startswith("\\centering")
        if center:
            b2 = b2[len("\\centering"):]
        width = w or int(self.textwidth / len(parts))
        specs.append((">{\\centering}" if center else "") + "p{" + f"{width / 1440 * 2.54:.2f}cm" + "}")
        b2 = re.sub(r"\n[ \t]*\n+", " \\\\newline ", b2.strip())
        b2 = " \\newline ".join(r for r in self._split_rows(b2) if r.strip())
        cells.append(b2)
    tab = "{" + "".join(specs) + "}\n" + " & ".join(cells) + "\n"
    self._table("tabular", tab, ctx)
    self._stat("minipage_rows")
    return True


def _h_end(self, s, i, j, ctx, buf):
    k = skip_spaces_only(s, j)
    if k < len(s) and s[k] == "{":
        _, k2 = read_group(s, k)
        return k2
    return j


HEADING_LEVELS = {"part": 0, "chapter": 0, "section": 1, "subsection": 2, "subsubsection": 3, "paragraph": 4,
                  "subparagraph": 5, "addsec": 1, "addchap": 0, "minisec": 4}
HEADING_SIZES = {0: 24.88, 1: 17.28, 2: 14.4, 3: 12.0}
HEADING_SPACING = {0: (24, 18), 1: (14, 8), 2: (12, 6), 3: (10, 4)}


def _h_heading(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    st, k = read_star(s, j)
    short, k = read_opt(s, k, allow_newline=False)
    title, k = read_arg(s, k)
    if title is None:
        return None
    self._flush(buf, ctx)
    if getattr(self, "_config_only", False):
        return k
    level = HEADING_LEVELS[name]
    if name in ("chapter", "addchap") and self.doc_class not in ("book", "report"):
        name, level = "section", 1
    numbered = not st and name not in ("addsec", "addchap", "minisec")
    depth_map = {"part": -1, "chapter": 0, "section": 1, "subsection": 2, "subsubsection": 3, "paragraph": 4,
                 "subparagraph": 5}
    num = ""
    counter = {"addsec": "section", "addchap": "chapter", "minisec": "paragraph"}.get(name, name)
    if numbered and depth_map.get(counter, 9) <= self.secnumdepth:
        num = self.refstep(counter, counter)
    elif numbered:
        self.step(counter)
    title_txt = self._inline(title, _ICtx(bold=True)).strip()
    if level >= 4:
        lead = (num + "\u2003") if num else ""
        prefix = self._wrap_fmt(_ICtx(), "b", lead + title_txt) + "\u2003"
        self._set_pending(ctx, prefix, {"sb": 8})
        return _eat_space_after_word(s, k)
    plain_title = self._plain(title)
    if name in ("chapter", "part", "addchap"):
        word = ("Chương" if name == "chapter" else "Phần") if self.lang == "vi" else ("Chapter" if name == "chapter" else "Part")
        if num:
            self._emit_para(self.span({"b": 1, "sz": 20.74}, f"{word} {num}"),
                            {"heading": 1, "keep": 1, "sb": 24, "sa": 12, "align": "c" if name == "part" else "l"}, ctx.child())
        self._emit_para(self.span({"b": 1, "sz": 24.88}, title_txt),
                        {"heading": 1, "keep": 1, "sa": 18, "sb": 0 if num else 24, "align": "c" if name == "part" else "l"}, ctx.child())
        self.headings.append((0, (f"{word} {num}. " if num else "") + plain_title))
        return k
    lead = (num + "\u2003") if num else ""
    sb, sa = HEADING_SPACING.get(level, (10, 4))
    if self.doc_class.startswith("ams") and level == 1:
        self._emit_para(self.span({"sc": 1}, ((num + ". ") if num else "") + self._inline(title, _ICtx())),
                        {"heading": level, "keep": 1, "sb": 12, "sa": 6, "align": "c"}, ctx.child(left=ctx.left))
    elif self.doc_class.startswith("ams") and level >= 2:
        prefix = self._wrap_fmt(_ICtx(), "b", ((num + ". ") if num else "") + title_txt + ".") + " "
        self._set_pending(ctx, prefix, {"sb": 8})
        self.headings.append((level, (num + " " if num else "") + plain_title))
        return k
    else:
        self._emit_para(self.span({"b": 1, "sz": HEADING_SIZES.get(level, 12.0)}, lit_safe(lead) + title_txt),
                        {"heading": level, "keep": 1, "sb": sb, "sa": sa}, ctx.child(align=ctx.align, left=ctx.left))
    self.headings.append((level, (num + " " if num else "") + plain_title))
    self._stat("headings")
    return k


def lit_safe(s: str) -> str:
    return s


def _h_frametitle(self, s, i, j, ctx, buf):
    o, k = read_opt(s, j, allow_newline=False)
    a, k = read_arg(s, k)
    if a is None:
        return None
    self._flush(buf, ctx)
    self._emit_para(self.span({"b": 1, "sz": 14.4, "c": "1F3A93"}, self._inline(a, _ICtx())), {"keep": 1, "sb": 6, "sa": 6}, ctx.child())
    return k


def _h_maketitle(self, s, i, j, ctx, buf):
    self._flush(buf, ctx)
    if getattr(self, "_config_only", False):
        return j
    sub = ctx.child(align="c")
    ams = self.doc_class.startswith("ams")
    if self.title:
        spec_t = {"b": 1, "caps": 1, "sz": 12.0} if ams else {"sz": 17.28}
        self._emit_para(self.span(spec_t, self._inline(self.title, _ICtx())), {"align": "c", "sb": 12, "sa": 12}, sub)
    if self.author:
        authors = re.split(r"\\and\b", self.author)
        if ams:
            txt = (" and " if self.lang != "vi" else " và ").join(self._inline(a.strip(), _ICtx()) for a in authors)
            self._emit_para(self.span({"caps": 1, "sz": 10.0}, txt), {"align": "c", "sa": 12}, sub)
        else:
            txt = "\u2003\u2003".join(self._inline(a.strip(), _ICtx()) for a in authors)
            self._emit_para(self.span({"sz": 14.4}, txt), {"align": "c", "sa": 6}, sub)
    date = self.date if self.date is not None else ("" if ams else "\\today")
    if date.strip():
        self._emit_para(self.span({"sz": 14.4}, self._inline(date, _ICtx())), {"align": "c", "sa": 18}, sub)
    self._title_done = True
    ab = getattr(self, "_deferred_abstract", None)
    if ab is not None:
        self._deferred_abstract = None
        title = self.names.get("abstractname") or ("Tóm tắt" if self.lang == "vi" else "Abstract")
        sub2 = ctx.child(left=ctx.left + 851, spans=ctx.spans + (self.span_id({"sz": 10.0}),))
        sub2.pending = _Pending(text=self.span({"sc": 1}, title + ".") + " ", used=False)
        self._blocks(ab, sub2)
    self._stat("title_block")
    return _eat_space_after_word(s, j)


def _h_toc(self, s, i, j, ctx, buf):
    self._flush(buf, ctx)
    if getattr(self, "_config_only", False):
        return j
    title = self.names.get("contentsname") or ("Mục lục" if self.lang == "vi" else "Contents")
    self._emit_para(self.span({"b": 1, "sz": 17.28}, title), {"keep": 1, "sb": 12, "sa": 8}, ctx.child())
    self.items.append(("para", "", {"toc": 1}))
    self.page["toc"] = True
    return _eat_space_after_word(s, j)


def _h_newpage(self, s, i, j, ctx, buf):
    self._flush(buf, ctx)
    name = read_cmd(s, i)[0]
    k = j
    if name in ("pagebreak", "newpage@"):
        _, k = read_opt(s, k, allow_newline=False)
    if not getattr(self, "_config_only", False):
        self.items.append(("para", "", {"pb": 1}))
    return _eat_space_after_word(s, k)


def _h_vspace(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    k = j
    pts = 0.0
    if name in ("vspace", "addvspace"):
        st, k = read_star(s, k)
        a, k = read_arg(s, k)
        tw = dim_to_twips(a or "") or 0
        pts = tw / 20.0
    elif name == "vskip":
        m = re.match(r"\s*(-?[\d.]+\s*(?:cm|mm|pt|em|ex|in))(?:\s*plus\s*[-\d.]+\s*\w+)?(?:\s*minus\s*[-\d.]+\s*\w+)?", s[k:])
        if m:
            pts = (dim_to_twips(m.group(1)) or 0) / 20.0
            k += m.end()
    else:
        pts = {"smallskip": 3.0, "medskip": 6.0, "bigskip": 12.0, "vfill": 0.0}.get(name, 0.0)
    if buf_has_text_local(buf):
        self._flush(buf, ctx)
    if pts > 0 and not getattr(self, "_config_only", False):
        self._pending_vspace = min(200, getattr(self, "_pending_vspace", 0) + int(round(pts)))
    return _eat_space_after_word(s, k)


def buf_has_text_local(buf: list[str]) -> bool:
    return any(x.strip() for x in buf)


def _h_label(self, s, i, j, ctx, buf):
    a, k = read_braced(s, j)
    if a is None:
        return None
    self.label_values[a.strip()] = self.current_label
    return k


def _h_caption(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    k = j
    kind = getattr(ctx, "float_kind", None) or "figure"
    if name == "captionof":
        st, k = read_star(s, k)
        a, k = read_braced(s, k)
        kind = (a or "figure").strip()
    st, k = read_star(s, k)
    short, k = read_opt(s, k, allow_newline=False)
    text, k = read_arg(s, k)
    if text is None:
        return None
    self._flush(buf, ctx)
    if getattr(self, "_config_only", False):
        return k
    kind_base = "table" if "table" in kind else "figure"
    if st or name == "caption*":
        label = ""
    else:
        num = self.refstep(kind_base, kind_base)
        nm = self.names.get(kind_base + "name") or (
            ("Bảng" if kind_base == "table" else "Hình") if self.lang == "vi" else ("Table" if kind_base == "table" else "Figure"))
        label = f"{nm} {num}: "
    body = self._inline(text, _ICtx())
    self._emit_para(label + body, {"align": "c", "sb": 4, "sa": 8, "keep": 1 if kind_base == "table" else 0}, ctx.child())
    self._stat("captions")
    return k


def _h_includegraphics(self, s, i, j, ctx, buf):
    st, k = read_star(s, j)
    o, k = read_opt(s, k, allow_newline=False)
    o2, k2 = read_opt(s, k, allow_newline=False)
    if o2 is not None:
        k = k2
    a, k = read_braced(s, k)
    if a is None:
        return None
    self._flush(buf, ctx)
    path = a.strip()
    found = None
    if self.base_dir is not None:
        for ext in ("", ".png", ".jpg", ".jpeg", ".gif", ".bmp"):
            p = self.base_dir / (path + ext)
            try:
                if p.is_file() and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
                    found = path + ext
                    break
            except OSError:
                pass
    if found:
        self._emit_block("\\includegraphics{" + found + "}", ctx)
        self._stat("images")
    else:
        self.res.warnings.append(f"Không có file ảnh '{path}' đi kèm — đã chèn khung giữ chỗ để bạn thay ảnh trong Word.")
        self._emit_para(self.span(PLACEHOLDER_IMG_SPAN, f"[Hình ảnh: {lit(path)}]"),
                        {"align": "c", "box": {"style": "dashed", "color": "94A3B8"}, "sb": 6, "sa": 6}, ctx.child())
        self._stat("image_placeholders")
    return k


def _choice_options_render(self, opts: list[str]) -> tuple[list[str], int | None]:
    rendered = []
    correct = None
    for idx, o in enumerate(opts):
        if re.search(r"\\True\b|\\true\b|\\dung\b", o):
            correct = idx
            o = re.sub(r"\\(True|true|dung)\b", "", o)
        rendered.append(self._inline(o.strip(), _ICtx(in_cell=True)).strip())
    return rendered, correct


def _emit_choice(self, opts: list[str], correct: int | None, force: int | None, ctx: _Ctx) -> None:
    if len(opts) == 4:
        cmd = {4: "boncot", 2: "haicot", 1: "motcot"}.get(force or 0, "choice")
        parts = []
        for idx, o in enumerate(opts):
            t = o.strip() or "\u00a0"
            if correct == idx:
                t = "\\True " + t
            parts.append("{" + t + "}")
        self._emit_block("\\" + cmd + "".join(parts), ctx)
        self._stat("choice_tables")
        return
    # other counts → tab-aligned rows
    labels = [chr(65 + k) + "." for k in range(len(opts))]
    cols = force or (len(opts) if max((len(self._plain(o)) for o in opts), default=0) < 14 else 2)
    self._options_rows(labels, opts, cols, ctx, correct)


def _options_rows(self, labels: list[str], opts: list[str], cols: int, ctx: _Ctx, correct: int | None = None) -> None:
    cols = max(1, min(cols, 6))
    width = self.textwidth - ctx.left
    stops = [ctx.left + int(width * c / cols) for c in range(1, cols)]
    for r in range(0, len(opts), cols):
        cells = []
        for c in range(cols):
            idx = r + c
            if idx >= len(opts):
                break
            lbl = "\\textbf{" + labels[idx] + "}"
            if correct == idx:
                lbl = "\\underline{" + lbl + "}"
            cells.append(lbl + " " + opts[idx])
        self._emit_para(MK_TAB.join(cells), {"tabs": stops, "left": ctx.left, "sb": 1, "sa": 1}, ctx.child())


def _h_choice(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    force = None
    if o is not None:
        k = k2
        if o.strip() in ("1", "2", "4"):
            force = int(o.strip())
    st, k = read_star(s, k)
    opts, k3 = self._read_n_braced(s, k, 4)
    if len(opts) < 2:
        return None
    self._flush(buf, ctx)
    if name in ("choiceA", "choiceB", "choiceC", "choiceD"):
        correct_fixed = "ABCD".index(name[-1])
    else:
        correct_fixed = None
    rendered, correct = self._choice_options_render(opts)
    if correct_fixed is not None:
        correct = correct_fixed
    if force is None:
        force = {"boncot": 4, "haicot": 2, "motcot": 1}.get(name)
        if force is None and name not in ("choice", "bonpa", "choiceA", "choiceB", "choiceC", "choiceD"):
            if re.search(r"(bon|four|4)", name):
                force = None
            elif re.search(r"(hai|two|2)", name):
                force = 2
            elif re.search(r"(mot|one|1)", name):
                force = 1
    if ctx.in_immini:
        force = force or None
    self._emit_choice(rendered, correct, force, ctx)
    return k3


def _h_choice_tf(self, s, i, j, ctx, buf):
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    if o is not None:
        k = k2
    args, k3 = self._read_n_braced(s, k, 8)
    if len(args) < 2:
        return None
    self._flush(buf, ctx)
    is8 = len(args) == 8 and all(len(self._plain(a)) <= 6 for a in args[1::2]) and any(
        self._plain(a).strip().upper() in ("Đ", "D", "S", "TRUE", "FALSE", "0", "1", "") for a in args[1::2])
    if is8:
        stmts = [(args[2 * q], args[2 * q + 1]) for q in range(4)]
        endk = k3
    else:
        args4, endk = self._read_n_braced(s, k, 4)
        has_true = any(re.search(r"\\True\b", a) for a in args4)
        stmts = []
        for a in args4:
            t = bool(re.search(r"\\True\b", a))
            stmts.append((re.sub(r"\\True\b", "", a), ("Đ" if t else "S") if has_true else ""))
    if len(stmts) == 4:
        parts = []
        for st_, ans in stmts:
            parts.append("{" + self._inline(st_.strip(), _ICtx(in_cell=True)).strip() + "}{" + self._plain(ans) + "}")
        self._emit_block("\\choiceTF" + "".join(parts), ctx)
        self._stat("true_false_tables")
    else:
        for q, (st_, ans) in enumerate(stmts):
            lbl = "abcdefgh"[q] + ")"
            self._emit_para("\\textbf{" + lbl + "} " + self._inline(st_, _ICtx()) + (f" ({ans})" if ans else ""), {"left": ctx.left + 284}, ctx.child())
    return endk


def _h_immini(self, s, i, j, ctx, buf):
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    if o is not None:
        k = k2
    args, k3 = self._read_n_braced(s, k, 2)
    if len(args) < 2:
        return None
    self._flush(buf, ctx)
    left, right = args
    m = re.search(r"\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}", right, re.S)
    if not m or getattr(self, "_config_only", False):
        # no TikZ on the right → sequential layout
        self._blocks(left, ctx.child(pending=ctx.pending))
        self._blocks(right, ctx.child(align="c"))
        return k3
    saved = self.items
    self.items = []
    sub = ctx.child(pending=ctx.pending, in_immini=True)
    self._blocks(left, sub)
    left_items = self.items
    self.items = saved
    pieces: list[str] = []
    ok = True
    for kind, a, b in left_items:
        if kind == "para":
            if (b or {}).get("pb"):
                continue
            if a.strip():
                pieces.append(("P", a.strip()))
        else:
            raw = a.strip()
            if raw.startswith(("\\choice", "\\boncot", "\\haicot", "\\motcot", "\\[")):
                pieces.append(("B", raw))
            else:
                ok = False
    if not ok:
        self.items.extend(left_items)
        self._emit_block(m.group(0), ctx)
        return k3
    out = []
    for idx2, (t, raw) in enumerate(pieces):
        if t == "P":
            if out and out[-1][0] == "P":
                out[-1] = ("P", out[-1][1] + " \\\\ " + raw)
            else:
                out.append(("P", raw))
        else:
            out.append(("B", raw))
    left_tex = "\n".join(raw for _, raw in out)
    self._emit_block("\\immini{" + left_tex + "}{" + m.group(0) + "}", ctx)
    self._stat("immini_layouts")
    return k3


def _h_loigiai(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    if o is not None:
        k = k2
    a, k = read_arg(s, k)
    if a is None:
        return None
    self._flush(buf, ctx)
    if not self.show_solutions and name in ("loigiai", "giai", "huongdan"):
        return k
    lbl = {"loigiai": "Lời giải.", "giai": "Giải.", "huongdan": "Hướng dẫn.", "shortans": "Đáp án:",
           "dapso": "Đáp số:", "traloi": "Trả lời:", "dapan": "Đáp án:"}.get(name, "Lời giải.")
    sub = ctx.child()
    self._set_pending(sub, "\\textbf{" + lbl + "} ", {"sb": 4})
    self._blocks(a, sub)
    if sub.pending is not None and not sub.pending.used:
        self._emit_para("", None, sub)
    self._stat("solutions")
    return k


def _h_question(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    if name in self.expander.macros:
        return None
    self._flush(buf, ctx)
    title = QUESTION_LIKE.get(name, "Câu")
    if name == "question":
        o, k = read_opt(s, j, allow_newline=False)
        num = self.refstep("question", "question")
        self._set_pending(ctx, "\\textbf{" + num + ".} ", {"sb": 6})
        return _eat_space_after_word(s, k)
    num, k = self._question_number(s, j, "ex")
    self._set_pending(ctx, "\\textbf{" + f"{title} {num}." + "} ", {"sb": 6})
    self._stat("questions")
    return k


def _h_item_stray(self, s, i, j, ctx, buf):
    self._flush(buf, ctx)
    o, k = read_opt(s, j, allow_newline=False)
    lbl = self._inline(o, _ICtx()) if o is not None else "•"
    self._set_pending(ctx, lbl + MK_TAB, {"left": ctx.left + 567, "hang": 340})
    return _eat_space_after_word(s, k)


def _h_hrule(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    k = j
    if name == "rule":
        o, k = read_opt(s, k, allow_newline=False)
        a, k = read_arg(s, k)
        b, k = read_arg(s, k)
        w = dim_to_twips(a or "", self.textwidth) or 0
        if w < self.textwidth * 0.5 or buf_has_text_local(buf):
            buf.append(s[i:k])
            return k
    elif buf_has_text_local(buf):
        buf.append(MK_RULEFILL)
        return k
    self._flush(buf, ctx)
    if not getattr(self, "_config_only", False):
        self.items.append(("para", "", {"hr": 1, "sb": 2, "sa": 2}))
    return _eat_space_after_word(s, k)


def _h_appendix(self, s, i, j, ctx, buf):
    self.appendix = True
    self.counters["section"] = 0
    self.counters["chapter"] = 0
    return j


def _h_columns(self, s, i, j, ctx, buf):
    name = read_cmd(s, i)[0]
    self._flush(buf, ctx)
    if getattr(self, "_config_only", False):
        return j
    if name == "twocolumn":
        o, k = read_opt(s, j, allow_newline=False)
        self.items.append(("para", "", {"cols": 2}))
        return k
    if name == "onecolumn":
        self.items.append(("para", "", {"cols_end": 1}))
        return j
    self.items.append(("para", "", {"colbreak": 1}))
    return _eat_space_after_word(s, j)


def _h_tikz_inline(self, s, i, j, ctx, buf):
    if buf_has_text_local(buf):
        return None
    k = j
    o, k = read_opt(s, k, allow_newline=False)
    kk = skip_ws(s, k)
    if kk < len(s) and s[kk] == "{":
        body, k = read_group(s, kk)
    else:
        e = s.find(";", k)
        if e == -1:
            return None
        body = s[k : e + 1]
        k = e + 1
    self._flush(buf, ctx)
    self._emit_block("\\begin{tikzpicture}" + (f"[{o}]" if o else "") + "\n" + body + "\n\\end{tikzpicture}", ctx)
    return k


def _try_choice_like(self, name: str, s: str, i: int, j: int, ctx: _Ctx, buf: list[str]) -> int | None:
    """Unknown command + exactly 4 braced args with a choice-like name → multiple-choice table."""
    if not name.isalpha() or name in self.expander.macros or name in TEXT_SYMBOLS or name in INLINE_KNOWN:
        return None
    if name in self._BLOCK_HANDLERS or name in self._CONFIG_CMDS or name in MATH_SYMBOLS_ALL:
        return None
    if not CHOICE_LIKE_RE.search(name):
        return None
    k = j
    o, k2 = read_opt(s, k, allow_newline=False)
    if o is not None:
        k = k2
    args, k3 = self._read_n_braced(s, k, 5)
    if len(args) != 4:
        return None
    self._flush(buf, ctx)
    rendered, correct = self._choice_options_render(args)
    force = None
    if re.search(r"(hai|two)", name, re.I):
        force = 2
    elif re.search(r"(mot|one)", name, re.I):
        force = 1
    self.res.warnings.append(f"Lệnh lạ \\{name}{{..}}{{..}}{{..}}{{..}} được hiểu là 4 phương án trắc nghiệm.")
    self._emit_choice(rendered, correct, force, ctx)
    return k3


# =====================================================================================
# CONFIG COMMANDS (preamble or body)
# =====================================================================================
CONFIG_CMDS = {
    "title", "author", "date", "subtitle", "institute", "keywords", "newtheorem", "theoremstyle", "definecolor",
    "colorlet", "numberwithin", "counterwithin", "counterwithout", "setcounter", "addtocounter", "stepcounter",
    "refstepcounter", "newcounter", "setlength", "addtolength", "renewcommand", "pagestyle", "thispagestyle",
    "fancyhead", "fancyfoot", "lhead", "chead", "rhead", "lfoot", "cfoot", "rfoot", "fancyhf", "geometry",
    "usepackage", "documentclass", "hypersetup", "usetikzlibrary", "pgfplotsset", "tikzset", "setmainfont",
    "setsansfont", "setmonofont", "newfontfamily", "captionsetup", "setlist", "settasks", "sisetup", "lstset",
    "graphicspath", "allowdisplaybreaks", "newlength", "linespread", "onehalfspacing", "doublespacing",
    "singlespacing", "setstretch", "pagenumbering", "phantomsection", "bibliographystyle", "nocite",
    "DeclareGraphicsExtensions", "RequirePackage", "PassOptionsToPackage", "selectlanguage", "foreignlanguage@",
    "setmathfont", "tcbuselibrary", "newtcolorbox", "tcbset", "newmdenv", "mdfsetup", "pgfmathsetmacro",
    "theoremstyle", "newtheoremstyle", "swapnumbers", "renewenvironment@", "AtBeginDocument", "AtEndDocument",
    "headheight", "footskip", "setcounter@", "raggedbottom", "hyphenpenalty", "tolerance", "parindent",
    "parskip", "baselineskip", "newsavebox", "ExplSyntaxOn", "ExplSyntaxOff", "Opensolutionfile",
    "Closesolutionfile", "Setcommandkey", "setkeys", "Writesolution", "Addlinespace", "renewtheorem",
    "fancypagestyle", "addbibresource", "bibliography", "printbibliography", "includeonly", "listfiles",
    "SetWatermarkText", "SetWatermarkScale", "newgeometry", "restoregeometry", "titleformat", "titlespacing",
    "setsecnumdepth", "settocdepth", "tableofcontents@", "newlist", "setlistdepth", "pgfdeclarelayer",
    "pgfsetlayers", "definecolorset", "colorbox@", "DeclareUnicodeCharacter", "urlstyle", "hfuzz", "vfuzz",
    "emergencystretch", "widowpenalty", "clubpenalty", "interfootnotelinepenalty", "pretolerance",
    "makeindex", "makeglossaries", "loadglsentries", "newacronym", "newglossaryentry", "sloppy@",
    "arraycolsep", "tabcolsep", "doublerulesep", "arrayrulewidth", "extrarowheight", "tikzexternalize",
    "pgfplotsset@", "usepgfplotslibrary", "newcolumntype@", "dimexpr", "numexpr", "providecommand@",
}


def _config_cmd(self, name: str, s: str, j: int, ctx: _Ctx | None) -> int | None:
    n = len(s)

    def args(spec: str, start: int) -> tuple[list[str | None], int]:
        res: list[str | None] = []
        k = start
        for ch in spec:
            if ch == "s":
                st, k = read_star(s, k)
                res.append("*" if st else None)
            elif ch == "o":
                o, k = read_opt(s, k)
                res.append(o)
            else:
                a, k = read_arg(s, k)
                res.append(a)
        return res, k

    if name in ("title", "author", "date", "subtitle", "institute", "keywords"):
        (o, a), k = args("om", j)
        if a is None:
            return None
        if name == "title":
            self.title = a
        elif name == "author":
            self.author = a
        elif name == "date":
            self.date = a
        elif name == "subtitle" and self.title:
            self.title = self.title + " \\\\ " + a
        return k
    if name == "newtheorem":
        st, k = read_star(s, j)
        a, k = read_braced(s, k)
        shared, k = read_opt(s, k)
        title, k = read_braced(s, k)
        within, k = read_opt(s, k)
        if a is None or title is None:
            return None
        env = a.strip()
        counter = shared.strip() if shared else env
        if not shared:
            self.counters.setdefault(env, 0)
            if within:
                self.counter_parent[env] = within.strip()
        self.theorems[env] = {"title": self._plain(title), "counter": counter, "numbered": not st,
                              "style": self.theorem_style}
        return k
    if name == "theoremstyle":
        a, k = read_braced(s, j)
        self.theorem_style = (a or "plain").strip()
        return k
    if name == "newtheoremstyle":
        k = j
        for _ in range(9):
            a, k = read_arg(s, k)
        return k
    if name == "definecolor":
        (o, a, b, c), k = args("ommm", j)
        if a and b and c:
            hx = _hexmodel(b, c)
            if hx:
                self.colors[a.strip()] = hx
        return k
    if name == "colorlet":
        (o, a, b), k = args("omm", j)
        if a and b:
            hx = self.color_hex(b)
            if hx:
                self.colors[a.strip()] = hx
        return k
    if name in ("numberwithin", "counterwithin"):
        st, k = read_star(s, j)
        o, k = read_opt(s, k)
        a, k = read_braced(s, k)
        b, k = read_braced(s, k)
        if a and b:
            self.counter_parent[a.strip()] = b.strip()
        return k
    if name == "counterwithout":
        st, k = read_star(s, j)
        a, k = read_braced(s, k)
        b, k = read_braced(s, k)
        if a:
            self.counter_parent.pop(a.strip(), None)
        return k
    if name in ("setcounter", "addtocounter"):
        a, k = read_braced(s, j)
        b, k = read_braced(s, k)
        if a is None or b is None:
            return None
        cname = a.strip()
        m = re.search(r"-?\d+", self._plain(b) or "0")
        v = int(m.group(0)) if m else 0
        mv = re.search(r"\\value\{([^}]*)\}", b)
        if mv:
            v = self.counters.get(mv.group(1).strip(), 0) + (v if "+" in b else 0)
        if cname == "secnumdepth":
            self.secnumdepth = v
            return k
        if cname in ("tocdepth",):
            return k
        if name == "setcounter":
            self.counters[cname] = v
        else:
            self.counters[cname] = self.counters.get(cname, 0) + v
        return k
    if name in ("stepcounter", "refstepcounter"):
        a, k = read_braced(s, j)
        if a is None:
            return None
        if name == "stepcounter":
            self.step(a.strip())
        else:
            self.refstep(a.strip())
        return k
    if name == "newcounter":
        a, k = read_braced(s, j)
        p, k = read_opt(s, k)
        if a:
            self.counters.setdefault(a.strip(), 0)
            if p:
                self.counter_parent[a.strip()] = p.strip()
        return k
    if name in ("setlength", "addtolength"):
        a, k = read_arg(s, j)
        b, k = read_arg(s, k)
        return k
    if name == "renewcommand":
        st, k = read_star(s, j)
        kk = skip_ws(s, k)
        if kk < n and s[kk] == "{":
            nm, k = read_group(s, kk)
            nm = nm.strip().lstrip("\\")
        elif kk < n and s[kk] == "\\":
            nm, k = read_cmd(s, kk)
        else:
            return None
        o, k = read_opt(s, k)
        body, k = read_arg(s, k)
        if body is None:
            return None
        if nm.startswith("the") and len(nm) > 3:
            self.the_formats[nm[3:]] = body
        elif nm.startswith("labelenum"):
            depth = {"labelenumi": 1, "labelenumii": 2, "labelenumiii": 3, "labelenumiv": 4}.get(nm, 1)
            self.enum_labels[depth] = body
        elif nm.startswith("labelitem"):
            depth = {"labelitemi": 1, "labelitemii": 2, "labelitemiii": 3, "labelitemiv": 4}.get(nm, 1)
            self.item_labels[depth] = body
        elif nm.endswith("name"):
            self.names[nm] = self._plain(body)
        return k
    if name in ("pagestyle", "thispagestyle"):
        a, k = read_braced(s, j)
        st = (a or "plain").strip()
        if name == "pagestyle":
            self.page["pagestyle"] = st
        elif st == "empty" and not self.items:
            self.page["first_empty"] = True
        return k
    if name == "fancyhf":
        a, k = read_braced(s, j)
        self.page["header"] = {}
        self.page["footer"] = {}
        self.page["pagestyle"] = "fancy"
        return k
    if name in ("fancyhead", "fancyfoot", "lhead", "chead", "rhead", "lfoot", "cfoot", "rfoot"):
        o, k = read_opt(s, j)
        a, k = read_arg(s, k)
        if a is None:
            return None
        where = "header" if name in ("fancyhead", "lhead", "chead", "rhead") else "footer"
        if name in ("fancyhead", "fancyfoot"):
            spots = (o or "LCR").upper()
            pos = [p for p in "LCR" if p in spots] if any(p in spots for p in "LCR") else ["L", "C", "R"]
        else:
            pos = [name[0].upper()]
        txt = self._plain_with_fields(a)
        for p in pos:
            self.page[where][p] = txt
        self.page["pagestyle"] = "fancy"
        return k
    if name in ("geometry", "newgeometry"):
        a, k = read_braced(s, j)
        if a:
            self._geometry(a)
        return k
    if name in ("usepackage", "RequirePackage"):
        o, k = read_opt(s, j)
        a, k = read_braced(s, k)
        for p in (a or "").split(","):
            p = p.strip()
            self.packages.add(p)
            if p == "geometry" and o:
                self._geometry(o)
        return k
    if name == "documentclass":
        o, k = read_opt(s, j)
        a, k = read_braced(s, k)
        return k
    if name == "settasks":
        a, k = read_braced(s, j)
        if a:
            m = re.search(r"label\s*=\s*(\{[^}]*\}|[^,]+)", a)
            if m:
                self.tasks_label = m.group(1).strip().strip("{}")
            m = re.search(r"label-format\s*=\s*\{?\\bfseries", a)
            if m:
                self.tasks_bold = True
        return k
    if name == "setlist":
        o, k = read_opt(s, j)
        a, k = read_braced(s, k)
        if a and o:
            m = re.search(r"label\s*=\s*(\{.*\}|[^,]+)", a)
            if m:
                lbl = m.group(1).strip()
                if lbl.startswith("{") and lbl.endswith("}"):
                    lbl = lbl[1:-1]
                depth = int(re.search(r"(\d)", o).group(1)) if re.search(r"(\d)", o) else 1
                if "enumerate" in o:
                    self.enum_labels[depth] = lbl
                elif "itemize" in o:
                    self.item_labels[depth] = lbl
        return k
    # generic swallow: optional/star + up to N braced args depending on command
    nargs = {
        "hypersetup": 1, "usetikzlibrary": 1, "pgfplotsset": 1, "tikzset": 1, "setmainfont": 1, "setsansfont": 1,
        "setmonofont": 1, "newfontfamily": 2, "captionsetup": 1, "sisetup": 1, "lstset": 1, "graphicspath": 1,
        "newlength": 1, "linespread": 1, "setstretch": 1, "pagenumbering": 1, "bibliographystyle": 1, "nocite": 1,
        "DeclareGraphicsExtensions": 1, "PassOptionsToPackage": 2, "selectlanguage": 1, "setmathfont": 1,
        "tcbuselibrary": 1, "newtcolorbox": 2, "tcbset": 1, "newmdenv": 1, "mdfsetup": 1, "pgfmathsetmacro": 2,
        "AtBeginDocument": 1, "AtEndDocument": 1, "Opensolutionfile": 1, "Closesolutionfile": 1,
        "Setcommandkey": 2, "setkeys": 2, "Writesolution": 1, "renewtheorem": 2, "fancypagestyle": 2,
        "addbibresource": 1, "bibliography": 1, "includeonly": 1, "SetWatermarkText": 1, "SetWatermarkScale": 1,
        "titleformat": 5, "titlespacing": 4, "newlist": 3, "setlistdepth": 1, "pgfdeclarelayer": 1,
        "pgfsetlayers": 1, "DeclareUnicodeCharacter": 2, "urlstyle": 1, "newacronym": 3, "newglossaryentry": 2,
        "loadglsentries": 1, "tikzexternalize": 0, "usepgfplotslibrary": 1, "newsavebox": 1, "allowdisplaybreaks": 0,
        "onehalfspacing": 0, "doublespacing": 0, "singlespacing": 0, "phantomsection": 0, "raggedbottom": 0,
        "ExplSyntaxOn": 0, "ExplSyntaxOff": 0, "makeindex": 0, "makeglossaries": 0, "restoregeometry": 0,
        "swapnumbers": 0, "listfiles": 0, "printbibliography": 0,
    }.get(name, 0)
    k = j
    st, k = read_star(s, k)
    o, k2 = read_opt(s, k)
    if o is not None:
        k = k2
    if name == "newtcolorbox":
        # \newtcolorbox{name}[n][default]{options} → treat env as a box
        a, k = read_braced(s, k)
        o, k = read_opt(s, k)
        o2, k = read_opt(s, k)
        b, k = read_braced(s, k)
        if a:
            self.box_envs_user = getattr(self, "box_envs_user", set()) | {a.strip()}
        return k
    if name == "newmdenv":
        b, k = read_braced(s, k)
        if b:
            self.box_envs_user = getattr(self, "box_envs_user", set()) | {b.strip()}
        return k
    for _ in range(nargs):
        a, k = read_arg(s, k)
        o, k2 = read_opt(s, k, allow_newline=False)
        if o is not None:
            k = k2
    if name in ("bibliography", "addbibresource"):
        self.res.warnings.append("Tài liệu tham khảo từ file .bib không có khi dán mã — bỏ qua.")
    if name in ("hfuzz", "vfuzz", "parindent", "parskip", "baselineskip", "tabcolsep", "arraycolsep",
                "doublerulesep", "arrayrulewidth", "extrarowheight", "headheight", "footskip", "emergencystretch"):
        m = re.match(r"\s*=?\s*-?[\d.]+\s*(cm|mm|pt|em|ex|in)?(\s*plus\s*[\d.]+\s*\w+)?(\s*minus\s*[\d.]+\s*\w+)?", s[k:])
        if m:
            k += m.end()
    if name in ("hyphenpenalty", "tolerance", "widowpenalty", "clubpenalty", "interfootnotelinepenalty", "pretolerance"):
        m = re.match(r"\s*=?\s*-?\d+", s[k:])
        if m:
            k += m.end()
    return k


def _hexmodel(model: str, spec: str) -> str | None:
    from .latex_preprocessor import _model_to_hex
    return _model_to_hex(model.strip(), spec.strip())


def _plain_with_fields(self, s: str) -> str:
    t = s.replace("\\thepage", MK_PAGE)
    t = re.sub(r"\\pageref\*?\{\s*LastPage\s*\}", MK_NUMPAGES, t)
    out = self._inline(t, _ICtx(plain=True))
    out = "".join(LIT_REVERSE.get(ch, ch) for ch in out)
    out = re.sub("[\ue000-\ue008\ue00b-\ue00f]", "", out)
    return re.sub(r"[ \t]+", " ", out).strip()


# =====================================================================================
# ENVIRONMENTS
# =====================================================================================
def _strip_env_args(content: str, spec: str) -> tuple[list[str | None], str]:
    res: list[str | None] = []
    k = 0
    for ch in spec:
        if ch == "o":
            o, k = read_opt(content, k)
            res.append(o)
        elif ch == "m":
            a, k2 = read_braced(content, k)
            if a is not None:
                k = k2
            res.append(a)
        elif ch == "p":
            kk = skip_ws(content, k)
            if kk < len(content) and content[kk] == "(":
                a, k = read_paren(content, kk)
                res.append(a)
            else:
                res.append(None)
    return res, content[k:]


def _env(self, env: str, content: str, ctx: _Ctx, buf: list[str]) -> None:
    base = env.rstrip("*")
    if getattr(self, "_config_only", False) and env not in ("document",):
        return
    # ---------- math ----------
    if env in MATH_DISPLAY_ENVS:
        self._flush(buf, ctx)
        numbered = not env.endswith("*") and env not in ("displaymath", "empheq") and base not in ("dgroup",)
        body = content
        if base == "empheq":
            o, body = _strip_env_args(content, "om")[0], _strip_env_args(content, "om")[1]
            numbered = False
            inner = re.search(r"\\begin\{(\w+\*?)\}(.*)\\end\{\1\}", body, re.S)
            if inner:
                self._display_math(inner.group(2), ctx, numbered=not inner.group(1).endswith("*"), env=inner.group(1))
                return
        self._display_math(body, ctx, numbered=numbered, env=env)
        return
    if env == "math":
        buf.append("$" + content + "$")
        return
    # ---------- drawings ----------
    if env in ("tikzpicture", "circuitikz", "pgfpicture"):
        self._flush(buf, ctx)
        if env != "tikzpicture":
            self.res.warnings.append(f"Môi trường {env} được vẽ bằng bộ dịch TikZ (có thể thiếu chi tiết).")
        self._emit_block("\\begin{tikzpicture}" + content + "\\end{tikzpicture}", ctx)
        self._stat("tikz_figures")
        return
    if env == "axis":
        self._flush(buf, ctx)
        self._emit_block("\\begin{axis}" + content + "\\end{axis}", ctx)
        return
    if env in ("tikzcd", "pspicture", "xy", "forest", "picture", "asy", "asydef", "mplibcode", "chemfig@",
               "circuit", "dot2tex", "gnuplot", "tikzpicture@"):
        self._flush(buf, ctx)
        self.res.warnings.append(f"Hình '{env}' chưa dịch được sang hình Word — đã chèn khung giữ chỗ.")
        self._emit_para(self.span(PLACEHOLDER_IMG_SPAN, f"[Hình vẽ {env}]"),
                        {"align": "c", "box": {"style": "dashed", "color": "94A3B8"}}, ctx.child())
        return
    # ---------- tables ----------
    if env in TABLE_ENVS:
        self._flush(buf, ctx)
        self._table(env, content, ctx)
        return
    if env == "tabbing":
        self._flush(buf, ctx)
        self._tabbing(content, ctx)
        return
    # ---------- lists ----------
    if env in INLINE_LIST_ENVS:
        (o,), body = _strip_env_args(content, "o")
        buf.append(self._inline_list_text(body, env, o))
        return
    if env in LIST_ENVS:
        self._flush(buf, ctx)
        self._list(env, content, ctx)
        return
    if env in CHOICE_LIST_ENVS:
        self._flush(buf, ctx)
        self._choices_env(env, content, ctx)
        return
    if env in ("tasks", "tasks*"):
        self._flush(buf, ctx)
        self._tasks(content, ctx)
        return
    # ---------- alignment ----------
    if env in ("center", "centering", "flushleft", "flushright", "raggedleft", "raggedright", "flushleft*"):
        self._flush(buf, ctx)
        al = {"center": "c", "centering": "c", "flushleft": "l", "flushleft*": "l", "raggedright": "l",
              "flushright": "r", "raggedleft": "r"}[env]
        sub = ctx.child(align=al, pending=ctx.pending)
        self._blocks(content, sub)
        return
    if env in ("quote", "quotation", "verse", "adjustwidth", "addmargin", "changemargin", "indentation", "narrow"):
        self._flush(buf, ctx)
        body = content
        add = 567
        if env in ("adjustwidth", "changemargin", "indentation"):
            (a, b), body = _strip_env_args(content, "mm")
            add = dim_to_twips(a or "", self.textwidth) or 0
        elif env == "addmargin":
            (o, a), body = _strip_env_args(content, "om")
            add = dim_to_twips(a or "", self.textwidth) or 567
        sub = ctx.child(left=ctx.left + add, pending=ctx.pending)
        self._blocks(body, sub)
        return
    if env == "abstract" and self.doc_class.startswith("ams") and not getattr(self, "_title_done", False):
        self._flush(buf, ctx)
        self._deferred_abstract = content
        return
    if env == "abstract":
        self._flush(buf, ctx)
        title = self.names.get("abstractname") or ("Tóm tắt" if self.lang == "vi" else "Abstract")
        self._emit_para(self.span({"b": 1, "sz": 10.95}, title), {"align": "c", "sb": 8, "sa": 4}, ctx.child())
        sub = ctx.child(left=ctx.left + 567, spans=ctx.spans + (self.span_id({"sz": 10.95}),))
        self._blocks(content, sub)
        return
    if env == "proof":
        self._flush(buf, ctx)
        (o,), body = _strip_env_args(content, "o")
        head = self._inline(o, _ICtx()) if o else (self.names.get("proofname") or ("Chứng minh" if self.lang == "vi" else "Proof"))
        sub = ctx.child(pending=_Pending(text="\\textit{" + head + ".} ", used=False))
        start = len(self.items)
        self._blocks(body, sub)
        self._append_qed(start, ctx)
        return
    if env in ("multicols", "multicols*", "paracol"):
        self._flush(buf, ctx)
        (ncols, hdr), body = _strip_env_args(content, "mo")
        try:
            nc = max(1, min(4, int(self._plain(ncols or "2"))))
        except ValueError:
            nc = 2
        if hdr:
            self._blocks(hdr, ctx.child())
        self.items.append(("para", "", {"cols": nc}))
        self._blocks(body, ctx.child(pending=ctx.pending))
        self.items.append(("para", "", {"cols_end": 1}))
        self._stat("multicols")
        return
    if env in ("thebibliography",):
        self._flush(buf, ctx)
        _, body = _strip_env_args(content, "m")
        self._bibliography(body, ctx)
        return
    if base in ("frame",):
        self._flush(buf, ctx)
        k = 0
        o, k = read_opt(content, 0)
        if content[skip_ws(content, k):skip_ws(content, k) + 1] == "<":
            e = content.find(">", k)
            k = e + 1
        o2, k = read_opt(content, k)
        title, k2 = read_braced(content, k)
        body = content
        if title is not None:
            body = content[k2:]
            self._emit_para(self.span({"b": 1, "sz": 14.4, "c": "1F3A93"}, self._inline(title, _ICtx())), {"keep": 1, "sb": 8, "sa": 6}, ctx.child())
            sub_t, k3 = read_braced(body, 0)
            if sub_t is not None and not sub_t.strip().startswith("\\"):
                body = body[k3:]
                self._emit_para(self.span({"i": 1}, self._inline(sub_t, _ICtx())), {"keep": 1}, ctx.child())
        else:
            body = content[k:]
        self._blocks(body, ctx.child())
        self.items.append(("para", "", {"pb": 1}))
        return
    if base in ("block", "alertblock", "exampleblock"):
        self._flush(buf, ctx)
        title, body = _strip_env_args(content, "m")
        col = {"block": "1F3A93", "alertblock": "B91C1C", "exampleblock": "047857"}[base]
        sub = ctx.child(box={"style": "single", "color": col}, left=ctx.left)
        if title and title[0]:
            self._emit_para(self.span({"b": 1, "c": col}, self._inline(title[0], _ICtx())), {}, sub)
        self._blocks(body, sub)
        return
    # ---------- boxes ----------
    if base in BOX_ENVS or base in getattr(self, "box_envs_user", set()):
        self._flush(buf, ctx)
        (o,), body = _strip_env_args(content, "o")
        box: dict[str, Any] = {"style": "single", "color": "000000"}
        title = None
        shd = None
        if o:
            kv = dict((p.split("=", 1)[0].strip(), p.split("=", 1)[1].strip().strip("{}")) for p in self._split_kv(o) if "=" in p)
            if "colframe" in kv:
                box["color"] = self.color_hex(kv["colframe"]) or "000000"
            if "linecolor" in kv:
                box["color"] = self.color_hex(kv["linecolor"]) or "000000"
            if "colback" in kv:
                shd = self.color_hex(kv["colback"])
            if "backgroundcolor" in kv:
                shd = self.color_hex(kv["backgroundcolor"])
            title = kv.get("title") or kv.get("frametitle")
        if base in ("leftbar",):
            box = {"style": "single", "color": "808080", "sides": "l"}
        if base == "shaded" or base == "snugshade":
            box = None
            shd = shd or "E5E7EB"
        sub = ctx.child(box=box, shd=shd if shd and shd != "FFFFFF" else None, pending=ctx.pending)
        if title:
            self._emit_para(self.span({"b": 1}, self._inline(title, _ICtx())), {}, sub)
        self._blocks(body, sub)
        self._stat("boxes")
        return
    # ---------- theorem / exam envs ----------
    if env in self.theorems or base in self.theorems:
        self._flush(buf, ctx)
        self._theorem(env if env in self.theorems else base, content, ctx, starred=env.endswith("*"))
        return
    if base in ("dang", "dangtoan", "chude", "vande"):
        self._flush(buf, ctx)
        (t,), body = _strip_env_args(content, "m")
        num = self.refstep("dang@" + base, base)
        label = {"dang": "Dạng", "dangtoan": "Dạng", "chude": "Chủ đề", "vande": "Vấn đề"}[base]
        head = f"{label} {num}" + (". " + self._inline(t, _ICtx()) if t else ".")
        self._emit_para(self.span({"b": 1, "c": "1F3A93"}, head), {"keep": 1, "sb": 8, "sa": 4}, ctx.child())
        self._blocks(body, ctx.child())
        return
    if base in THEOREM_BUILTINS_VI or base in THEOREM_BUILTINS_EN:
        self._flush(buf, ctx)
        self._theorem(base, content, ctx, starred=env.endswith("*"), builtin=True)
        return
    # ---------- floats & transparent ----------
    if env in TRANSPARENT_ENVS:
        self._flush(buf, ctx)
        body = content
        sub = ctx.child(pending=ctx.pending)
        if base in ("figure", "table", "sidewaysfigure", "sidewaystable"):
            (o,), body = _strip_env_args(content, "o")
            sub.float_kind = "table" if "table" in base else "figure"
        elif base in ("wrapfigure", "wraptable"):
            (o1, a, o2, w), body = _strip_env_args(content, "omom")
            sub.float_kind = "table" if "table" in base else "figure"
        elif base in ("subfigure", "subtable"):
            (o, w), body = _strip_env_args(content, "om")
            sub.float_kind = "table" if "table" in base else "figure"
        elif base in ("minipage", "varwidth", "boxedminipage", "minipage*"):
            (o1, o2, o3, w), body = _strip_env_args(content, "ooom")
        elif base in ("adjustbox",):
            (a,), body = _strip_env_args(content, "m")
        elif base in ("spacing", "otherlanguage", "turn", "rotate", "column", "Form"):
            (a,), body = _strip_env_args(content, "m")
        elif base in ("columns",):
            (o,), body = _strip_env_args(content, "o")
        elif base in ("titlepage",):
            self._blocks(body, sub)
            self.items.append(("para", "", {"pb": 1}))
            return
        elif base in SIZE_SWITCH_PT:
            sub.spans = ctx.spans + (self.span_id({"sz": SIZE_SWITCH_PT[base]}),)
        elif base in ("bfseries", "bf"):
            sub.spans = ctx.spans + (self.span_id({"b": 1}),)
        elif base in ("itshape", "it", "em", "sl", "slshape"):
            sub.spans = ctx.spans + (self.span_id({"i": 1}),)
        elif base == "color":
            (a,), body = _strip_env_args(content, "m")
            hx = self.color_hex(a or "")
            if hx:
                sub.spans = ctx.spans + (self.span_id({"c": hx}),)
        elif base == "appendices":
            self.appendix = True
            self.counters["section"] = 0
        self._blocks(body, sub)
        return
    if base in ("loigiai", "giai", "huongdan", "solution", "dapan", "Solution"):
        self._flush(buf, ctx)
        if not self.show_solutions:
            return
        title = {"loigiai": "Lời giải.", "giai": "Giải.", "huongdan": "Hướng dẫn.", "solution": "Lời giải.",
                 "dapan": "Đáp án.", "Solution": "Solution."}[base]
        sub = ctx.child(pending=_Pending(text="\\textbf{" + title + "} ", used=False))
        self._blocks(content, sub)
        if sub.pending is not None and not sub.pending.used:
            self._emit_para("", None, sub)
        return
    if env == "document":
        self._blocks(content, ctx)
        return
    # ---------- unknown environment → transparent ----------
    self.res.unknown_environments[env] = self.res.unknown_environments.get(env, 0) + 1
    self._flush(buf, ctx)
    body = content
    k = 0
    o, k = read_opt(body, 0)
    for _ in range(3):
        kk = skip_ws(body, k)
        if kk < len(body) and body[kk] == "{":
            a, k2 = read_group(body, kk)
            if re.fullmatch(r"\s*-?[\d.]*\s*(cm|mm|pt|em|ex|in)?\s*(\\(textwidth|linewidth|columnwidth|hsize))?\s*", a) or re.fullmatch(r"\s*[a-zA-Z]{1,2}\s*", a):
                k = k2
                continue
        break
    body = body[k:]
    kk = skip_ws(body, 0)
    if kk < len(body) and body[kk] == "{":
        a, k2 = read_group(body, kk)
        if a.strip() and not a.strip().startswith("\\") and len(a) < 200:
            self._emit_para(self.span({"b": 1}, self._inline(a, _ICtx())), {"keep": 1}, ctx.child())
            body = body[k2:]
    self._blocks(body, ctx.child(pending=ctx.pending))


def _append_qed(self, start: int, ctx: _Ctx) -> None:
    for idx in range(len(self.items) - 1, start - 1, -1):
        kind, a, b = self.items[idx]
        if kind == "para":
            self.items[idx] = ("para", a + MK_HFILL + "□", b)
            return
        break
    self._emit_para("□", {"align": "r"}, ctx.child())


def _theorem(self, env: str, content: str, ctx: _Ctx, *, starred: bool = False, builtin: bool = False) -> None:
    if env in self.theorems:
        th = self.theorems[env]
        title, numbered, style, counter = th["title"], th["numbered"] and not starred, th["style"], th["counter"]
    elif env in THEOREM_BUILTINS_VI:
        title, numbered, style = THEOREM_BUILTINS_VI[env]
        numbered = numbered and not starred
        counter = "ex" if env in ("ex", "cau", "question", "cauhoi") else env
    else:
        en, vi, numbered, style = THEOREM_BUILTINS_EN[env]
        title = vi if self.lang == "vi" else en
        numbered = numbered and not starred
        counter = env
    (o,), body = _strip_env_args(content, "o")
    if style == "solution" and not self.show_solutions:
        return
    if style == "exam":
        o = None  # nguồn/ID của câu hỏi (ex_test) — không in
    num = self.refstep(counter, env) if numbered else ""
    head = title + (f" {num}" if num else "")
    note = f" ({self._inline(o, _ICtx())})" if o else ""
    if style == "remark":
        head_txt = "\\textit{" + head + note + ".} "
    elif style == "plain" and o:
        head_txt = "\\textbf{" + head + "} (" + self._inline(o, _ICtx()) + ")\\textbf{.} "
    else:
        head_txt = "\\textbf{" + head + note + ".} "
    spans = ctx.spans
    if style == "plain" and not builtin:
        spans = spans + (self.span_id({"i": 1}),)
    sub = ctx.child(spans=spans, pending=_Pending(text=head_txt, spec={"sb": 6}, used=False))
    self._blocks(body, sub)
    if sub.pending is not None and not sub.pending.used:
        self._emit_para("", None, sub)
    self._stat("theorem_like" if style != "exam" else "exam_questions")


# ---------------- lists ----------------
def _enum_format(self, template: str, n: int) -> str:
    t = template
    for star, fn in (("\\arabic*", str), ("\\alph*", _alph), ("\\Alph*", lambda x: _alph(x).upper()),
                     ("\\roman*", _roman), ("\\Roman*", lambda x: _roman(x).upper())):
        if star in t:
            t = t.replace(star, fn(n))
            return self._inline(t, _ICtx())
    # enumerate-package short syntax: first of 1 a A i I outside braces is the counter
    out = []
    depth = 0
    done = False
    i = 0
    while i < len(t):
        ch = t[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
        elif depth == 0 and not done and ch in "1aAiI" and (i == 0 or not t[i - 1].isalpha()) and (i + 1 >= len(t) or not t[i + 1].isalpha()):
            out.append({"1": str(n), "a": _alph(n), "A": _alph(n).upper(), "i": _roman(n), "I": _roman(n).upper()}[ch])
            done = True
            i += 1
            continue
        if ch not in "{}":
            out.append(ch)
        i += 1
    if not done:
        return self._inline(t, _ICtx())
    return self._inline("".join(out), _ICtx())


ENUM_DEFAULTS = {1: "\\arabic*.", 2: "(\\alph*)", 3: "\\roman*.", 4: "\\Alph*."}
ITEM_DEFAULTS = {1: "•", 2: "–", 3: "∗", 4: "·"}


def _list(self, env: str, content: str, ctx: _Ctx) -> None:
    base = env.rstrip("*")
    is_enum = base in ("enumerate", "compactenum", "asparaenum", "questions", "parts", "subparts", "subsubparts", "steps")
    is_desc = base in ("description", "compactdesc")
    opt, body = _strip_env_args(content, "o")
    o = opt[0]
    depth_e = ctx.enum_depth + (1 if is_enum else 0)
    depth_i = ctx.item_depth + (1 if not is_enum and not is_desc else 0)
    template = None
    start = 1
    resume = False
    indent_step = 567 if (ctx.list_depth == 0) else 454
    if o:
        if "=" in o or o.strip() in ("resume", "noitemsep", "nosep", "nolistsep"):
            kv = {}
            for p in self._split_kv(o):
                if "=" in p:
                    a, b = p.split("=", 1)
                    kv[a.strip()] = b.strip()
                else:
                    kv[p.strip()] = True
            if "label" in kv and isinstance(kv["label"], str):
                lb = kv["label"]
                if lb.startswith("{") and lb.endswith("}"):
                    lb = lb[1:-1]
                template = lb
            if "start" in kv:
                try:
                    start = int(str(kv["start"]).strip("{} "))
                except ValueError:
                    pass
            if "resume" in kv or "resume*" in kv:
                resume = True
            if "leftmargin" in kv and isinstance(kv["leftmargin"], str):
                tw = dim_to_twips(kv["leftmargin"].strip("{}"), self.textwidth)
                if tw is not None and tw > 0:
                    indent_step = tw
                elif kv["leftmargin"] in ("*", "0pt", "0"):
                    indent_step = 0 if kv["leftmargin"] != "*" else indent_step
        else:
            template = o
    if base == "questions":
        template = template or "\\arabic*."
    elif base == "parts":
        template = template or "(\\alph*)"
    elif base == "subparts":
        template = template or "\\roman*."
    if is_enum and template is None:
        template = self.enum_labels.get(depth_e) or ENUM_DEFAULTS.get(depth_e, "\\arabic*.")
    sep = {"questions": "question", "parts": "part", "subparts": "subpart", "subsubparts": "subsubpart"}.get(base, "item")
    chunks = split_top_level(body, (sep, "item"))
    counter = f"enum{'i' * depth_e if depth_e <= 3 else 'iv'}" if depth_e <= 4 else "enumiv"
    if base == "questions":
        counter = "question"
    n = (self.list_resume.get(env, 0) if resume else start - 1)
    left = ctx.left + indent_step
    hang = 397 if not is_enum else 425
    if is_desc:
        hang = 0
    pre = chunks[0][1] if chunks and chunks[0][0] == "" else ""
    if pre.strip():
        self._blocks(pre, ctx.child(left=left))
    for cmd, chunk in chunks:
        if cmd == "":
            continue
        mo = re.match(r"\s*<[^>]*>", chunk)
        if mo:
            chunk = chunk[mo.end():]
        o2, k = read_opt(chunk, 0)
        item_body = chunk[k:] if o2 is not None else chunk
        if cmd in ("question", "part") and o2 is not None and re.fullmatch(r"\s*\d+\s*", o2):
            o2 = None  # điểm số của câu hỏi (exam class)
        if o2 is not None:
            label = self._inline(o2, _ICtx())
            if is_desc:
                label = "\\textbf{" + label + "}"
        elif is_enum:
            n += 1
            self.counters[counter] = n
            label = self._enum_format(template or "\\arabic*.", n)
            fmt = {1: str, 2: _alph, 3: _roman, 4: lambda x: _alph(x).upper()}.get(depth_e, str)
            self.current_label = (fmt(n), "item")
            if base == "questions":
                label = "\\textbf{" + label + "}"
        elif is_desc:
            label = ""
        else:
            lbl_t = self.item_labels.get(depth_i)
            if template:
                label = self._inline(template, _ICtx())
            elif lbl_t:
                label = self._inline(lbl_t, _ICtx())
            else:
                label = ITEM_DEFAULTS.get(min(depth_i, 4), "•")
        sub = ctx.child(left=left, list_depth=ctx.list_depth + 1, enum_depth=depth_e, item_depth=depth_i)
        if is_desc:
            sub.pending = _Pending(text=(label + " ") if label else "", spec={"left": left, "hang": 567}, used=False)
            sub.first_para_spec = None
        else:
            sub.pending = _Pending(text=label + MK_TAB, spec={"left": left, "hang": hang}, used=False)
        self._blocks(item_body, sub)
        if sub.pending is not None and not sub.pending.used:
            self._emit_para("", None, sub)
    if is_enum:
        self.list_resume[env] = n
    self._stat("list_items", sum(1 for c, _ in chunks if c))


def _inline_list_text(self, body: str, env: str, template: str | None = None) -> str:
    chunks = split_top_level(body, ("item",))
    parts = []
    cnt = 0
    if template and "label=" in template:
        m = re.search(r"label\s*=\s*(\{[^}]*\}|[^,]+)", template)
        template = m.group(1).strip("{} ") if m else None
    for cmd, chunk in chunks:
        if cmd != "item":
            continue
        cnt += 1
        o, k = read_opt(chunk, 0)
        if o is not None:
            lbl = o
            chunk = chunk[k:]
        elif "enum" in env:
            lbl = self._enum_format(template, cnt) if template else f"({cnt})"
        else:
            lbl = "•"
        parts.append(lbl + " " + chunk.strip())
    return " " + " ".join(parts) + " "


def _choices_env(self, env: str, content: str, ctx: _Ctx) -> None:
    chunks = split_top_level(content, ("choice", "CorrectChoice", "correctchoice", "item"))
    opts = []
    correct = None
    for cmd, chunk in chunks:
        if cmd == "":
            continue
        if cmd in ("CorrectChoice", "correctchoice"):
            correct = len(opts)
        opts.append(self._inline(chunk.strip(), _ICtx(in_cell=True)).strip())
    if not opts:
        return
    if "checkbox" in env:
        labels = ["☐"] * len(opts)
        self._options_rows(labels, opts, len(opts) if "onepar" in env else 1, ctx)
        return
    force = None if "onepar" not in env else 4
    self._emit_choice(opts, correct, force, ctx)


def _tasks(self, content: str, ctx: _Ctx) -> None:
    (o, cols), body = _strip_env_args(content, "op")
    try:
        ncols = int(cols) if cols else 1
    except ValueError:
        ncols = 1
    chunks = split_top_level(body, ("task",))
    tasks = []
    for cmd, chunk in chunks:
        if cmd != "task":
            continue
        k = 0
        if chunk.startswith("*"):
            k = 1
            if k < len(chunk) and chunk[k] == "(":
                _, k = read_paren(chunk, k)
        o2, k2 = read_opt(chunk, k)
        tasks.append(chunk[k2:] if o2 is not None else chunk[k:])
    template = None
    if o:
        m = re.search(r"label\s*=\s*(\{[^}]*\}|[^,\]]+)", o)
        if m:
            template = m.group(1).strip().strip("{}")
    template = template or self.tasks_label or "\\alph*)"
    rendered = [self._inline(t.strip(), _ICtx(in_cell=True)).strip() for t in tasks]
    labels = [self._enum_format(template, idx + 1) for idx in range(len(rendered))]
    if len(rendered) == 4 and template.replace(" ", "") in ("\\Alph*.", "A.", "\\textbf{\\Alph*.}", "\\bfseries\\Alph*."):
        self._emit_choice(rendered, None, ncols if ncols in (1, 2, 4) else None, ctx)
        return
    width = self.textwidth - ctx.left
    stops = [ctx.left + int(width * c / ncols) for c in range(1, ncols)]
    for r in range(0, len(rendered), ncols):
        cells = []
        for c in range(ncols):
            idx = r + c
            if idx >= len(rendered):
                break
            cells.append("\\textbf{" + labels[idx] + "} " + rendered[idx] if getattr(self, "tasks_bold", False) else labels[idx] + " " + rendered[idx])
        self._emit_para(MK_TAB.join(cells), {"tabs": stops, "left": ctx.left, "sb": 1, "sa": 1}, ctx.child())
    self._stat("tasks_rows")


def _tabbing(self, content: str, ctx: _Ctx) -> None:
    rows = self._split_rows(content)
    stops: list[int] = []
    for r in rows:
        if "\\=" in r:
            pieces = r.replace("\\kill", "").split("\\=")
            acc = ctx.left
            stops = []
            for p in pieces[:-1]:
                acc += max(400, len(self._plain(p)) * 125 + 200)
                stops.append(acc)
            if "\\kill" in r:
                continue
        if "\\kill" in r:
            continue
        t = r.replace("\\=", "\\>").replace("\\kill", "")
        cells = [self._inline(c, _ICtx()) for c in t.split("\\>")]
        txt = MK_TAB.join(cells)
        if txt.strip():
            self._emit_para(txt, {"tabs": stops or [ctx.left + 1134 * k2 for k2 in range(1, 8)], "sb": 0, "sa": 0}, ctx.child())


def _bibliography(self, body: str, ctx: _Ctx) -> None:
    title = self.names.get("refname") or self.names.get("bibname") or ("Tài liệu tham khảo" if self.lang == "vi" else "References")
    if self.doc_class.startswith("ams"):
        self._emit_para(self.span({"sc": 1}, title), {"heading": 1, "keep": 1, "sb": 12, "sa": 6, "align": "c"}, ctx.child())
    else:
        self._emit_para(self.span({"b": 1, "sz": 17.28}, title), {"heading": 1, "keep": 1, "sb": 14, "sa": 8}, ctx.child())
    chunks = split_top_level(body, ("bibitem",))
    n = 0
    for cmd, chunk in chunks:
        if cmd != "bibitem":
            continue
        o, k = read_opt(chunk, 0)
        key, k2 = read_braced(chunk, k)
        n += 1
        lbl = self._inline(o, _ICtx()) if o else str(n)
        if key:
            self.bib_keys[key.strip()] = lbl if o else n
        text = self._inline(chunk[k2:].strip(), _ICtx())
        self._emit_para("[" + lbl + "]" + MK_TAB + text, {"left": ctx.left + 567, "hang": 567}, ctx.child())


# =====================================================================================
# TABLES
# =====================================================================================
def _parse_colspec(self, spec: str) -> tuple[list[dict[str, Any]], list[int]]:
    cols: list[dict[str, Any]] = []
    vlines = [0]
    i = 0
    n = len(spec)
    pending_align: str | None = None
    guard = 0
    while i < n and guard < 2000:
        guard += 1
        c = spec[i]
        if c in " \t\n":
            i += 1
            continue
        if c == "|":
            vlines[-1] += 1
            i += 1
            continue
        if c == "!" or c == "@":
            a, j = read_arg(spec, i + 1)
            if a and ("\\vline" in a or "|" in a):
                vlines[-1] += 1
            i = j
            continue
        if c == ">":
            a, j = read_arg(spec, i + 1)
            if a:
                if "\\centering" in a:
                    pending_align = "c"
                elif "\\raggedleft" in a:
                    pending_align = "r"
                elif "\\raggedright" in a:
                    pending_align = "l"
            i = j
            continue
        if c == "<":
            _, j = read_arg(spec, i + 1)
            i = j
            continue
        if c == "*":
            a, j = read_arg(spec, i + 1)
            b, j2 = read_arg(spec, j)
            try:
                rep = int((a or "1").strip())
            except ValueError:
                rep = 1
            spec = spec[:i] + (b or "") * rep + spec[j2:]
            n = len(spec)
            continue
        if c in "lcrLCRXJSZYD" or c in "pmbw" or c == "Q" or c in self.expander.coltypes:
            if c in self.expander.coltypes and c not in "lcrpmbX":
                nargs, body = self.expander.coltypes[c]
                args = []
                j = i + 1
                for _ in range(nargs):
                    a, j = read_arg(spec, j)
                    args.append(a or "")
                rep = re.sub(r"#([1-9])", lambda m: args[int(m.group(1)) - 1] if int(m.group(1)) <= len(args) else "", body)
                spec = spec[:i] + rep + spec[j:]
                n = len(spec)
                continue
            col: dict[str, Any] = {"align": "l", "width": None, "x": False}
            j = i + 1
            if c in "lL":
                col["align"] = "l"
            elif c in "cCS":
                col["align"] = "c"
            elif c in "rR":
                col["align"] = "r"
            elif c in "XYZJ":
                col["x"] = True
                col["align"] = "c" if c in "YZ" else "l"
            elif c in "pmb":
                a, j = read_arg(spec, j)
                col["width"] = dim_to_twips(a or "", self.textwidth)
            elif c == "w":
                a, j = read_arg(spec, j)
                b, j = read_arg(spec, j)
                col["align"] = (a or "l").strip()[:1] or "l"
                col["width"] = dim_to_twips(b or "", self.textwidth)
            elif c == "D":
                for _ in range(3):
                    _, j = read_arg(spec, j)
                col["align"] = "c"
            elif c == "Q":
                o, j = read_opt(spec, j) if spec[j : j + 1] == "[" else (None, j)
                if o and re.search(r"\bc\b", o):
                    col["align"] = "c"
                if o and re.search(r"\br\b", o):
                    col["align"] = "r"
            if pending_align:
                col["align"] = pending_align
                pending_align = None
            cols.append(col)
            vlines.append(0)
            i = j
            continue
        i += 1
    if not cols:
        cols = [{"align": "c", "width": None, "x": False}]
        vlines = [0, 0]
    return cols, vlines


def _split_cells(self, row: str) -> list[str]:
    cells: list[str] = []
    cur: list[str] = []
    depth = 0
    env_depth = 0
    i = 0
    n = len(row)
    while i < n:
        c = row[i]
        if c == "\\":
            if row.startswith("\\begin{", i):
                env_depth += 1
            elif row.startswith("\\end{", i):
                env_depth -= 1
            cur.append(row[i : i + 2])
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth = max(0, depth - 1)
        if c == "&" and depth == 0 and env_depth == 0:
            cells.append("".join(cur))
            cur = []
        else:
            cur.append(c)
        i += 1
    cells.append("".join(cur))
    return cells


RULE_RE = re.compile(
    r"^\s*(\\(?:hline|toprule|midrule|bottomrule|hdashline|Hline|firsthline|lasthline)\b(?:\[[^\]]*\])?|"
    r"\\(?:cline|cmidrule|cdashline)\s*(?:\([^)]*\))?\s*(?:\[[^\]]*\])?\s*\{[^}]*\}|"
    r"\\(?:Xhline|specialrule)\s*\{[^}]*\}(?:\s*\{[^}]*\}\s*\{[^}]*\})?|\\hhline\s*\{[^}]*\}|"
    r"\\(?:rowcolor|rowcolors)\s*(?:\[[^\]]*\])?\s*\{[^}]*\}(?:\s*\{[^}]*\})?|\\arrayrulecolor\s*(?:\[[^\]]*\])?\s*\{[^}]*\}|"
    r"\\addlinespace(?:\[[^\]]*\])?|\\morecmidrules|\\noalign\s*\{[^}]*\}|\\(?:endhead|endfirsthead|endfoot|endlastfoot)\b|"
    r"\\(?:centering|small|footnotesize|scriptsize|normalsize|large|tiny)\b|\\label\{[^}]*\}|\\caption\*?(?:\[[^\]]*\])?\{(?:[^{}]|\{[^{}]*\})*\})"
)


def _parse_rule(self, tok: str, ncols: int) -> tuple[str, dict[int, tuple[str, int]] | None, str | None]:
    """Return (kind, coverage, color) for one rule token."""
    t = tok.strip()
    full = {c: ("single", 4) for c in range(ncols)}
    if t.startswith("\\hline") or t.startswith("\\firsthline") or t.startswith("\\lasthline") or t.startswith("\\Hline"):
        return "line", full, None
    if t.startswith("\\hdashline"):
        return "line", {c: ("dashed", 4) for c in range(ncols)}, None
    if t.startswith("\\toprule") or t.startswith("\\bottomrule"):
        return "line", {c: ("single", 8) for c in range(ncols)}, None
    if t.startswith("\\midrule"):
        return "line", {c: ("single", 6) for c in range(ncols)}, None
    if t.startswith("\\Xhline") or t.startswith("\\specialrule"):
        m = re.search(r"\{([^}]*)\}", t)
        w = (dim_to_twips(m.group(1)) or 20) if m else 20
        sz = max(4, min(24, int(w / 20 * 8)))
        return "line", {c: ("single", sz) for c in range(ncols)}, None
    if t.startswith("\\hhline"):
        m = re.search(r"\{([^}]*)\}", t)
        style = "double" if m and "=" in m.group(1) else "single"
        return "line", {c: (style, 4) for c in range(ncols)}, None
    m = re.search(r"\{\s*(\d+)\s*-\s*(\d+)\s*\}", t)
    if m and (t.startswith("\\cline") or t.startswith("\\cmidrule") or t.startswith("\\cdashline")):
        a, b = int(m.group(1)) - 1, int(m.group(2)) - 1
        style = "dashed" if t.startswith("\\cdashline") else "single"
        return "line", {c: (style, 4) for c in range(max(0, a), min(ncols, b + 1))}, None
    if t.startswith("\\rowcolor"):
        mm = re.match(r"\\rowcolor\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}", t)
        if mm:
            return "rowcolor", None, self.color_hex(mm.group(2), mm.group(1))
    if t.startswith("\\caption"):
        return "caption", None, t
    if t.startswith("\\label"):
        return "label", None, t
    return "other", None, None


def _table(self, env: str, content: str, ctx: _Ctx) -> None:
    base = env.rstrip("*")
    k = 0
    target_width = None
    spec = ""
    body = content
    if base in ("tabular", "longtable", "array", "supertabular", "xtabular", "NiceTabular", "TAB"):
        if env in ("tabular*",):
            w, k = read_braced(content, 0)
            target_width = dim_to_twips(w or "", self.textwidth)
        o, k = read_opt(content, k)
        sp, k = read_braced(content, k)
        spec = sp or "c"
        if base == "NiceTabular":
            o2, k2 = read_opt(content, k)
            if o2 is not None:
                k = k2
        body = content[k:]
    elif base in ("tabularx", "tabulary", "xltabular"):
        w, k = read_braced(content, 0)
        target_width = dim_to_twips(w or "", self.textwidth)
        o, k = read_opt(content, k)
        sp, k = read_braced(content, k)
        spec = sp or "X"
        body = content[k:]
    elif base in ("tabu", "longtabu"):
        m = re.match(r"\s*(?:to|spread)\s*([^{]*)", content)
        if m:
            target_width = dim_to_twips(m.group(1).strip(), self.textwidth)
            k = m.end()
        sp, k = read_braced(content, k)
        spec = sp or "c"
        body = content[k:]
    elif base in ("tblr", "longtblr", "talltblr"):
        o, k = read_opt(content, 0)
        sp, k = read_braced(content, k)
        sp = sp or ""
        m = re.search(r"colspec\s*=\s*\{", sp)
        if m:
            spec, _ = read_group(sp, m.end() - 1)
        elif "=" not in sp:
            spec = sp
        else:
            spec = ""
        body = content[k:]
        if not spec:
            first = self._split_rows(body)[0] if body.strip() else ""
            spec = "c" * max(1, len(self._split_cells(first)))
    cols, vlines0 = self._parse_colspec(spec)
    ncols = len(cols)
    # longtable heads/feet
    if base in ("longtable", "xltabular", "longtabu", "supertabular", "xtabular"):
        if "\\endfirsthead" in body:
            first, rest = body.split("\\endfirsthead", 1)
            if "\\endhead" in rest:
                rest = rest.split("\\endhead", 1)[1]
            body = first + "\\\\" + rest if not first.rstrip().endswith("\\\\") else first + rest
        elif "\\endhead" in body:
            body = body.replace("\\endhead", "")
        body = body.replace("\\endfoot", "").replace("\\endlastfoot", "")
    raw_rows = self._split_rows(body)
    # collect rules/captions per row
    rows: list[dict[str, Any]] = []
    pending_rules: list[str] = []
    captions: list[str] = []
    labels: list[str] = []
    for rr in raw_rows:
        rest = rr
        lead: list[str] = []
        while True:
            m = RULE_RE.match(rest)
            if not m:
                break
            lead.append(m.group(1))
            rest = rest[m.end():]
        rest_stripped = rest.strip()
        pending_rules.extend(lead)
        if not rest_stripped:
            continue
        rows.append({"rules": pending_rules, "src": rest})
        pending_rules = []
    bottom_rules = pending_rules
    for row in rows:
        new_rules = []
        for t in row["rules"]:
            kind, cov, extra = self._parse_rule(t, ncols)
            if kind == "caption":
                captions.append(extra)
            elif kind == "label":
                labels.append(extra)
            else:
                new_rules.append(t)
        row["rules"] = new_rules
    for cap in captions:
        self._blocks(cap, ctx.child(float_kind="table"))
    for lb in labels:
        self._blocks(lb, ctx)
    # parse cells
    grid_rows: list[list[dict[str, Any]]] = []
    for r_i, row in enumerate(rows):
        cells_src = self._split_cells(row["src"])
        cells = []
        for c_src in cells_src:
            cells.append(self._parse_table_cell(c_src))
        grid_rows.append(cells)
    real_cols = max([sum(c["colspan"] for c in r) for r in grid_rows] + [ncols])
    if real_cols > ncols:
        for _ in range(real_cols - ncols):
            cols.append({"align": "c", "width": None, "x": False})
            vlines0.append(0)
        ncols = real_cols
    # choice-table detection (A. B. C. D. laid out in a tabular)
    flat = [c for r in grid_rows for c in r if c["content"].strip()]
    if len(flat) == 4:
        labels_ok = []
        correct = None
        opts = []
        for idx, c in enumerate(flat):
            m = re.match(r"^\s*(?:\\(?:textbf|underline|textcolor\{[^}]*\}|bf)\s*)?\{?\s*(?:\\(?:textbf|underline)\s*\{)?\s*([A-D])\.\s*\}*\s*", c["content"])
            if m and m.group(1) == "ABCD"[idx]:
                if "\\underline" in c["content"][: m.end()] or "\\True" in c["content"]:
                    correct = idx
                labels_ok.append(True)
                opts.append(c["content"][m.end():].replace("\\True", ""))
            else:
                labels_ok.append(False)
        if all(labels_ok):
            rendered = [self._inline(o2.strip(), _ICtx(in_cell=True)).strip() for o2 in opts]
            ncol_eff = max(len(r) for r in grid_rows)
            self._emit_choice(rendered, correct, ncol_eff if ncol_eff in (1, 2, 4) else None, ctx)
            return
    # build output rows + coverage
    vmerge = [0] * ncols
    out_rows: list[str] = []
    marker_placed = False
    row_rules: list[dict[int, tuple[str, int]]] = []
    row_shd: dict[int, str] = {}
    cell_shd: dict[str, str] = {}
    mc_bars: dict[str, tuple[int, int]] = {}
    tid = len(self.res.table_specs)
    for r_i, (row, cells) in enumerate(zip(rows, grid_rows)):
        cov: dict[int, tuple[str, int]] = {}
        for t in row["rules"]:
            kind, cv, extra = self._parse_rule(t, ncols)
            if kind == "line" and cv:
                for c, st in cv.items():
                    if c in cov and cov[c][0] == st[0] == "single" and t.startswith("\\hline"):
                        cov[c] = ("double", 4)
                    else:
                        cov[c] = st
            elif kind == "rowcolor" and extra:
                row_shd[r_i] = extra
        row_rules.append(cov)
        emitted: list[str] = []
        c_pos = 0
        for cell in cells:
            if c_pos < ncols and vmerge[c_pos] > 0:
                # placeholder under a \multirow → engine covers it automatically
                for cc in range(c_pos, min(ncols, c_pos + cell["colspan"])):
                    vmerge[cc] = max(0, vmerge[cc] - 1)
                c_pos += cell["colspan"]
                continue
            txt = self._inline(cell["content"], _ICtx(in_cell=True, bold_via_span=(r_i == 0)))
            txt = txt.strip()
            if txt.startswith("\\\\"):
                txt = txt[2:].strip()
            if " \\\\ " in txt or txt.endswith("\\\\"):
                txt = "{" + txt.rstrip("\\ ") + "}"
            if not marker_placed and cell["diag"] is None:
                txt = MK_TABLE + idx_char(tid) + txt
                marker_placed = True
            if cell["shd"]:
                cell_shd[f"{r_i},{c_pos}"] = cell["shd"]
            if cell["mc_spec"] is not None:
                _, lb = self._parse_colspec(cell["mc_spec"])
                mc_bars[f"{r_i},{c_pos}"] = (lb[0], lb[-1])
            piece = txt
            if cell["diag"] is not None:
                a, b = cell["diag"]
                piece = "\\diagbox{" + self._plain(a) + "}{" + self._plain(b) + "}"
            if cell["rowspan"] > 1:
                piece = "\\multirow{" + str(cell["rowspan"]) + "}{*}{" + piece + "}"
                for cc in range(c_pos, min(ncols, c_pos + cell["colspan"])):
                    vmerge[cc] = cell["rowspan"] - 1
            if cell["colspan"] > 1 or cell["mc_spec"] is not None:
                al = cell.get("mc_align") or "c"
                piece = "\\multicolumn{" + str(cell["colspan"]) + "}{" + al + "}{" + piece + "}"
            emitted.append(piece)
            c_pos += cell["colspan"]
        # trailing covered columns
        for cc in range(c_pos, ncols):
            if vmerge[cc] > 0:
                vmerge[cc] -= 1
        line = " & ".join(emitted)
        if not line.strip():
            line = " \u200b "
        out_rows.append(line)
    if not out_rows:
        return
    bottom: dict[int, tuple[str, int]] = {}
    for t in bottom_rules:
        kind, cv, extra = self._parse_rule(t, ncols)
        if kind == "line" and cv:
            for c, st in cv.items():
                if c in bottom and t.startswith("\\hline"):
                    bottom[c] = ("double", 4)
                else:
                    bottom[c] = st
    if not marker_placed:
        out_rows[0] = MK_TABLE + idx_char(tid) + out_rows[0]
    spec_simple = "".join(c["align"] if c["align"] in "lcr" else "l" for c in cols)
    tex = "\\begin{tabular}{" + spec_simple + "}\n" + " \\\\\n".join(out_rows) + "\n\\end{tabular}"
    self.res.table_specs.append({
        "ncols": ncols,
        "nrows": len(out_rows),
        "vlines": vlines0,
        "mc_bars": mc_bars,
        "rules": [{str(c): list(v) for c, v in rr.items()} for rr in row_rules],
        "bottom": {str(c): list(v) for c, v in bottom.items()},
        "widths": [c["width"] for c in cols],
        "x": [c["x"] for c in cols],
        "aligns": [c["align"] for c in cols],
        "target_width": target_width,
        "row_shd": {str(k2): v for k2, v in row_shd.items()},
        "cell_shd": cell_shd,
        "font_size": 12.0,
        "header_plain": True,
        "max_width": self.textwidth - ctx.left,
    })
    self._emit_block(tex, ctx)
    self._stat("data_tables")


def _parse_table_cell(self, c_src: str) -> dict[str, Any]:
    cell = {"content": c_src, "colspan": 1, "rowspan": 1, "diag": None, "shd": None, "mc_spec": None, "mc_align": None}
    s = c_src.strip()
    changed = True
    guard = 0
    while changed and guard < 6:
        guard += 1
        changed = False
        m = re.match(r"^\\cellcolor\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}", s)
        if m:
            cell["shd"] = self.color_hex(m.group(2), m.group(1))
            s = s[m.end():].strip()
            changed = True
            continue
        if s.startswith("\\multicolumn"):
            a, k = read_arg(s, len("\\multicolumn"))
            b, k = read_arg(s, k)
            c, k = read_arg(s, k)
            if a is not None and c is not None:
                try:
                    cell["colspan"] = max(1, int(self._plain(a)))
                except ValueError:
                    pass
                cell["mc_spec"] = b or "c"
                cols_mc, _ = self._parse_colspec(b or "c")
                cell["mc_align"] = cols_mc[0]["align"] if cols_mc else "c"
                s = (c + s[k:]).strip()
                changed = True
                continue
        if s.startswith("\\multirow") or s.startswith("\\multirowcell"):
            nm = "\\multirowcell" if s.startswith("\\multirowcell") else "\\multirow"
            k = len(nm)
            o, k = read_opt(s, k)
            a, k = read_arg(s, k)
            o, k = read_opt(s, k)
            if nm == "\\multirow":
                kk = skip_ws(s, k)
                if kk < len(s) and s[kk] == "*":
                    k = kk + 1
                else:
                    b, k = read_arg(s, k)
                o, k = read_opt(s, k)
            c, k = read_arg(s, k)
            if a is not None and c is not None:
                try:
                    cell["rowspan"] = abs(int(self._plain(a)))
                except ValueError:
                    pass
                s = (c + s[k:]).strip()
                changed = True
                continue
        if s.startswith("\\diagbox") or s.startswith("\\backslashbox") or s.startswith("\\slashbox"):
            nm = re.match(r"\\[A-Za-z]+", s).group(0)
            k = len(nm)
            o, k = read_opt(s, k)
            a, k = read_arg(s, k)
            b, k = read_arg(s, k)
            if a is not None and b is not None:
                cell["diag"] = (a, b)
                s = ""
                changed = True
                continue
    cell["content"] = s
    return cell


# =====================================================================================
# INSTALL
# =====================================================================================
BLOCK_HANDLERS = {
    "par": _h_par, "begin": _h_begin, "end": _h_end,
    "part": _h_heading, "chapter": _h_heading, "section": _h_heading, "subsection": _h_heading,
    "subsubsection": _h_heading, "paragraph": _h_heading, "subparagraph": _h_heading, "addsec": _h_heading,
    "addchap": _h_heading, "minisec": _h_heading,
    "frametitle": _h_frametitle, "maketitle": _h_maketitle, "tableofcontents": _h_toc,
    "newpage": _h_newpage, "clearpage": _h_newpage, "cleardoublepage": _h_newpage, "pagebreak": _h_newpage,
    "eject": _h_newpage, "newpage@": _h_newpage,
    "vspace": _h_vspace, "vskip": _h_vspace, "smallskip": _h_vspace, "medskip": _h_vspace, "bigskip": _h_vspace,
    "addvspace": _h_vspace, "vfill": _h_vspace,
    "label": _h_label, "caption": _h_caption, "captionof": _h_caption,
    "includegraphics": _h_includegraphics,
    "choice": _h_choice, "choiceA": _h_choice, "choiceB": _h_choice, "choiceC": _h_choice, "choiceD": _h_choice,
    "boncot": _h_choice, "haicot": _h_choice, "motcot": _h_choice, "bonpa": _h_choice, "bonPA": _h_choice,
    "haipa": _h_choice, "motpa": _h_choice, "fourch": _h_choice, "twoch": _h_choice, "onech": _h_choice,
    "choiceTF": _h_choice_tf, "choiceTFt": _h_choice_tf, "choiceTFn": _h_choice_tf,
    "immini": _h_immini, "impicinpar": _h_immini, "imminiR": _h_immini,
    "loigiai": _h_loigiai, "giai": _h_loigiai, "huongdan": _h_loigiai, "shortans": _h_loigiai, "dapso": _h_loigiai,
    "traloi": _h_loigiai,
    "cau": _h_question, "Cau": _h_question, "bai": _h_question, "Bai": _h_question, "question": _h_question,
    "item": _h_item_stray, "hrule": _h_hrule, "rule": _h_hrule, "hline": _h_hrule,
    "appendix": _h_appendix, "twocolumn": _h_columns, "onecolumn": _h_columns, "columnbreak": _h_columns,
    "tikz": _h_tikz_inline, "noindent": _h_skip, "indent": _h_skip, "pause": _h_skip,
}

INLINE_KNOWN = set(FMT_CMDS) | set(DECLARATIONS) | {
    "textcolor", "color", "colorbox", "fcolorbox", "footnote", "ref", "eqref", "cite", "url", "href", "label",
    "hspace", "hfill", "dotfill", "hrulefill", "quad", "qquad", "today", "ce", "SI", "si", "num", "ensuremath",
}


def install_parts(cls: Any) -> None:
    funcs = [
        _inline, _inline_math, _wrap_fmt, _apply_fmt, _inline_cmd, _accent, _plain, _inline_env, _read_n_braced,
        _question_number, _math, _math_plain_arg, _math_simple_symbolic, _math_walk, _math_cmd, _math_env,
        _strip_hlines, _math_to_plain, _chem_to_math, _si_unit, _si_num, _siunitx, _set_pending,
        _choice_options_render, _emit_choice, _options_rows, _config_cmd, _try_choice_like, _plain_with_fields, _env, _append_qed,
        _theorem, _enum_format, _list, _inline_list_text, _choices_env, _tasks, _tabbing, _bibliography,
        _parse_colspec, _split_cells, _parse_rule, _table, _parse_table_cell, _minipage_row,
    ]
    for f in funcs:
        setattr(cls, f.__name__, f)
    cls._BLOCK_HANDLERS = BLOCK_HANDLERS
    cls._CONFIG_CMDS = CONFIG_CMDS
    cls._INLINE_KNOWN = INLINE_KNOWN
    cls.float_kind = None
    cls._config_only = False
    cls._pending_vspace = 0
    cls.tasks_bold = False
