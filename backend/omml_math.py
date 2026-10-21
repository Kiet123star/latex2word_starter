"""Pure-Python LaTeX Math to Office Math Markup Language (OMML) Converter (`backend/omml_math.py`).

Converts inline ($...$) and display (\\[...\\], $$...$$, equation/align/cases/matrix/binom/xrightarrow) LaTeX math
directly into native Microsoft Word <m:oMath> / <m:oMathPara> XML elements.
Includes strict XML 1.0 character sanitization so malformed or decompiled PDF/LaTeX strings never trigger XMLSyntaxError.
"""
from __future__ import annotations

import re
from xml.etree import ElementTree as ET
from docx.oxml import parse_xml

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ET.register_namespace("m", M_NS)
ET.register_namespace("w", W_NS)

# Matches all characters forbidden in XML 1.0 (NULL \x00, control chars \x01-\x1f except \t\n\r,
# DEL/C1 control chars \x7f-\x9f, surrogates, non-characters, and Unicode Variation Selectors \ufe00-\ufe0f)
_ILLEGAL_XML_CHARS_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ud800-\udfff\ufdd0-\ufdef\ufffe\uffff\ufe00-\ufe0f]"
)


def sanitize_xml_text(s: str) -> str:
    """Strip all characters illegal in XML 1.0 so lxml.etree.fromstring / python-docx never raises XMLSyntaxError."""
    if not s:
        return ""
    return _ILLEGAL_XML_CHARS_RE.sub("", s)


LATEX_SYMBOLS: dict[str, str] = {
    r"\infty": "∞",
    r"\pm": "±",
    r"\mp": "∓",
    r"\times": "×",
    r"\cdot": "⋅",
    r"\div": "÷",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\neq": "≠",
    r"\ne": "≠",
    r"\approx": "≈",
    r"\equiv": "≡",
    r"\sim": "∼",
    r"\simeq": "≃",
    r"\cong": "≅",
    r"\propto": "∝",
    r"\in": "∈",
    r"\notin": "∉",
    r"\ni": "∋",
    r"\subset": "⊂",
    r"\subseteq": "⊆",
    r"\supset": "⊃",
    r"\supseteq": "⊇",
    r"\cup": "∪",
    r"\cap": "∩",
    r"\setminus": "∖",
    r"\emptyset": "∅",
    r"\varnothing": "∅",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\nexists": "∄",
    r"\lor": "∨",
    r"\vee": "∨",
    r"\land": "∧",
    r"\wedge": "∧",
    r"\neg": "¬",
    r"\lnot": "¬",
    r"\partial": "∂",
    r"\nabla": "∇",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\leftarrow": "←",
    r"\leftrightarrow": "↔",
    r"\Rightarrow": "⇒",
    r"\Leftarrow": "⇐",
    r"\Leftrightarrow": "⇔",
    r"\iff": "⇔",
    r"\implies": "⇒",
    r"\mapsto": "↦",
    r"\uparrow": "↑",
    r"\downarrow": "↓",
    r"\mid": " ∣ ",
    r"\colon": ":",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ϵ",
    r"\varepsilon": "ε",
    r"\zeta": "ζ",
    r"\eta": "η",
    r"\theta": "θ",
    r"\vartheta": "ϑ",
    r"\iota": "ι",
    r"\kappa": "κ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\nu": "ν",
    r"\xi": "ξ",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\upsilon": "υ",
    r"\phi": "ϕ",
    r"\varphi": "φ",
    r"\chi": "χ",
    r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ",
    r"\Delta": "Δ",
    r"\Theta": "Θ",
    r"\Lambda": "Λ",
    r"\Xi": "Ξ",
    r"\Pi": "Π",
    r"\Sigma": "Σ",
    r"\Upsilon": "Υ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
    r"\ell": "ℓ",
    r"\hbar": "ℏ",
    r"\angle": "∠",
    r"\perp": "⊥",
    r"\parallel": "∥",
    r"\triangle": "△",
    r"\circ": "∘",
    r"\degree": "°",
    r"\dots": "…",
    r"\cdots": "⋯",
    r"\ldots": "…",
    r"\vdots": "⋮",
    r"\ddots": "⋱",
    r"\prime": "′",
    r"\Vert": "‖",
    r"\|": "‖",
    r"\{": "{",
    r"\}": "}",
    r"\%": "%",
}

BLACKBOARD_BOLD: dict[str, str] = {
    "R": "ℝ",
    "N": "ℕ",
    "Z": "ℤ",
    "Q": "ℚ",
    "C": "ℂ",
    "P": "ℙ",
    "E": "𝔼",
    "F": "𝔽",
    "K": "𝕂",
}

FUNC_NAMES: dict[str, str] = {
    r"\sin": "sin",
    r"\cos": "cos",
    r"\tan": "tan",
    r"\cot": "cot",
    r"\arcsin": "arcsin",
    r"\arccos": "arccos",
    r"\arctan": "arctan",
    r"\sinh": "sinh",
    r"\cosh": "cosh",
    r"\tanh": "tanh",
    r"\ln": "ln",
    r"\log": "log",
    r"\lg": "lg",
    r"\exp": "exp",
    r"\det": "det",
    r"\gcd": "gcd",
    r"\lcm": "lcm",
    r"\deg": "deg",
    r"\dim": "dim",
    r"\ker": "ker",
    r"\span": "span",
    r"\Span": "Span",
    r"\Im": "Im",
    r"\Re": "Re",
    r"\rank": "rank",
    r"\rk": "rk",
    r"\tr": "tr",
}

LIM_OPS: dict[str, str] = {
    r"\lim": "lim",
    r"\max": "max",
    r"\min": "min",
    r"\sup": "sup",
    r"\inf": "inf",
}

NARY_OPS: dict[str, str] = {
    r"\int": "∫",
    r"\iint": "∬",
    r"\iiint": "∭",
    r"\oint": "∮",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\bigcup": "⋃",
    r"\bigcap": "⋂",
}

ACCENT_MAP: dict[str, str] = {
    r"\vec": "→",
    r"\overrightarrow": "→",
    r"\hat": "^",
    r"\widehat": "^",
    r"\bar": "¯",
    r"\tilde": "~",
    r"\widetilde": "~",
    r"\dot": "˙",
    r"\ddot": "¨",
}

MATRIX_DELIMS: dict[str, tuple[str, str]] = {
    "bmatrix": ("[", "]"),
    "bmatrix*": ("[", "]"),
    "pmatrix": ("(", ")"),
    "pmatrix*": ("(", ")"),
    "Bmatrix": ("{", "}"),
    "Bmatrix*": ("{", "}"),
    "vmatrix": ("|", "|"),
    "vmatrix*": ("|", "|"),
    "Vmatrix": ("‖", "‖"),
    "Vmatrix*": ("‖", "‖"),
}

FONT_SIZE_CMDS: dict[str, float] = {
    "Huge": 22.0,
    "huge": 19.0,
    "LARGE": 17.0,
    "Large": 15.0,
    "large": 13.5,
    "normalsize": 12.0,
    "small": 11.0,
    "footnotesize": 10.0,
    "scriptsize": 9.0,
    "tiny": 8.0,
}


def _m(tag: str) -> ET.Element:
    return ET.Element(f"{{{M_NS}}}{tag}")


def _w(tag: str) -> ET.Element:
    return ET.Element(f"{{{W_NS}}}{tag}")


def _make_run(text: str, *, normal_text: bool = False, bold: bool = False, italic: bool | None = None) -> ET.Element:
    text = sanitize_xml_text(text)
    r = _m("r")
    rPr_m = _m("rPr")
    has_m_pr = False
    if normal_text:
        nor = _m("nor")
        rPr_m.append(nor)
        has_m_pr = True
    elif italic is False:
        sty = _m("sty")
        sty.set(f"{{{M_NS}}}val", "p")
        rPr_m.append(sty)
        has_m_pr = True
    if has_m_pr:
        r.append(rPr_m)

    rPr_w = _w("rPr")
    rFonts = _w("rFonts")
    rFonts.set(f"{{{W_NS}}}ascii", "Cambria Math")
    rFonts.set(f"{{{W_NS}}}hAnsi", "Cambria Math")
    rPr_w.append(rFonts)
    if bold:
        rPr_w.append(_w("b"))
    if italic is False and not normal_text:
        i_el = _w("i")
        i_el.set(f"{{{W_NS}}}val", "0")
        rPr_w.append(i_el)
    r.append(rPr_w)

    t = _m("t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _wrap_in_delimiter(inner_elements: list[ET.Element], beg_chr: str, end_chr: str) -> ET.Element:
    """Wrap one or more OMML elements in an <m:d> delimiter element."""
    beg_chr = sanitize_xml_text(beg_chr)
    end_chr = sanitize_xml_text(end_chr)
    d = _m("d")
    dPr = _m("dPr")
    b_el = _m("begChr")
    b_el.set(f"{{{M_NS}}}val", beg_chr)
    dPr.append(b_el)
    e_el = _m("endChr")
    e_el.set(f"{{{M_NS}}}val", end_chr)
    dPr.append(e_el)
    d.append(dPr)
    e_box = _m("e")
    for el in inner_elements or [_make_run("", normal_text=True)]:
        e_box.append(el)
    d.append(e_box)
    return d


class _LatexMathParser:
    def __init__(self, src: str) -> None:
        s = sanitize_xml_text(src).strip()
        # Extract \tag{...} or \tag*{...} if present
        self.tag: str | None = None
        m_tag = re.search(r"\\tag\*?\{([^{}]*)\}", s)
        if m_tag:
            self.tag = m_tag.group(1).strip()
            s = (s[: m_tag.start()] + s[m_tag.end() :]).strip()

        # Convert legacy {n \choose k} to \binom{n}{k}
        s = re.sub(r"\{\s*([^{}]+?)\s*\\choose\s*([^{}]+?)\s*\}", r"\\binom{\1}{\2}", s)

        s = re.sub(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b", "", s)
        s = re.sub(r"\\(?:quad|qquad)\b", "  ", s)
        s = re.sub(r"\\(?:,|;|:|!)\b", " ", s)
        s = s.replace(r"\,", " ").replace(r"\;", " ").replace(r"\!", "")
        self.s = s
        self.pos = 0
        self.n = len(s)

    def skip_ws(self) -> None:
        while self.pos < self.n and self.s[self.pos] in " \t\r\n":
            self.pos += 1

    def read_command(self) -> str:
        start = self.pos
        self.pos += 1
        if self.pos < self.n and not self.s[self.pos].isalpha():
            self.pos += 1
            return self.s[start:self.pos]
        while self.pos < self.n and self.s[self.pos].isalpha():
            self.pos += 1
        return self.s[start:self.pos]

    def read_braced_raw(self) -> str:
        self.skip_ws()
        if self.pos >= self.n:
            return ""
        if self.s[self.pos] != "{":
            if self.s[self.pos] == "\\":
                return self.read_command()
            ch = self.s[self.pos]
            self.pos += 1
            return ch
        self.pos += 1
        depth = 1
        start = self.pos
        while self.pos < self.n and depth > 0:
            if self.s[self.pos] == "\\":
                self.pos += 2
                continue
            if self.s[self.pos] == "{":
                depth += 1
            elif self.s[self.pos] == "}":
                depth -= 1
                if depth == 0:
                    res = self.s[start : self.pos]
                    self.pos += 1
                    return res
            self.pos += 1
        return self.s[start : self.pos]

    def read_bracket_raw(self) -> str | None:
        self.skip_ws()
        if self.pos >= self.n or self.s[self.pos] != "[":
            return None
        self.pos += 1
        depth = 1
        start = self.pos
        while self.pos < self.n and depth > 0:
            if self.s[self.pos] == "[":
                depth += 1
            elif self.s[self.pos] == "]":
                depth -= 1
                if depth == 0:
                    res = self.s[start : self.pos]
                    self.pos += 1
                    return res
            self.pos += 1
        return self.s[start : self.pos]

    def parse_group_elements(self) -> list[ET.Element]:
        raw = self.read_braced_raw()
        sub = _LatexMathParser(raw)
        return sub.parse_all()

    def _parse_paren_group(self, open_ch: str, close_ch: str) -> list[ET.Element]:
        """Parse a balanced (...) group so exponents like (x-1)^2 wrap the entire group."""
        start_pos = self.pos
        self.pos += 1
        depth = 1
        inner_start = self.pos
        while self.pos < self.n and depth > 0:
            if self.s[self.pos] == "\\":
                if self.s.startswith(r"\left", self.pos) or self.s.startswith(r"\right", self.pos):
                    break
                self.pos += 2
                continue
            if self.s[self.pos] == open_ch:
                depth += 1
            elif self.s[self.pos] == close_ch:
                depth -= 1
                if depth == 0:
                    inner_raw = self.s[inner_start : self.pos]
                    self.pos += 1
                    inner_els = _LatexMathParser(inner_raw).parse_all()
                    return (
                        [_make_run(open_ch, normal_text=True)]
                        + inner_els
                        + [_make_run(close_ch, normal_text=True)]
                    )
            self.pos += 1
        self.pos = start_pos + 1
        return [_make_run(open_ch, normal_text=True)]

    def _parse_escaped_brace_group(self) -> list[ET.Element]:
        """Parse a balanced \\{ ... \\} set or span group into crisp brace runs."""
        start_pos = self.pos
        depth = 1
        inner_start = self.pos
        while self.pos < self.n and depth > 0:
            if self.s.startswith(r"\{", self.pos):
                depth += 1
                self.pos += 2
                continue
            if self.s.startswith(r"\}", self.pos):
                depth -= 1
                if depth == 0:
                    inner_raw = self.s[inner_start : self.pos]
                    self.pos += 2
                    inner_els = _LatexMathParser(inner_raw).parse_all()
                    return [_make_run("{", italic=False)] + inner_els + [_make_run("}", italic=False)]
                self.pos += 2
                continue
            if self.s[self.pos] == "\\":
                self.pos += 2
                continue
            self.pos += 1
        self.pos = start_pos
        return [_make_run("{", normal_text=True)]

    def parse_single_atom(self) -> list[ET.Element]:
        self.skip_ws()
        if self.pos >= self.n:
            return [_make_run("")]

        ch = self.s[self.pos]
        if ch == "{":
            return self.parse_group_elements()

        if ch == "(":
            return self._parse_paren_group("(", ")")

        if ch in ")[ ]|":
            self.pos += 1
            return [_make_run(ch, normal_text=True)]

        if ch == "\\":
            saved = self.pos
            cmd = self.read_command()

            if cmd == r"\right":
                self.pos = saved
                return []

            if cmd == r"\left":
                return self._parse_left_right()

            if cmd == r"\{":
                return self._parse_escaped_brace_group()

            if cmd == r"\}":
                return [_make_run("}", normal_text=True)]

            if cmd in (r"\frac", r"\dfrac", r"\tfrac", r"\cfrac"):
                num_els = self.parse_group_elements()
                den_els = self.parse_group_elements()
                f = _m("f")
                num = _m("num")
                for el in num_els or [_make_run("")]:
                    num.append(el)
                den = _m("den")
                for el in den_els or [_make_run("")]:
                    den.append(el)
                f.append(num)
                f.append(den)
                return [f]

            if cmd in (r"\binom", r"\dbinom", r"\tbinom"):
                num_els = self.parse_group_elements()
                den_els = self.parse_group_elements()
                f = _m("f")
                fPr = _m("fPr")
                t_el = _m("type")
                t_el.set(f"{{{M_NS}}}val", "noBar")
                fPr.append(t_el)
                f.append(fPr)
                num = _m("num")
                for el in num_els or [_make_run("")]:
                    num.append(el)
                den = _m("den")
                for el in den_els or [_make_run("")]:
                    den.append(el)
                f.append(num)
                f.append(den)
                return [_wrap_in_delimiter([f], "(", ")")]

            if cmd in (r"\xrightarrow", r"\xleftarrow"):
                sub_raw = self.read_bracket_raw()
                sup_els = self.parse_group_elements()
                arrow_str = " ⎯⎯⎯⎯→ " if cmd == r"\xrightarrow" else " ←⎯⎯⎯⎯ "
                base_node: ET.Element = _make_run(arrow_str, normal_text=True)
                if sup_els:
                    limUpp = _m("limUpp")
                    e_up = _m("e")
                    e_up.append(base_node)
                    lim_up = _m("lim")
                    for el in sup_els:
                        lim_up.append(el)
                    limUpp.append(e_up)
                    limUpp.append(lim_up)
                    base_node = limUpp
                if sub_raw:
                    limLow = _m("limLow")
                    e_lo = _m("e")
                    e_lo.append(base_node)
                    lim_lo = _m("lim")
                    for el in _LatexMathParser(sub_raw).parse_all():
                        lim_lo.append(el)
                    limLow.append(e_lo)
                    limLow.append(lim_lo)
                    base_node = limLow
                return [base_node]

            if cmd in (r"\overset", r"\underset"):
                lim_els = self.parse_group_elements()
                base_els = self.parse_group_elements()
                node = _m("limUpp" if cmd == r"\overset" else "limLow")
                e_box = _m("e")
                for el in base_els or [_make_run("")]:
                    e_box.append(el)
                lim_box = _m("lim")
                for el in lim_els or [_make_run("")]:
                    lim_box.append(el)
                node.append(e_box)
                node.append(lim_box)
                return [node]

            if cmd == r"\sqrt":
                deg_raw = self.read_bracket_raw()
                body_els = self.parse_group_elements()
                rad = _m("rad")
                radPr = _m("radPr")
                degHide = _m("degHide")
                degHide.set(f"{{{M_NS}}}val", "0" if deg_raw else "1")
                radPr.append(degHide)
                rad.append(radPr)
                deg = _m("deg")
                if deg_raw:
                    for el in _LatexMathParser(deg_raw).parse_all():
                        deg.append(el)
                rad.append(deg)
                e = _m("e")
                for el in body_els or [_make_run("")]:
                    e.append(el)
                rad.append(e)
                return [rad]

            if cmd in (r"\overline", r"\underline"):
                body_els = self.parse_group_elements()
                bar = _m("bar")
                barPr = _m("barPr")
                pos_el = _m("pos")
                pos_el.set(f"{{{M_NS}}}val", "top" if cmd == r"\overline" else "bot")
                barPr.append(pos_el)
                bar.append(barPr)
                e = _m("e")
                for el in body_els or [_make_run("")]:
                    e.append(el)
                bar.append(e)
                return [bar]

            if cmd == r"\mathbb":
                arg = self.read_braced_raw().strip()
                return [_make_run(BLACKBOARD_BOLD.get(arg, arg), italic=False)]

            if cmd in (
                r"\text", r"\textrm", r"\textbf", r"\textit",
                r"\mathrm", r"\mathbf", r"\mathcal", r"\operatorname",
            ):
                arg = self.read_braced_raw()
                is_bold = cmd in (r"\textbf", r"\mathbf")
                is_nor = cmd in (r"\text", r"\textrm", r"\mathrm", r"\operatorname")
                return [_make_run(arg, normal_text=is_nor, bold=is_bold, italic=False if is_bold else None)]

            if cmd in ACCENT_MAP:
                body_els = self.parse_group_elements()
                acc = _m("acc")
                accPr = _m("accPr")
                chr_el = _m("chr")
                chr_el.set(f"{{{M_NS}}}val", ACCENT_MAP[cmd])
                accPr.append(chr_el)
                acc.append(accPr)
                e = _m("e")
                for el in body_els or [_make_run("")]:
                    e.append(el)
                acc.append(e)
                return [acc]

            if cmd in NARY_OPS:
                return [self._parse_nary(NARY_OPS[cmd])]

            if cmd in LIM_OPS:
                return [self._parse_lim_op(LIM_OPS[cmd])]

            if cmd in FUNC_NAMES:
                return [_make_run(FUNC_NAMES[cmd], normal_text=True)]

            if cmd in (r"\heva", r"\hoac"):
                body_raw = self.read_braced_raw()
                delim = "{" if cmd == r"\heva" else "["
                return [self._build_cases(body_raw, left_chr=delim)]

            if cmd == r"\begin":
                env_name = self.read_braced_raw().strip()
                return [self._parse_environment(env_name)]

            if cmd in LATEX_SYMBOLS:
                sym = LATEX_SYMBOLS[cmd]
                use_nor = cmd in (
                    r"\mid", r"\Rightarrow", r"\Leftarrow", r"\Leftrightarrow",
                    r"\iff", r"\implies", r"\lor", r"\vee", r"\land", r"\wedge",
                )
                return [_make_run(sym, normal_text=use_nor, italic=False)]

            clean_cmd = cmd.lstrip("\\")
            return [_make_run(clean_cmd)]

        self.pos += 1
        if ch == "'":
            return [_make_run("′")]
        if ch == "-":
            return [_make_run("−")]
        if ch == "·":
            return [_make_run("⋅", italic=False)]
        if ch == "=":
            return [_make_run("=", normal_text=True)]
        if ch == "*":
            return [_make_run("∗", italic=False)]
        return [_make_run(ch)]

    def _parse_left_right(self) -> list[ET.Element]:
        self.skip_ws()
        left_chr = "("
        if self.pos < self.n:
            if self.s[self.pos] == "\\":
                c = self.read_command()
                left_chr = LATEX_SYMBOLS.get(c, c.lstrip("\\"))
            else:
                left_chr = self.s[self.pos]
                self.pos += 1
        if left_chr == ".":
            left_chr = ""

        inner_els: list[ET.Element] = []
        while self.pos < self.n:
            self.skip_ws()
            if self.s.startswith(r"\right", self.pos):
                self.read_command()
                break
            el_list = self.parse_next_with_scripts()
            if not el_list:
                break
            inner_els.extend(el_list)

        self.skip_ws()
        right_chr = ")"
        if self.pos < self.n:
            if self.s[self.pos] == "\\":
                c = self.read_command()
                right_chr = LATEX_SYMBOLS.get(c, c.lstrip("\\"))
            else:
                right_chr = self.s[self.pos]
                self.pos += 1
        if right_chr == ".":
            right_chr = ""

        if left_chr in ("(", "[", "{", "|", "‖") and right_chr in (")", "]", "}", "|", "‖", ""):
            return [_wrap_in_delimiter(inner_els, left_chr, right_chr)]

        res: list[ET.Element] = []
        if left_chr:
            res.append(_make_run(left_chr, normal_text=True))
        res.extend(inner_els)
        if right_chr:
            res.append(_make_run(right_chr, normal_text=True))
        return res

    def _parse_nary(self, symbol: str) -> ET.Element:
        sub_els: list[ET.Element] = []
        sup_els: list[ET.Element] = []
        for _ in range(2):
            self.skip_ws()
            if self.pos < self.n and self.s[self.pos] == "_":
                self.pos += 1
                sub_els = self.parse_single_atom()
            elif self.pos < self.n and self.s[self.pos] == "^":
                self.pos += 1
                sup_els = self.parse_single_atom()

        operand_els: list[ET.Element] = []
        while self.pos < self.n:
            self.skip_ws()
            if self.pos >= self.n or self.s[self.pos] in "=," or self.s.startswith(r"\right", self.pos):
                break
            nxt = self.parse_next_with_scripts()
            if not nxt:
                break
            operand_els.extend(nxt)
        if not operand_els:
            operand_els = [_make_run("")]

        nary = _m("nary")
        naryPr = _m("naryPr")
        chr_el = _m("chr")
        chr_el.set(f"{{{M_NS}}}val", symbol)
        naryPr.append(chr_el)
        limLoc = _m("limLoc")
        limLoc.set(f"{{{M_NS}}}val", "subSup")
        naryPr.append(limLoc)
        subHide = _m("subHide")
        subHide.set(f"{{{M_NS}}}val", "0" if sub_els else "1")
        naryPr.append(subHide)
        supHide = _m("supHide")
        supHide.set(f"{{{M_NS}}}val", "0" if sup_els else "1")
        naryPr.append(supHide)
        nary.append(naryPr)

        sub = _m("sub")
        for el in sub_els:
            sub.append(el)
        sup = _m("sup")
        for el in sup_els:
            sup.append(el)
        e = _m("e")
        for el in operand_els:
            e.append(el)
        nary.append(sub)
        nary.append(sup)
        nary.append(e)
        return nary

    def _parse_lim_op(self, op_name: str) -> ET.Element:
        self.skip_ws()
        if self.pos < self.n and self.s[self.pos] == "_":
            self.pos += 1
            lim_els = self.parse_single_atom()
            limLow = _m("limLow")
            e = _m("e")
            e.append(_make_run(op_name, italic=False))
            lim = _m("lim")
            for el in lim_els or [_make_run("")]:
                lim.append(el)
            limLow.append(e)
            limLow.append(lim)
            return limLow
        return _make_run(op_name, italic=False)

    def _parse_environment(self, env_name: str) -> ET.Element:
        end_token = f"\\end{{{env_name}}}"
        idx = self.s.find(end_token, self.pos)
        if idx == -1:
            body = self.s[self.pos :]
            self.pos = self.n
        else:
            body = self.s[self.pos : idx]
            self.pos = idx + len(end_token)

        if env_name in ("cases", "dcases"):
            return self._build_cases(body, left_chr="{")

        if "matrix" in env_name or env_name == "array":
            return self._build_matrix(env_name, body)

        return self._build_aligned(body)

    def _build_cases(self, body: str, left_chr: str = "{") -> ET.Element:
        """Build a native OMML system of equations with a single stretchy left brace <m:d> + <m:m>."""
        raw_rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        parsed_rows: list[list[str]] = []
        for r_str in raw_rows:
            if "&" in r_str:
                cols = [c.strip() for c in r_str.split("&")]
            else:
                m_eqnum = re.match(r"^(.*?)(?:\\quad|\\qquad|\s{2,})(\(\d+\))\s*$", r_str)
                if m_eqnum:
                    cols = [m_eqnum.group(1).strip(), m_eqnum.group(2).strip()]
                else:
                    cols = [r_str]
            parsed_rows.append(cols)

        max_cols = max((len(r) for r in parsed_rows), default=1)
        m_el = _m("m")
        mPr = _m("mPr")
        mcs = _m("mcs")
        mc = _m("mc")
        mcPr = _m("mcPr")
        cnt = _m("count")
        cnt.set(f"{{{M_NS}}}val", str(max_cols))
        jc = _m("mcJc")
        jc.set(f"{{{M_NS}}}val", "left")
        mcPr.append(cnt)
        mcPr.append(jc)
        mc.append(mcPr)
        mcs.append(mc)
        mPr.append(mcs)
        m_el.append(mPr)

        for cols in parsed_rows:
            mr = _m("mr")
            for c_i in range(max_cols):
                e = _m("e")
                c_str = cols[c_i] if c_i < len(cols) else ""
                if c_str:
                    for el in _LatexMathParser(c_str).parse_all():
                        e.append(el)
                else:
                    e.append(_make_run("", normal_text=True))
                mr.append(e)
            m_el.append(mr)

        return _wrap_in_delimiter([m_el], left_chr, "")

    def _build_matrix(self, env_name: str, body: str) -> ET.Element:
        """Build a native OMML matrix <m:m>, wrapped in <m:d> brackets for bmatrix/pmatrix/vmatrix/Bmatrix."""
        if env_name == "array":
            body = re.sub(r"^\s*\{[^{}]*\}", "", body, count=1)
        raw_rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        parsed_rows = [[c.strip() for c in r_str.split("&")] for r_str in raw_rows]
        max_cols = max((len(r) for r in parsed_rows), default=1)

        m_el = _m("m")
        mPr = _m("mPr")
        mcs = _m("mcs")
        mc = _m("mc")
        mcPr = _m("mcPr")
        cnt = _m("count")
        cnt.set(f"{{{M_NS}}}val", str(max_cols))
        jc = _m("mcJc")
        jc.set(f"{{{M_NS}}}val", "center")
        mcPr.append(cnt)
        mcPr.append(jc)
        mc.append(mcPr)
        mcs.append(mc)
        mPr.append(mcs)
        m_el.append(mPr)

        for cols in parsed_rows:
            mr = _m("mr")
            for c_i in range(max_cols):
                e = _m("e")
                c_str = cols[c_i] if c_i < len(cols) else ""
                if c_str:
                    for el in _LatexMathParser(c_str).parse_all():
                        e.append(el)
                else:
                    e.append(_make_run("", normal_text=True))
                mr.append(e)
            m_el.append(mr)

        if env_name in MATRIX_DELIMS:
            beg_c, end_c = MATRIX_DELIMS[env_name]
            return _wrap_in_delimiter([m_el], beg_c, end_c)
        return m_el

    def _build_aligned(self, body: str) -> ET.Element:
        """Build a multi-line aligned equation array using a left-aligned <m:m> matrix."""
        raw_rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        if not raw_rows:
            return _make_run("")
        parsed_rows = [[c.strip() for c in r_str.split("&")] for r_str in raw_rows]
        max_cols = max((len(r) for r in parsed_rows), default=1)

        m_el = _m("m")
        mPr = _m("mPr")
        mcs = _m("mcs")
        mc = _m("mc")
        mcPr = _m("mcPr")
        cnt = _m("count")
        cnt.set(f"{{{M_NS}}}val", str(max_cols))
        jc = _m("mcJc")
        jc.set(f"{{{M_NS}}}val", "left")
        mcPr.append(cnt)
        mcPr.append(jc)
        mc.append(mcPr)
        mcs.append(mc)
        mPr.append(mcs)
        m_el.append(mPr)

        for cols in parsed_rows:
            mr = _m("mr")
            for c_i in range(max_cols):
                e = _m("e")
                c_str = cols[c_i] if c_i < len(cols) else ""
                if c_str:
                    for el in _LatexMathParser(c_str).parse_all():
                        e.append(el)
                else:
                    e.append(_make_run(" ", normal_text=True))
                mr.append(e)
            m_el.append(mr)
        return m_el

    def parse_next_with_scripts(self) -> list[ET.Element]:
        self.skip_ws()
        if self.pos >= self.n:
            return []
        base_els = self.parse_single_atom()
        if not base_els:
            return []

        while self.pos < self.n and self.s[self.pos] == "'":
            self.pos += 1
            base_els.append(_make_run("′"))

        sub_els: list[ET.Element] | None = None
        sup_els: list[ET.Element] | None = None

        for _ in range(2):
            self.skip_ws()
            if self.pos < self.n and self.s[self.pos] == "_" and sub_els is None:
                self.pos += 1
                sub_els = self.parse_single_atom()
            elif self.pos < self.n and self.s[self.pos] == "^" and sup_els is None:
                self.pos += 1
                sup_els = self.parse_single_atom()

        if sub_els is None and sup_els is None:
            return base_els

        if sub_els is not None and sup_els is not None:
            node = _m("sSubSup")
            e = _m("e")
            for el in base_els:
                e.append(el)
            sub = _m("sub")
            for el in sub_els:
                sub.append(el)
            sup = _m("sup")
            for el in sup_els:
                sup.append(el)
            node.append(e)
            node.append(sub)
            node.append(sup)
            return [node]
        elif sub_els is not None:
            node = _m("sSub")
            e = _m("e")
            for el in base_els:
                e.append(el)
            sub = _m("sub")
            for el in sub_els:
                sub.append(el)
            node.append(e)
            node.append(sub)
            return [node]
        else:
            assert sup_els is not None
            node = _m("sSup")
            e = _m("e")
            for el in base_els:
                e.append(el)
            sup = _m("sup")
            for el in sup_els:
                sup.append(el)
            node.append(e)
            node.append(sup)
            return [node]

    def parse_all(self) -> list[ET.Element]:
        elements: list[ET.Element] = []
        while self.pos < self.n:
            self.skip_ws()
            if self.pos >= self.n:
                break
            if self.s.startswith(r"\right", self.pos):
                break
            nxt = self.parse_next_with_scripts()
            if not nxt:
                break
            elements.extend(nxt)
        if self.tag:
            elements.append(_make_run(f"   ({self.tag})", normal_text=True))
        return elements


def latex_to_omml_element(latex_math: str, display: bool = False):
    """Convert a LaTeX math string into a python-docx OxmlElement (<m:oMath> or <m:oMathPara>)."""
    clean = sanitize_xml_text(latex_math).strip()
    if clean.startswith("$$") and clean.endswith("$$"):
        clean = clean[2:-2].strip()
        display = True
    elif clean.startswith("\\[") and clean.endswith("\\]"):
        clean = clean[2:-2].strip()
        display = True
    elif clean.startswith("$") and clean.endswith("$"):
        clean = clean[1:-1].strip()

    omath = _m("oMath")
    if clean in ("+", "-", "−", "±"):
        omath.append(_make_run("−" if clean == "-" else clean, normal_text=True))
    else:
        parser = _LatexMathParser(clean)
        children = parser.parse_all()
        if not children:
            omath.append(_make_run(clean or " "))
        else:
            for ch in children:
                omath.append(ch)

    if display:
        para = _m("oMathPara")
        para.append(omath)
        xml_str = sanitize_xml_text(ET.tostring(para, encoding="unicode"))
    else:
        xml_str = sanitize_xml_text(ET.tostring(omath, encoding="unicode"))
    return parse_xml(xml_str)


def latex_to_omml_xml_string(latex_math: str, display: bool = False) -> str:
    """Return raw XML string of <m:oMath> suitable for embedding inside DrawingML textboxes."""
    clean = sanitize_xml_text(latex_math).strip()
    if clean.startswith("$") and clean.endswith("$"):
        clean = clean[1:-1].strip()
    omath = _m("oMath")
    if clean in ("+", "-", "−", "±"):
        omath.append(_make_run("−" if clean == "-" else clean, normal_text=True))
    else:
        parser = _LatexMathParser(clean)
        children = parser.parse_all()
        if not children:
            omath.append(_make_run(clean or " "))
        else:
            for ch in children:
                omath.append(ch)
    if display:
        para = _m("oMathPara")
        para.append(omath)
        return sanitize_xml_text(ET.tostring(para, encoding="unicode"))
    return sanitize_xml_text(ET.tostring(omath, encoding="unicode"))


def _extract_braced_content(s: str, open_brace_idx: int) -> tuple[str, int] | None:
    """Given index of `{`, return `(inner_content, index_after_closing_brace)` supporting arbitrary nesting."""
    if open_brace_idx >= len(s) or s[open_brace_idx] != "{":
        return None
    depth = 1
    i = open_brace_idx + 1
    n = len(s)
    while i < n and depth > 0:
        if s[i] == "\\":
            i += 2
            continue
        if s[i] == "{":
            depth += 1
        elif s[i] == "}":
            depth -= 1
            if depth == 0:
                return s[open_brace_idx + 1 : i], i + 1
        i += 1
    return s[open_brace_idx + 1 :], n


def _hspace_to_nbsp(match: re.Match[str]) -> str:
    """Convert \\hspace*{0.5cm} or \\hspace{2em} into proportional non-breaking spaces."""
    arg = match.group(1).strip()
    m_num = re.search(r"(\d+(?:\.\d+)?)", arg)
    val = float(m_num.group(1)) if m_num else 0.5
    if "cm" in arg:
        spaces = max(2, int(round(val * 6)))
    elif "em" in arg:
        spaces = max(2, int(round(val * 3)))
    elif "pt" in arg:
        spaces = max(1, int(round(val / 3.0)))
    else:
        spaces = 4
    return "\u00A0" * min(spaces, 32)


def append_rich_latex_text_to_paragraph(
    paragraph,
    text: str,
    *,
    default_bold: bool = False,
    default_underline: bool = False,
    default_italic: bool = False,
    font_size_pt: float = 12.0,
) -> None:
    """Parse mixed Vietnamese text, nested formatting (\\textbf, \\underline, \\textit, {\\LARGE ...}), line breaks (\\\\), and $...$ math."""
    from docx.shared import Pt

    s = sanitize_xml_text(text).replace(r"\(", "$").replace(r"\)", "$")
    s = re.sub(r"\\hspace\*?\s*\{([^{}]*)\}", _hspace_to_nbsp, s)
    s = s.replace(r"\textbackslash ", "\\").replace(r"\textbackslash", "\\")

    i = 0
    n = len(s)
    plain_buf: list[str] = []

    def _flush_plain() -> None:
        if not plain_buf:
            return
        raw_frag = "".join(plain_buf)
        plain_buf.clear()
        clean = _clean_plain_tex_fragment(raw_frag)
        if clean:
            r = paragraph.add_run(clean)
            r.bold = default_bold
            r.underline = default_underline
            r.italic = default_italic
            r.font.size = Pt(font_size_pt)

    while i < n:
        # 1. Explicit LaTeX line break \\ or \newline outside math
        if s.startswith(r"\\", i):
            _flush_plain()
            i += 2
            while i < n and s[i] in " \t":
                i += 1
            if i < n and s[i] == "[":
                end_b = s.find("]", i)
                if end_b != -1:
                    i = end_b + 1
            if i < n and s[i] == "\n":
                i += 1
            if s[i:].strip():
                paragraph.add_run().add_break()
            continue

        if s.startswith(r"\newline", i) and (i + 8 == n or not s[i + 8].isalpha()):
            _flush_plain()
            i += 8
            if s[i:].strip():
                paragraph.add_run().add_break()
            continue

        # 2. Escaped literal braces \{ and \} in text mode
        if s.startswith(r"\{", i):
            plain_buf.append("{")
            i += 2
            continue
        if s.startswith(r"\}", i):
            plain_buf.append("}")
            i += 2
            continue

        # 3. Display math \[ ... \] or $$ ... $$
        if s.startswith(r"\[", i):
            end_m = s.find(r"\]", i + 2)
            if end_m != -1:
                _flush_plain()
                math_src = sanitize_xml_text(s[i + 2 : end_m]).strip()
                if math_src:
                    paragraph._p.append(latex_to_omml_element(math_src, display=True))
                i = end_m + 2
                continue

        if s.startswith("$$", i):
            end_m = s.find("$$", i + 2)
            if end_m != -1:
                _flush_plain()
                math_src = sanitize_xml_text(s[i + 2 : end_m]).strip()
                if math_src:
                    paragraph._p.append(latex_to_omml_element(math_src, display=True))
                i = end_m + 2
                continue

        # 4. Inline math $ ... $
        if s[i] == "$":
            end_m = i + 1
            while end_m < n:
                if s[end_m] == "\\":
                    end_m += 2
                    continue
                if s[end_m] == "$":
                    break
                end_m += 1
            if end_m < n and s[end_m] == "$":
                _flush_plain()
                math_src = sanitize_xml_text(s[i + 1 : end_m]).strip()
                if math_src:
                    paragraph._p.append(latex_to_omml_element(math_src, display=False))
                i = end_m + 1
                continue

        # 5. Formatting commands \textbf{...}, \underline{...}, \textit{...}, \emph{...}, \texttt{...}
        m_fmt = re.match(r"^\\(textbf|underline|textit|emph|texttt)\s*\{", s[i:])
        if m_fmt:
            cmd_name = m_fmt.group(1)
            open_idx = i + m_fmt.end() - 1
            extracted = _extract_braced_content(s, open_idx)
            if extracted is not None:
                _flush_plain()
                inner_str, next_i = extracted
                append_rich_latex_text_to_paragraph(
                    paragraph,
                    inner_str,
                    default_bold=default_bold or (cmd_name == "textbf"),
                    default_underline=default_underline or (cmd_name == "underline"),
                    default_italic=default_italic or (cmd_name in ("textit", "emph")),
                    font_size_pt=font_size_pt,
                )
                i = next_i
                continue

        # 6. Braced group in text mode, e.g. {\LARGE \textbf{Bài toán đếm}} or {\large ...} or {\bf ...}
        if s[i] == "{":
            extracted = _extract_braced_content(s, i)
            if extracted is not None:
                _flush_plain()
                inner_str, next_i = extracted
                sub_sz = font_size_pt
                sub_bold = default_bold
                sub_italic = default_italic
                # Check leading size/style switches inside the group
                while True:
                    m_sw = re.match(
                        r"^\s*\\(Huge|huge|LARGE|Large|large|normalsize|small|footnotesize|scriptsize|tiny|bfseries|bf|itshape|it|em)\b\s*",
                        inner_str,
                    )
                    if not m_sw:
                        break
                    sw = m_sw.group(1)
                    if sw in FONT_SIZE_CMDS:
                        sub_sz = FONT_SIZE_CMDS[sw]
                    elif sw in ("bfseries", "bf"):
                        sub_bold = True
                    elif sw in ("itshape", "it", "em"):
                        sub_italic = True
                    inner_str = inner_str[m_sw.end() :]
                append_rich_latex_text_to_paragraph(
                    paragraph,
                    inner_str,
                    default_bold=sub_bold,
                    default_underline=default_underline,
                    default_italic=sub_italic,
                    font_size_pt=sub_sz,
                )
                i = next_i
                continue

        plain_buf.append(s[i])
        i += 1

    _flush_plain()


def _clean_plain_tex_fragment(s: str) -> str:
    s = sanitize_xml_text(s)
    s = re.sub(r"\\(?:dots|ldots|cdots)\b\s*", "…", s)
    s = re.sub(r"\\(?:par|noindent|centering|raggedright|raggedleft|hfill|vfill|hrulefill)\b\s*", "", s)
    s = re.sub(r"\\(?:vspace|hspace)\*?\s*\{[^{}]*\}", "", s)
    s = re.sub(r"\\(?:setstretch|linespread)\s*\{[^{}]*\}", "", s)
    s = re.sub(r"\\(?:small|footnotesize|scriptsize|tiny|large|Large|LARGE|huge|Huge|normalsize|bfseries|itshape)\b\s*", "", s)
    s = re.sub(r"\\(?:begin|end)\{(?:itemize|enumerate|description|quote|quotation|center)\}(?:\[[^\]]*\])?", "", s)
    s = re.sub(r"\\item(?:\s*\[[^\]]*\])?\s*", "", s)
    s = s.replace("``", "“").replace("''", "”").replace("~", " ").replace(r"\&", "&").replace(r"\%", "%")
    s = re.sub(r"[ \t]+", " ", s)
    return sanitize_xml_text(s)
