"""Pure-Python LaTeX Math to Office Math Markup Language (OMML) Converter.

Converts inline ($...$) and display (\\[...\\], $$...$$, equation/align) LaTeX math
directly into native Microsoft Word <m:oMath> / <m:oMathPara> XML elements.
"""
from __future__ import annotations

import re
from xml.etree import ElementTree as ET
from docx.oxml import parse_xml

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ET.register_namespace("m", M_NS)
ET.register_namespace("w", W_NS)

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
    r"\propto": "∝",
    r"\in": "∈",
    r"\notin": "∉",
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
}

FUNC_NAMES = {
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
    r"\deg": "deg",
    r"\dim": "dim",
    r"\ker": "ker",
}

LIM_OPS = {
    r"\lim": "lim",
    r"\max": "max",
    r"\min": "min",
    r"\sup": "sup",
    r"\inf": "inf",
}

NARY_OPS = {
    r"\int": "∫",
    r"\iint": "∬",
    r"\iiint": "∭",
    r"\oint": "∮",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\bigcup": "⋃",
    r"\bigcap": "⋂",
}

ACCENT_MAP = {
    r"\vec": "→",
    r"\overrightarrow": "→",
    r"\hat": "^",
    r"\widehat": "^",
    r"\bar": "¯",
    r"\overline": "¯",
    r"\tilde": "~",
    r"\widetilde": "~",
    r"\dot": "˙",
    r"\ddot": "¨",
}


def _m(tag: str) -> ET.Element:
    return ET.Element(f"{{{M_NS}}}{tag}")


def _w(tag: str) -> ET.Element:
    return ET.Element(f"{{{W_NS}}}{tag}")


def _make_run(text: str, *, normal_text: bool = False, bold: bool = False, italic: bool | None = None) -> ET.Element:
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
    t.text = text
    r.append(t)
    return r


class _LatexMathParser:
    def __init__(self, src: str) -> None:
        s = src.strip()
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
                    res = self.s[start:self.pos]
                    self.pos += 1
                    return res
            self.pos += 1
        return self.s[start:self.pos]

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
                    res = self.s[start:self.pos]
                    self.pos += 1
                    return res
            self.pos += 1
        return self.s[start:self.pos]

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
                    inner_raw = self.s[inner_start:self.pos]
                    self.pos += 1
                    inner_els = _LatexMathParser(inner_raw).parse_all()
                    return (
                        [_make_run(open_ch, normal_text=True)]
                        + inner_els
                        + [_make_run(close_ch, normal_text=True)]
                    )
            self.pos += 1
        # Unbalanced (e.g. half-open interval (-2; 0]): emit literal bracket with normal_text=True
        self.pos = start_pos + 1
        return [_make_run(open_ch, normal_text=True)]

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

            if cmd == r"\mathbb":
                arg = self.read_braced_raw().strip()
                return [_make_run(BLACKBOARD_BOLD.get(arg, arg), italic=False)]

            if cmd in (r"\text", r"\textrm", r"\textbf", r"\textit", r"\mathrm", r"\mathbf", r"\mathcal"):
                arg = self.read_braced_raw()
                is_bold = cmd in (r"\textbf", r"\mathbf")
                is_nor = cmd in (r"\text", r"\textrm", r"\mathrm")
                return [_make_run(arg, normal_text=is_nor, bold=is_bold)]

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
                return [_make_run(FUNC_NAMES[cmd], italic=False)]

            if cmd in (r"\heva", r"\hoac"):
                body_raw = self.read_braced_raw()
                delim = "{" if cmd == r"\heva" else "["
                return [self._build_cases(body_raw, left_chr=delim)]

            if cmd == r"\begin":
                env_name = self.read_braced_raw().strip()
                return [self._parse_environment(env_name)]

            if cmd in LATEX_SYMBOLS:
                return [_make_run(LATEX_SYMBOLS[cmd], italic=False)]

            clean_cmd = cmd.lstrip("\\")
            return [_make_run(clean_cmd)]

        self.pos += 1
        if ch == "'":
            return [_make_run("′")]
        if ch == "-":
            return [_make_run("−")]
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

        # Collect integrand / summand tokens up to '=', ',', ';', or end of expression
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
            body = self.s[self.pos:]
            self.pos = self.n
        else:
            body = self.s[self.pos:idx]
            self.pos = idx + len(end_token)

        if env_name in ("cases", "dcases"):
            return self._build_cases(body, left_chr="{")

        if "matrix" in env_name or env_name == "array":
            return self._build_matrix(env_name, body)

        rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        if len(rows) <= 1:
            els = _LatexMathParser(body.replace("&", " ")).parse_all()
            return els[0] if len(els) == 1 else _make_run(body)
        eqArr = _m("eqArr")
        for row_str in rows:
            e = _m("e")
            for el in _LatexMathParser(row_str.replace("&", " ")).parse_all():
                e.append(el)
            eqArr.append(e)
        return eqArr

    def _build_cases(self, body: str, left_chr: str = "{") -> ET.Element:
        eqArr = _m("eqArr")
        rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        for r_str in rows:
            e_row = _m("e")
            e_row.append(_make_run(left_chr + " ", normal_text=True))
            clean_r = r_str.replace("&", "  ")
            for el in _LatexMathParser(clean_r).parse_all():
                e_row.append(el)
            eqArr.append(e_row)
        return eqArr

    def _build_matrix(self, env_name: str, body: str) -> ET.Element:
        if env_name == "array":
            body = re.sub(r"^\s*\{[^{}]*\}", "", body, count=1)
        rows = [r.strip() for r in re.split(r"\\\\", body) if r.strip()]
        m_el = _m("m")
        for r_str in rows:
            mr = _m("mr")
            cols = r_str.split("&")
            for c_str in cols:
                e = _m("e")
                for el in _LatexMathParser(c_str.strip()).parse_all():
                    e.append(el)
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
        return elements


def latex_to_omml_element(latex_math: str, display: bool = False):
    """Convert a LaTeX math string into a python-docx OxmlElement (<m:oMath> or <m:oMathPara>)."""
    clean = latex_math.strip()
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
            omath.append(_make_run(clean))
        else:
            for ch in children:
                omath.append(ch)

    if display:
        para = _m("oMathPara")
        para.append(omath)
        xml_str = ET.tostring(para, encoding="unicode")
    else:
        xml_str = ET.tostring(omath, encoding="unicode")
    return parse_xml(xml_str)


def latex_to_omml_xml_string(latex_math: str, display: bool = False) -> str:
    """Return raw XML string of <m:oMath> suitable for embedding inside DrawingML textboxes."""
    clean = latex_math.strip()
    if clean.startswith("$") and clean.endswith("$"):
        clean = clean[1:-1].strip()
    omath = _m("oMath")
    if clean in ("+", "-", "−", "±"):
        omath.append(_make_run("−" if clean == "-" else clean, normal_text=True))
    else:
        parser = _LatexMathParser(clean)
        children = parser.parse_all()
        if not children:
            omath.append(_make_run(clean))
        else:
            for ch in children:
                omath.append(ch)
    if display:
        para = _m("oMathPara")
        para.append(omath)
        return ET.tostring(para, encoding="unicode")
    return ET.tostring(omath, encoding="unicode")


def append_rich_latex_text_to_paragraph(
    paragraph, text: str, *, default_bold: bool = False, font_size_pt: float = 12.0
) -> None:
    """Parse mixed Vietnamese text, formatting (\\textbf, \\underline, \\textit), and $...$ math into a Word paragraph."""
    from docx.shared import Pt

    s = text.replace(r"\(", "$").replace(r"\)", "$")
    s = re.sub(r"\\texttt\{([^{}]*)\}", r"\1", s)
    s = s.replace(r"\textbackslash ", "\\").replace(r"\textbackslash", "\\")

    token_re = re.compile(
        r"(\\\[.*?\\\]|\$\$.*?\$\$|\$[^$]+\$"
        r"|\\underline\{\\textbf\{[^{}]*\}\}"
        r"|\\textbf\{\\underline\{[^{}]*\}\}"
        r"|\\textbf\{[^{}]*\}"
        r"|\\underline\{[^{}]*\}"
        r"|\\textit\{[^{}]*\}"
        r"|\\emph\{[^{}]*\})",
        re.S,
    )

    pos = 0
    for m in token_re.finditer(s):
        if m.start() > pos:
            plain = _clean_plain_tex_fragment(s[pos:m.start()])
            if plain:
                r = paragraph.add_run(plain)
                r.bold = default_bold
                r.font.size = Pt(font_size_pt)

        tok = m.group(0)
        if tok.startswith("\\[") or tok.startswith("$$"):
            omml = latex_to_omml_element(tok, display=True)
            paragraph._p.append(omml)
        elif tok.startswith("$") and tok.endswith("$"):
            omml = latex_to_omml_element(tok[1:-1], display=False)
            paragraph._p.append(omml)
        elif tok.startswith(r"\underline{\textbf{") or tok.startswith(r"\textbf{\underline{"):
            inner = tok[19:-2]
            _append_formatted_runs(paragraph, inner, bold=True, underline=True, font_size_pt=font_size_pt)
        elif tok.startswith(r"\textbf{"):
            inner = tok[8:-1]
            _append_formatted_runs(paragraph, inner, bold=True, underline=False, font_size_pt=font_size_pt)
        elif tok.startswith(r"\underline{"):
            inner = tok[11:-1]
            _append_formatted_runs(paragraph, inner, bold=default_bold, underline=True, font_size_pt=font_size_pt)
        elif tok.startswith(r"\textit{") or tok.startswith(r"\emph{"):
            inner = tok.split("{", 1)[1][:-1]
            _append_formatted_runs(paragraph, inner, bold=default_bold, italic=True, font_size_pt=font_size_pt)

        pos = m.end()

    if pos < len(s):
        plain = _clean_plain_tex_fragment(s[pos:])
        if plain:
            r = paragraph.add_run(plain)
            r.bold = default_bold
            r.font.size = Pt(font_size_pt)


def _append_formatted_runs(
    paragraph,
    text: str,
    *,
    bold: bool = False,
    underline: bool = False,
    italic: bool = False,
    font_size_pt: float = 12.0,
) -> None:
    from docx.shared import Pt

    parts = re.split(r"(\$[^$]+\$)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("$") and part.endswith("$"):
            paragraph._p.append(latex_to_omml_element(part[1:-1], display=False))
        else:
            clean = _clean_plain_tex_fragment(part)
            if clean:
                r = paragraph.add_run(clean)
                r.bold = bold
                r.underline = underline
                r.italic = italic
                r.font.size = Pt(font_size_pt)


def _clean_plain_tex_fragment(s: str) -> str:
    s = re.sub(r"\\(?:par|noindent|centering|raggedright|raggedleft|hfill|vfill)\b\s*", "", s)
    s = re.sub(r"\\(?:vspace|hspace)\*?\s*\{[^{}]*\}", "", s)
    s = re.sub(r"\\(?:setstretch|linespread)\s*\{[^{}]*\}", "", s)
    s = re.sub(r"\\(?:small|footnotesize|scriptsize|large|Large|LARGE|huge|Huge|normalsize)\b\s*", "", s)
    s = s.replace("``", "“").replace("''", "”").replace("~", " ").replace(r"\&", "&").replace(r"\%", "%")
    s = re.sub(r"[ \t]+", " ", s)
    return s
