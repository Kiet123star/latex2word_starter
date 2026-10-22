"""Smart LaTeX Preprocessor (`backend/latex_preprocessor.py`).

Lớp tiền xử lý chạy TRƯỚC Nhân Xử Lý Hợp Nhất (core_converter) — chỉ dùng cho nhánh "Dán mã LaTeX".
Không sửa đổi engine gốc: mọi cấu trúc lạ được dịch về đúng tập lệnh mà engine đã hỗ trợ
(\\textbf, \\textit, \\underline, $...$, \\[...\\], tabular, tikzpicture, \\choice, \\choiceTF, \\immini ...),
còn các định dạng engine chưa có (căn giữa, màu chữ, gạch ngang, chú thích chân trang, cột, khung,
viền bảng chuẩn LaTeX, tab, ngắt trang, header/footer...) được mã hoá bằng ký tự đánh dấu vô hình
(Unicode Private Use Area) để `backend/smart_paste.py` áp dụng lên file .docx sau khi engine chạy xong.

Các tầng xử lý:
  1. Chuẩn hoá: xuống dòng, \\input, bảo vệ verbatim, bỏ chú thích `%`.
  2. Phase A — Mở rộng macro người dùng: \\newcommand, \\renewcommand, \\providecommand, \\def (cả tham số
     có dấu phân cách), \\let, \\DeclareMathOperator, \\DeclarePairedDelimiter, \\NewDocumentCommand (xparse),
     \\newenvironment, \\NewDocumentEnvironment, \\newif/\\ifX, \\iftrue/\\iffalse, \\ifthenelse...
  3. Phase B — Dựng cấu trúc: đề thi (ex, \\cau, \\bonpa, \\choice, \\choiceTF, \\loigiai, \\shortans...),
     section/đánh số, định lý (\\newtheorem), danh sách (itemize/enumerate/description/tasks), bảng,
     công thức đánh số (equation/align/gather/multline/eqnarray...), hình/caption, label/ref/cite,
     footnote, màu, hộp, căn lề, multicols, trang tiêu đề, mục lục, header/footer...
  4. Chuẩn hoá toán: ký hiệu thiếu → Unicode, \\vv, \\overparen, \\underbrace, \\substack, \\pmod,
     \\operatorname*, \\mathcal/\\mathbb/\\mathfrak → chữ toán Unicode, \\ce (hoá học), siunitx...
  5. Fallback an toàn: lệnh lạ không ai định nghĩa → bỏ tên lệnh, giữ nội dung (không bao giờ in ra "\\abc").
"""
from __future__ import annotations

import datetime as _dt
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

# =====================================================================================
# 0. KÝ TỰ ĐÁNH DẤU (Unicode Private Use Area) — được smart_paste.py xử lý rồi xoá khỏi .docx
# =====================================================================================
MK_PARA = "\ue000"        # + idx  → thuộc tính đoạn (căn lề, thụt lề, heading, ngắt trang, khung...)
MK_SPAN = "\ue001"        # + idx  → mở vùng định dạng ký tự (màu, gạch, cỡ chữ, font...)
MK_END = "\ue002"         # đóng vùng định dạng gần nhất
MK_TAB = "\ue003"         # tab thường
MK_HFILL = "\ue004"       # \hfill
MK_DOTFILL = "\ue005"     # \dotfill
MK_RULEFILL = "\ue006"    # \hrulefill
MK_FOOTNOTE = "\ue007"    # + idx  → chú thích chân trang
MK_TABLE = "\ue008"       # + idx  → thông số viền/độ rộng bảng
MK_PAGE = "\ue009"        # trường số trang
MK_NUMPAGES = "\ue00a"    # trường tổng số trang
_REF_O = "\ue00b"         # (nội bộ) tham chiếu chờ giải
_REF_C = "\ue00c"
_VERB_BLOCK = "\ue00d"    # (nội bộ) + idx
_VERB_INLINE = "\ue00e"   # (nội bộ) + idx

LIT_CHARS: dict[str, str] = {
    "\\": "\ue010", "{": "\ue011", "}": "\ue012", "$": "\ue013", "%": "\ue014", "&": "\ue015",
    "~": "\ue016", "#": "\ue017", "^": "\ue018", "_": "\ue019", "`": "\ue01a", "'": "\ue01b",
    "-": "\ue01c", "[": "\ue01d", "]": "\ue01e",
}
LIT_REVERSE: dict[str, str] = {v: k for k, v in LIT_CHARS.items()}
IDX_BASE = 0xE100


def idx_char(i: int) -> str:
    return chr(IDX_BASE + i)


def char_idx(c: str) -> int:
    return ord(c) - IDX_BASE


def lit(s: str) -> str:
    """Escape every engine-sensitive ASCII char so the core engine prints it literally."""
    return "".join(LIT_CHARS.get(ch, ch) for ch in s)


# =====================================================================================
# 1. HÀM QUÉT CẤP THẤP
# =====================================================================================
def _is_name_char(c: str) -> bool:
    return c.isalpha() or c == "@"


def read_cmd(s: str, i: int) -> tuple[str, int]:
    """s[i] == '\\' → (name, end). Control symbols return their single char."""
    j = i + 1
    n = len(s)
    if j >= n:
        return "", j
    if _is_name_char(s[j]) and s[j].isascii():
        k = j
        while k < n and s[k].isascii() and _is_name_char(s[k]):
            k += 1
        return s[j:k], k
    return s[j], j + 1


def skip_ws(s: str, i: int, allow_par: bool = False) -> int:
    """Skip spaces/tabs and single newlines (stop before a blank line unless allow_par)."""
    n = len(s)
    while i < n:
        c = s[i]
        if c in " \t\r":
            i += 1
        elif c == "\n":
            if not allow_par:
                k = i + 1
                while k < n and s[k] in " \t\r":
                    k += 1
                if k < n and s[k] == "\n":
                    return i
            i += 1
        else:
            break
    return i


def skip_spaces_only(s: str, i: int) -> int:
    n = len(s)
    while i < n and s[i] in " \t":
        i += 1
    return i


def read_group(s: str, i: int) -> tuple[str, int]:
    """s[i] == '{' → (inner, index after matching '}')."""
    n = len(s)
    depth = 1
    j = i + 1
    while j < n:
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1 : j], j + 1
        j += 1
    return s[i + 1 :], n


def read_bracket(s: str, i: int) -> tuple[str, int]:
    """s[i] == '[' → (inner, end) — braces protect ']'."""
    n = len(s)
    depth_b = 0
    depth_sq = 1
    j = i + 1
    while j < n:
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth_b += 1
        elif c == "}":
            depth_b = max(0, depth_b - 1)
        elif depth_b == 0:
            if c == "[":
                depth_sq += 1
            elif c == "]":
                depth_sq -= 1
                if depth_sq == 0:
                    return s[i + 1 : j], j + 1
        j += 1
    return s[i + 1 :], n


def read_paren(s: str, i: int) -> tuple[str, int]:
    n = len(s)
    depth = 1
    j = i + 1
    while j < n:
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return s[i + 1 : j], j + 1
        j += 1
    return s[i + 1 :], n


def read_opt(s: str, i: int, allow_newline: bool = True) -> tuple[str | None, int]:
    j = skip_ws(s, i) if allow_newline else skip_spaces_only(s, i)
    if j < len(s) and s[j] == "[":
        return read_bracket(s, j)
    return None, i


def read_arg(s: str, i: int) -> tuple[str | None, int]:
    """Read one TeX argument: braced group, control sequence, or single char."""
    j = skip_ws(s, i)
    n = len(s)
    if j >= n:
        return None, i
    c = s[j]
    if c == "{":
        return read_group(s, j)
    if c == "\\":
        name, k = read_cmd(s, j)
        return s[j:k], k
    if c in "}]":
        return None, i
    if c == "\n":
        return None, i
    return c, j + 1


def read_braced(s: str, i: int) -> tuple[str | None, int]:
    """Read a mandatory braced argument only (no bare-token fallback)."""
    j = skip_ws(s, i)
    if j < len(s) and s[j] == "{":
        return read_group(s, j)
    return None, i


def read_star(s: str, i: int) -> tuple[bool, int]:
    j = skip_spaces_only(s, i)
    if j < len(s) and s[j] == "*":
        return True, j + 1
    return False, i


def find_env_end(s: str, i: int, env: str) -> tuple[int, int]:
    """From index i (inside the env), find matching \\end{env}. Returns (content_end, after_end)."""
    pat = re.compile(r"\\(begin|end)\s*\{" + re.escape(env) + r"\}")
    depth = 1
    pos = i
    while True:
        m = pat.search(s, pos)
        if not m:
            return len(s), len(s)
        # ignore escaped backslash (\\begin is a linebreak + "begin")
        k = m.start()
        bs = 0
        while k - 1 - bs >= 0 and s[k - 1 - bs] == "\\":
            bs += 1
        if bs % 2 == 1:
            pos = m.end()
            continue
        if m.group(1) == "begin":
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                return m.start(), m.end()
        pos = m.end()


def split_top_level(s: str, sep_cmds: tuple[str, ...]) -> list[tuple[str, str]]:
    """Split s at top-level occurrences of \\cmd (not inside groups or nested envs).
    Returns [(cmd_or_'', chunk)] where chunk is the text after the command up to the next one."""
    out: list[tuple[str, str]] = []
    depth = 0
    env_depth = 0
    i = 0
    n = len(s)
    cur_cmd = ""
    cur_start = 0
    while i < n:
        c = s[i]
        if c == "\\":
            name, j = read_cmd(s, i)
            if name in ("begin", "end"):
                k = skip_spaces_only(s, j)
                if k < n and s[k] == "{":
                    _, k2 = read_group(s, k)
                    env_depth += 1 if name == "begin" else -1
                    i = k2
                    continue
            if depth == 0 and env_depth == 0 and name in sep_cmds and (
                j >= n or not (s[j].isascii() and _is_name_char(s[j])) or not name.isalpha()
            ):
                out.append((cur_cmd, s[cur_start:i]))
                cur_cmd = name
                cur_start = j
            i = j
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth = max(0, depth - 1)
        i += 1
    out.append((cur_cmd, s[cur_start:]))
    return out


# =====================================================================================
# 2. ĐƠN VỊ ĐO, MÀU
# =====================================================================================
TEXTWIDTH_TWIPS_DEFAULT = 9638  # A4 21cm − 2×2cm


def dim_to_twips(spec: str, textwidth: int = TEXTWIDTH_TWIPS_DEFAULT, font_pt: float = 12.0) -> int | None:
    s = (spec or "").strip().replace(" ", "")
    if not s:
        return None
    m = re.match(r"^(-?\d*\.?\d*)\\(textwidth|linewidth|columnwidth|hsize|paperwidth|textheight)$", s)
    if m:
        f = float(m.group(1)) if m.group(1) not in ("", "-", ".") else 1.0
        return int(f * textwidth)
    m = re.match(r"^(-?\d*\.?\d+)(cm|mm|in|pt|bp|em|ex|pc|dd|sp)?$", s)
    if not m:
        return None
    v = float(m.group(1))
    u = m.group(2) or "pt"
    pt = {
        "cm": v * 28.3465, "mm": v * 2.83465, "in": v * 72.0, "pt": v * 0.99626, "bp": v,
        "em": v * font_pt, "ex": v * font_pt * 0.45, "pc": v * 12.0, "dd": v * 1.07, "sp": v / 65536.0,
    }[u]
    return int(pt * 20)


XCOLOR_NAMED: dict[str, str] = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "00FF00", "blue": "0000FF",
    "cyan": "00FFFF", "magenta": "FF00FF", "yellow": "FFFF00", "gray": "808080", "grey": "808080",
    "darkgray": "404040", "lightgray": "BFBFBF", "brown": "BF8040", "lime": "BFFF00", "olive": "808000",
    "orange": "FF8000", "pink": "FFBFBF", "purple": "BF0040", "teal": "008080", "violet": "800080",
    "darkblue": "00008B", "darkred": "8B0000", "darkgreen": "006400", "navy": "000080", "maroon": "800000",
    "gold": "FFD700", "silver": "C0C0C0", "indigo": "4B0082", "skyblue": "87CEEB", "turquoise": "40E0D0",
    "crimson": "DC143C", "forestgreen": "228B22", "royalblue": "4169E1", "tomato": "FF6347",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "lightyellow": "FFFFE0", "darkorange": "FF8C00",
    "Red": "ED1B23", "Blue": "2E3192", "Green": "00A64F", "Orange": "F58137", "Purple": "99479B",
    "Brown": "792500", "Maroon": "AF3235", "NavyBlue": "006EB8", "RoyalBlue": "0071BC", "ForestGreen": "009B55",
    "OliveGreen": "3C8031", "Black": "000000", "Gray": "949698", "White": "FFFFFF", "Cyan": "00AEEF",
    "Magenta": "EC008C", "Yellow": "FFF200", "Violet": "58429B", "RedViolet": "A1246B", "BlueViolet": "473992",
    "Emerald": "00A99D", "JungleGreen": "00A99A", "SeaGreen": "3FBC9D", "Turquoise": "00B4CE",
    "Aquamarine": "00B5BE", "Bittersweet": "C04F17", "BrickRed": "B6321C", "Mahogany": "A9341F",
    "Peach": "F7965A", "Salmon": "F69289", "Rhodamine": "EF559F", "Orchid": "AF72B0", "Plum": "92268F",
    "Periwinkle": "7977B8", "Cerulean": "00A2E3", "ProcessBlue": "00B0F0", "SkyBlue": "46C5DD",
    "MidnightBlue": "006795", "Dandelion": "FDBC42", "Goldenrod": "FFDF42", "Melon": "F89E7B",
    "YellowOrange": "FAA21A", "RedOrange": "F26035", "Lavender": "F49EC4", "Thistle": "D883B7",
    "CarnationPink": "F282B4", "WildStrawberry": "EE2967", "Sepia": "671800", "Tan": "DA9D76",
    "Apricot": "FBB982", "LimeGreen": "8DC73E", "YellowGreen": "98CC70", "SpringGreen": "C6DC67",
    "PineGreen": "008B72", "TealBlue": "00AEB3", "CadetBlue": "74729A", "CornflowerBlue": "41B0E4",
    "DarkOrchid": "A4538A", "Fuchsia": "8C368C", "Mulberry": "A93C93", "Lilac": "AE6ABA", "Gray0": "000000",
}


def _hex_to_rgb(h: str) -> tuple[int, int, int]:
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgb_to_hex(r: float, g: float, b: float) -> str:
    return "%02X%02X%02X" % (max(0, min(255, round(r))), max(0, min(255, round(g))), max(0, min(255, round(b))))


def _model_to_hex(model: str, spec: str) -> str | None:
    model = (model or "").strip()
    spec = (spec or "").strip()
    try:
        if model == "HTML":
            h = spec.lstrip("#")
            return h.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", h) else None
        nums = [float(x) for x in re.split(r"[,\s]+", spec) if x]
        if model == "rgb" and len(nums) == 3:
            return _rgb_to_hex(nums[0] * 255, nums[1] * 255, nums[2] * 255)
        if model == "RGB" and len(nums) == 3:
            return _rgb_to_hex(*nums)
        if model == "cmyk" and len(nums) == 4:
            c, m, y, k = nums
            return _rgb_to_hex(255 * (1 - c) * (1 - k), 255 * (1 - m) * (1 - k), 255 * (1 - y) * (1 - k))
        if model == "cmy" and len(nums) == 3:
            return _rgb_to_hex(255 * (1 - nums[0]), 255 * (1 - nums[1]), 255 * (1 - nums[2]))
        if model == "gray" and len(nums) == 1:
            return _rgb_to_hex(nums[0] * 255, nums[0] * 255, nums[0] * 255)
        if model == "Gray" and len(nums) == 1:
            v = nums[0] * 255 / 15
            return _rgb_to_hex(v, v, v)
    except ValueError:
        return None
    return None


# =====================================================================================
# 3. BẢNG KÝ HIỆU
# =====================================================================================
TEXT_SYMBOLS: dict[str, str] = {
    "textbar": "|", "textless": "<", "textgreater": ">", "textasciicircum": "^", "textunderscore": "_",
    "textendash": "–", "textemdash": "—", "textquoteleft": "‘", "textquoteright": "’",
    "textquotedblleft": "“", "textquotedblright": "”", "textquotesingle": "'", "textquotedbl": '"',
    "guillemotleft": "«", "guillemotright": "»", "guillemetleft": "«", "guillemetright": "»",
    "guilsinglleft": "‹", "guilsinglright": "›", "quotedblbase": "„", "quotesinglbase": "‚",
    "textbullet": "•", "textperiodcentered": "·", "textellipsis": "…", "dots": "…", "ldots": "…",
    "textdagger": "†", "dag": "†", "textdaggerdbl": "‡", "ddag": "‡", "textsection": "§", "S": "§",
    "textparagraph": "¶", "P": "¶", "copyright": "©", "textcopyright": "©", "textregistered": "®",
    "texttrademark": "™", "pounds": "£", "textsterling": "£", "euro": "€", "texteuro": "€", "EUR": "€",
    "yen": "¥", "textyen": "¥", "textcent": "¢", "textdegree": "°", "degree": "°", "textcelsius": "℃",
    "celsius": "°C", "textperthousand": "‰", "textonehalf": "½", "textonequarter": "¼",
    "textthreequarters": "¾", "texttimes": "×", "textdiv": "÷", "textpm": "±", "textmu": "µ",
    "textminus": "−", "textonesuperior": "¹", "texttwosuperior": "²", "textthreesuperior": "³",
    "textordfeminine": "ª", "textordmasculine": "º", "textexclamdown": "¡", "textquestiondown": "¿",
    "i": "ı", "j": "ȷ", "ss": "ß", "SS": "SS", "ae": "æ", "AE": "Æ", "oe": "œ", "OE": "Œ", "o": "ø",
    "O": "Ø", "aa": "å", "AA": "Å", "l": "ł", "L": "Ł", "dh": "ð", "DH": "Ð", "th": "þ", "TH": "Þ",
    "ng": "ŋ", "NG": "Ŋ", "dj": "đ", "DJ": "Đ", "ohorn": "ơ", "OHORN": "Ơ", "uhorn": "ư", "UHORN": "Ư",
    "abreve": "ă", "ABREVE": "Ă", "checkmark": "✓", "Checkmark": "✔", "XSolidBrush": "✘",
    "ballotx": "✗", "Square": "☐", "CheckedBox": "☑", "XBox": "☒", "smiley": "☺", "frownie": "☹",
    "Writinghand": "✍", "Pointinghand": "☞", "phone": "☎", "Letter": "✉", "LaTeX": "LaTeX",
    "TeX": "TeX", "LaTeXe": "LaTeX2ε", "XeLaTeX": "XeLaTeX", "LuaLaTeX": "LuaLaTeX", "XeTeX": "XeTeX",
    "BibTeX": "BibTeX", "AmS": "AMS", "AmSTeX": "AMS-TeX", "eTeX": "ε-TeX", "pdfLaTeX": "pdfLaTeX",
    "textvisiblespace": "␣", "textnumero": "№", "textohm": "Ω", "lq": "‘", "rq": "’",
    "textasteriskcentered": "∗", "textbigcircle": "◯", "textopenbullet": "◦", "textreferencemark": "※",
    "textlangle": "⟨", "textrangle": "⟩", "textleftarrow": "←", "textrightarrow": "→",
    "textuparrow": "↑", "textdownarrow": "↓", "slash": "/", "textslash": "/", "textfractionsolidus": "⁄",
    "nobreakspace": "\u00a0", "space": " ", "enspace": "\u2002", "enskip": "\u2002", "quad": "\u2003",
    "qquad": "\u2003\u2003", "thinspace": "\u2009", "negthinspace": "", "medspace": "\u205f",
    "thickspace": "\u2004", "nobreak": "", "allowbreak": "", "break": "", "hss": "", "relax": "",
    "protect": "", "leavevmode": "", "ignorespaces": "", "unskip": "", "null": "", "empty": "",
    "strut": "", "mathstrut": "", "indent": "", "noindent": "", "xspace": "", "sloppy": "", "fussy": "",
    "selectfont": "", "nolinebreak": "", "nopagebreak": "", "noexpand": "", "expandafter": "",
    "makeatletter": "", "makeatother": "", "frenchspacing": "", "nonfrenchspacing": "", "obeylines": "",
    "raggedbottom": "", "flushbottom": "", "maketitle@": "", "hfil": "", "vfil": "", "filbreak": "",
    "textsquare": "□", "blacksquare": "■", "qedsymbol": "□", "QED": "∎", "qed": "□", "dotfill@": "",
    "textpilcrow": "¶", "textdollar": LIT_CHARS["$"], "textbackslash": LIT_CHARS["\\"],
    "textasciitilde": LIT_CHARS["~"], "textbraceleft": LIT_CHARS["{"], "textbraceright": LIT_CHARS["}"],
    "textasciigrave": LIT_CHARS["`"], "textquotestraightbase": "‚", "textcircledP": "℗",
    "textinterrobang": "‽", "textestimated": "℮", "textmho": "℧", "textrecipe": "℞", "textdied": "†",
    "textborn": "*", "textmarried": "⚭", "textdivorced": "⚮", "textleaf": "❦", "textmusicalnote": "♪",
}

TEXT_ACCENTS: dict[str, str] = {
    "'": "\u0301", "`": "\u0300", "^": "\u0302", '"': "\u0308", "~": "\u0303", "=": "\u0304",
    ".": "\u0307", "u": "\u0306", "v": "\u030c", "H": "\u030b", "c": "\u0327", "d": "\u0323",
    "b": "\u0331", "r": "\u030a", "k": "\u0328", "t": "\u0361", "h": "\u0309", "G": "\u030f",
    "U": "\u030e", "textcommabelow": "\u0326", "textdotbelow": "\u0323", "texthookabove": "\u0309",
}

# Ký hiệu toán mà parser OMML của engine CHƯA có → Unicode (giữ nguyên các ký hiệu engine đã hỗ trợ).
MATH_SYMBOLS_EXTRA: dict[str, str] = {
    "lvert": "|", "rvert": "|", "vert": "|", "lVert": "‖", "rVert": "‖", "langle": "⟨", "rangle": "⟩",
    "lceil": "⌈", "rceil": "⌉", "lfloor": "⌊", "rfloor": "⌋", "lbrace": "\\{", "rbrace": "\\}",
    "lbrack": "[", "rbrack": "]", "ulcorner": "⌜", "urcorner": "⌝", "llcorner": "⌞", "lrcorner": "⌟",
    "lgroup": "⟮", "rgroup": "⟯", "llbracket": "⟦", "rrbracket": "⟧", "backslash": "∖",
    "leqslant": "⩽", "geqslant": "⩾", "leqq": "≦", "geqq": "≧", "lneq": "⪇", "gneq": "⪈", "lneqq": "≨",
    "gneqq": "≩", "nleq": "≰", "ngeq": "≱", "nless": "≮", "ngtr": "≯", "nleqslant": "≰", "ngeqslant": "≱",
    "ll": "≪", "gg": "≫", "lll": "⋘", "ggg": "⋙", "lesssim": "≲", "gtrsim": "≳", "lessapprox": "⪅",
    "gtrapprox": "⪆", "lessgtr": "≶", "gtrless": "≷", "prec": "≺", "succ": "≻", "preceq": "⪯",
    "succeq": "⪰", "preccurlyeq": "≼", "succcurlyeq": "≽", "nprec": "⊀", "nsucc": "⊁", "nsim": "≁",
    "ncong": "≇", "napprox": "≉", "doteq": "≐", "doteqdot": "≑", "Doteq": "≑", "fallingdotseq": "≒",
    "risingdotseq": "≓", "triangleq": "≜", "eqdef": "≝", "coloneqq": "≔", "coloneq": "≔", "eqqcolon": "≕",
    "Coloneqq": "⩴", "approxeq": "≊", "backsim": "∽", "backsimeq": "⋍", "bumpeq": "≏", "Bumpeq": "≎",
    "asymp": "≍", "circeq": "≗", "eqcirc": "≖", "thicksim": "∼", "thickapprox": "≈", "models": "⊨",
    "vdash": "⊢", "dashv": "⊣", "vDash": "⊨", "Vdash": "⊩", "Vvdash": "⊪", "nvdash": "⊬", "nvDash": "⊭",
    "top": "⊤", "bot": "⊥", "nmid": "∤", "nparallel": "∦", "shortmid": "∣", "shortparallel": "∥",
    "nshortmid": "∤", "nshortparallel": "∦", "smile": "⌣", "frown": "⌢", "bowtie": "⋈", "Join": "⨝",
    "between": "≬", "pitchfork": "⋔", "varpropto": "∝", "subsetneq": "⊊", "supsetneq": "⊋",
    "subsetneqq": "⫋", "supsetneqq": "⫌", "nsubseteq": "⊈", "nsupseteq": "⊉", "nsubset": "⊄",
    "nsupset": "⊅", "sqsubset": "⊏", "sqsupset": "⊐", "sqsubseteq": "⊑", "sqsupseteq": "⊒",
    "Subset": "⋐", "Supset": "⋑", "subseteqq": "⫅", "supseteqq": "⫆", "varsubsetneq": "⊊",
    "varsupsetneq": "⊋", "owns": "∋", "notni": "∌", "uplus": "⊎", "sqcup": "⊔", "sqcap": "⊓",
    "amalg": "⨿", "oplus": "⊕", "ominus": "⊖", "otimes": "⊗", "oslash": "⊘", "odot": "⊙",
    "circledcirc": "⊚", "circledast": "⊛", "circleddash": "⊝", "boxplus": "⊞", "boxminus": "⊟",
    "boxtimes": "⊠", "boxdot": "⊡", "bullet": "∙", "star": "⋆", "ast": "∗", "diamond": "⋄",
    "Diamond": "◇", "lozenge": "◊", "blacklozenge": "⧫", "dagger": "†", "ddagger": "‡", "wr": "≀",
    "bigcirc": "◯", "cdotp": "⋅", "ldotp": ".", "centerdot": "⋅", "intercal": "⊺", "dotplus": "∔",
    "ltimes": "⋉", "rtimes": "⋊", "leftthreetimes": "⋋", "rightthreetimes": "⋌", "curlywedge": "⋏",
    "curlyvee": "⋎", "barwedge": "⌅", "veebar": "⊻", "doublebarwedge": "⩞", "divideontimes": "⋇",
    "smallsetminus": "∖", "triangleleft": "◁", "triangleright": "▷", "bigtriangleup": "△",
    "bigtriangledown": "▽", "vartriangle": "△", "triangledown": "▽", "blacktriangle": "▲",
    "blacktriangledown": "▼", "blacktriangleleft": "◀", "blacktriangleright": "▶",
    "vartriangleleft": "⊲", "vartriangleright": "⊳", "trianglelefteq": "⊴", "trianglerighteq": "⊵",
    "lhd": "⊲", "rhd": "⊳", "unlhd": "⊴", "unrhd": "⊵", "square": "□", "Box": "□", "blacksquare": "■",
    "measuredangle": "∡", "sphericalangle": "∢", "varangle": "∡", "rightangle": "∟", "diameter": "⌀",
    "complement": "∁", "aleph": "ℵ", "beth": "ℶ", "gimel": "ℷ", "daleth": "ℸ", "wp": "℘",
    "imath": "ı", "jmath": "ȷ", "eth": "ð", "hslash": "ℏ", "mho": "℧", "Finv": "Ⅎ", "Game": "⅁",
    "Bbbk": "𝕜", "surd": "√", "flat": "♭", "natural": "♮", "sharp": "♯", "clubsuit": "♣",
    "diamondsuit": "♢", "heartsuit": "♡", "spadesuit": "♠", "checkmark": "✓", "maltese": "✠",
    "therefore": "∴", "because": "∵", "qed": "∎", "Longrightarrow": "⟹", "longrightarrow": "⟶",
    "Longleftarrow": "⟸", "longleftarrow": "⟵", "Longleftrightarrow": "⟺", "longleftrightarrow": "⟷",
    "longmapsto": "⟼", "impliedby": "⟸", "gets": "←", "Uparrow": "⇑", "Downarrow": "⇓",
    "Updownarrow": "⇕", "updownarrow": "↕", "nearrow": "↗", "searrow": "↘", "swarrow": "↙",
    "nwarrow": "↖", "hookrightarrow": "↪", "hookleftarrow": "↩", "leftharpoonup": "↼",
    "leftharpoondown": "↽", "rightharpoonup": "⇀", "rightharpoondown": "⇁", "rightleftharpoons": "⇌",
    "leftrightharpoons": "⇋", "upharpoonleft": "↿", "upharpoonright": "↾", "downharpoonleft": "⇃",
    "downharpoonright": "⇂", "leadsto": "⇝", "rightsquigarrow": "⇝", "leftrightsquigarrow": "↭",
    "twoheadrightarrow": "↠", "twoheadleftarrow": "↞", "rightarrowtail": "↣", "leftarrowtail": "↢",
    "looparrowright": "↬", "looparrowleft": "↫", "curvearrowright": "↷", "curvearrowleft": "↶",
    "circlearrowright": "↻", "circlearrowleft": "↺", "Rrightarrow": "⇛", "Lleftarrow": "⇚",
    "rightrightarrows": "⇉", "leftleftarrows": "⇇", "leftrightarrows": "⇆", "rightleftarrows": "⇄",
    "upuparrows": "⇈", "downdownarrows": "⇊", "Lsh": "↰", "Rsh": "↱", "multimap": "⊸",
    "nrightarrow": "↛", "nleftarrow": "↚", "nRightarrow": "⇏", "nLeftarrow": "⇍", "nleftrightarrow": "↮",
    "nLeftrightarrow": "⇎", "dashrightarrow": "⇢", "dashleftarrow": "⇠", "mapsfrom": "↤", "Mapsto": "⤇",
    "varpi": "ϖ", "varrho": "ϱ", "varsigma": "ς", "varkappa": "ϰ", "digamma": "ϝ", "omicron": "ο",
    "Alpha": "Α", "Beta": "Β", "Epsilon": "Ε", "Zeta": "Ζ", "Eta": "Η", "Iota": "Ι", "Kappa": "Κ",
    "Mu": "Μ", "Nu": "Ν", "Omicron": "Ο", "Rho": "Ρ", "Tau": "Τ", "Chi": "Χ", "varGamma": "𝛤",
    "varDelta": "𝛥", "varTheta": "𝛩", "varLambda": "𝛬", "varXi": "𝛯", "varPi": "𝛱", "varSigma": "𝛴",
    "varUpsilon": "𝛶", "varPhi": "𝛷", "varPsi": "𝛹", "varOmega": "𝛺", "upalpha": "α", "upbeta": "β",
    "upgamma": "γ", "updelta": "δ", "uppi": "π", "upmu": "μ", "upepsilon": "ε", "uptheta": "θ",
    "uplambda": "λ", "upsigma": "σ", "upomega": "ω", "upphi": "φ", "coprod": "∐", "bigoplus": "⨁",
    "bigotimes": "⨂", "bigodot": "⨀", "biguplus": "⨄", "bigsqcup": "⨆", "bigvee": "⋁", "bigwedge": "⋀",
    "iiiint": "⨌", "oiint": "∯", "oiiint": "∰", "intop": "∫", "smallint": "∫", "dotsc": "…",
    "dotsb": "⋯", "dotsm": "⋯", "dotsi": "⋯", "dotso": "…", "hdots": "…", "iddots": "⋰", "adots": "⋰",
    "vcentcolon": ":", "dblcolon": "∷", "Colon": "∷", "backprime": "‵", "dprime": "″", "trprime": "‴",
    "diagup": "╱", "diagdown": "╲", "mathellipsis": "…", "lightning": "↯", "nexists": "∄",
    "varnothing": "∅", "emptyset": "∅", "Box": "□", "AA": "Å", "lnot": "¬", "neg": "¬",
    "textdegree": "°", "dots": "…", "colon": ":", "ell": "ℓ", "hbar": "ℏ", "infty": "∞",
    "mathsection": "§", "mathparagraph": "¶", "S": "§", "P": "¶", "pounds": "£", "copyright": "©",
    "euro": "€", "yen": "¥", "dag": "†", "ddag": "‡", "Diamondblack": "◆", "bigstar": "★",
    "blackstar": "★", "circledS": "Ⓢ", "circledR": "®", "nVdash": "⊮", "nVDash": "⊯",
    "leftmoon": "☾", "rightmoon": "☽", "earth": "♁", "male": "♂", "female": "♀", "sun": "☼",
    "Square": "□", "openbox": "□", "notin": "∉", "nni": "∌",
}

# Ký hiệu (cả những cái engine biết) — dùng bên trong \text{}, \mathbf{}, \mathrm{}... (engine in thô)
MATH_SYMBOLS_ALL: dict[str, str] = {
    "infty": "∞", "pm": "±", "mp": "∓", "times": "×", "cdot": "⋅", "div": "÷", "le": "≤", "leq": "≤",
    "ge": "≥", "geq": "≥", "neq": "≠", "ne": "≠", "approx": "≈", "equiv": "≡", "sim": "∼", "simeq": "≃",
    "cong": "≅", "propto": "∝", "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "supseteq": "⊇", "cup": "∪", "cap": "∩", "setminus": "∖", "emptyset": "∅",
    "varnothing": "∅", "forall": "∀", "exists": "∃", "nexists": "∄", "lor": "∨", "vee": "∨",
    "land": "∧", "wedge": "∧", "neg": "¬", "lnot": "¬", "partial": "∂", "nabla": "∇", "to": "→",
    "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "Leftrightarrow": "⇔", "iff": "⇔", "implies": "⇒", "mapsto": "↦", "uparrow": "↑",
    "downarrow": "↓", "mid": "∣", "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ϵ", "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ",
    "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "ϕ", "varphi": "φ", "chi": "χ", "psi": "ψ",
    "omega": "ω", "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π",
    "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω", "ell": "ℓ", "hbar": "ℏ",
    "angle": "∠", "perp": "⊥", "parallel": "∥", "triangle": "△", "circ": "∘", "degree": "°",
    "dots": "…", "cdots": "⋯", "ldots": "…", "vdots": "⋮", "ddots": "⋱", "prime": "′", "Vert": "‖",
    "sum": "∑", "prod": "∏", "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮", "bigcup": "⋃",
    "bigcap": "⋂", **MATH_SYMBOLS_EXTRA,
}

MATH_OPERATOR_NAMES = {
    "sec", "csc", "cosec", "coth", "sech", "csch", "arccot", "arcsec", "arccsc", "arcsinh", "arccosh",
    "arctanh", "arg", "hom", "Pr", "injlim", "projlim", "tg", "cotg", "ctg", "arctg", "arccotg",
    "arcctg", "sh", "ch", "th", "cth", "lb", "sgn", "sign", "diag", "Tr", "ord", "mes", "card", "Var",
    "Cov", "erf", "grad", "rot", "curl", "Ker", "Hom", "End", "Aut", "Log", "Arg", "Res", "supp",
    "esssup", "argmax", "argmin", "trace", "Id", "id", "codim", "coker", "Spec", "Gal", "lcm", "tr",
    "rank", "rk", "Span", "span",
}
MATH_OPERATOR_SPECIAL = {"limsup": "lim sup", "liminf": "lim inf", "varlimsup": "lim sup", "varliminf": "lim inf"}

# Lệnh engine/omml_math đã tự xử lý trong toán — KHÔNG đụng vào.
ENGINE_MATH_CMDS = {
    "frac", "dfrac", "tfrac", "cfrac", "binom", "dbinom", "tbinom", "xrightarrow", "xleftarrow",
    "overset", "underset", "sqrt", "overline", "underline", "mathbb", "text", "textrm", "textbf",
    "textit", "mathrm", "mathbf", "mathcal", "operatorname", "vec", "overrightarrow", "hat",
    "widehat", "bar", "tilde", "widetilde", "dot", "ddot", "int", "iint", "iiint", "oint", "sum",
    "prod", "bigcup", "bigcap", "lim", "max", "min", "sup", "inf", "sin", "cos", "tan", "cot", "arcsin",
    "arccos", "arctan", "sinh", "cosh", "tanh", "ln", "log", "lg", "exp", "det", "gcd", "lcm", "deg",
    "dim", "ker", "span", "Span", "Im", "Re", "rank", "rk", "tr", "heva", "hoac", "begin", "end",
    "left", "right", "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle", "quad", "qquad",
    "tag", "choose", "{", "}", "|", ",", ";", "!", " ", "%", "#", "&", "_", "$", "\\",
}

ENGINE_MATH_SYMBOL_NAMES = {
    "infty", "pm", "mp", "times", "cdot", "div", "le", "leq", "ge", "geq", "neq", "ne", "approx", "equiv",
    "sim", "simeq", "cong", "propto", "in", "notin", "ni", "subset", "subseteq", "supset", "supseteq",
    "cup", "cap", "setminus", "emptyset", "varnothing", "forall", "exists", "nexists", "lor", "vee",
    "land", "wedge", "neg", "lnot", "partial", "nabla", "to", "rightarrow", "leftarrow", "leftrightarrow",
    "Rightarrow", "Leftarrow", "Leftrightarrow", "iff", "implies", "mapsto", "uparrow", "downarrow",
    "mid", "colon", "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta", "theta",
    "vartheta", "iota", "kappa", "lambda", "mu", "nu", "xi", "pi", "rho", "sigma", "tau", "upsilon",
    "phi", "varphi", "chi", "psi", "omega", "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma",
    "Upsilon", "Phi", "Psi", "Omega", "ell", "hbar", "angle", "perp", "parallel", "triangle", "circ",
    "degree", "dots", "cdots", "ldots", "vdots", "ddots", "prime", "Vert",
}

SCRIPT_UPPER = "𝒜ℬ𝒞𝒟ℰℱ𝒢ℋℐ𝒥𝒦ℒℳ𝒩𝒪𝒫𝒬ℛ𝒮𝒯𝒰𝒱𝒲𝒳𝒴𝒵"
SCRIPT_LOWER = "𝒶𝒷𝒸𝒹ℯ𝒻ℊ𝒽𝒾𝒿𝓀𝓁𝓂𝓃ℴ𝓅𝓆𝓇𝓈𝓉𝓊𝓋𝓌𝓍𝓎𝓏"
FRAK_UPPER = "𝔄𝔅ℭ𝔇𝔈𝔉𝔊ℌℑ𝔍𝔎𝔏𝔐𝔑𝔒𝔓𝔔ℜ𝔖𝔗𝔘𝔙𝔚𝔛𝔜ℨ"
FRAK_LOWER = "𝔞𝔟𝔠𝔡𝔢𝔣𝔤𝔥𝔦𝔧𝔨𝔩𝔪𝔫𝔬𝔭𝔮𝔯𝔰𝔱𝔲𝔳𝔴𝔵𝔶𝔷"
BB_UPPER = "𝔸𝔹ℂ𝔻𝔼𝔽𝔾ℍ𝕀𝕁𝕂𝕃𝕄ℕ𝕆ℙℚℝ𝕊𝕋𝕌𝕍𝕎𝕏𝕐ℤ"
BB_LOWER = "𝕒𝕓𝕔𝕕𝕖𝕗𝕘𝕙𝕚𝕛𝕜𝕝𝕞𝕟𝕠𝕡𝕢𝕣𝕤𝕥𝕦𝕧𝕨𝕩𝕪𝕫"
BB_DIGITS = "𝟘𝟙𝟚𝟛𝟜𝟝𝟞𝟟𝟠𝟡"


def _map_alpha(text: str, upper: str, lower: str, digits: str | None = None) -> str:
    out = []
    for ch in text:
        if "A" <= ch <= "Z":
            out.append(upper[ord(ch) - 65])
        elif "a" <= ch <= "z":
            out.append(lower[ord(ch) - 97])
        elif digits and "0" <= ch <= "9":
            out.append(digits[ord(ch) - 48])
        else:
            out.append(ch)
    return "".join(out)


NOT_MAP = {
    "in": "∉", "=": "≠", "equiv": "≢", "subset": "⊄", "supset": "⊅", "subseteq": "⊈", "supseteq": "⊉",
    "parallel": "∦", "mid": "∤", "<": "≮", ">": "≯", "le": "≰", "leq": "≰", "ge": "≱", "geq": "≱",
    "sim": "≁", "approx": "≉", "cong": "≇", "exists": "∄", "ni": "∌", "simeq": "≄", "prec": "⊀",
    "succ": "⊁", "vdash": "⊬", "models": "⊭", "perp": "⊥\u0338", "leqslant": "≰", "geqslant": "≱",
}

CIRCLED = "⓪①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
DING_MAP = {"51": "✓", "52": "✔", "55": "✗", "56": "✘", "43": "☛", "42": "☛", "45": "✍", "73": "★",
            "72": "★", "108": "●", "110": "■", "111": "❏", "113": "❑", "117": "◆", "118": "❖",
            "164": "❤", "171": "①", "172": "②", "173": "③", "174": "④", "175": "⑤", "182": "①",
            "183": "②", "184": "③", "185": "④", "186": "⑤", "192": "①", "202": "➊", "212": "➔",
            "217": "➙", "220": "➜", "224": "➠", "228": "➤", "232": "➨", "233": "➩", "36": "✄",
            "33": "✁", "34": "✂", "46": "☞", "47": "✏", "48": "✐", "41": "✉", "37": "☎", "40": "✈"}

SI_UNITS = {
    "meter": "m", "metre": "m", "second": "s", "kilogram": "kg", "gram": "g", "newton": "N", "joule": "J",
    "watt": "W", "volt": "V", "ampere": "A", "ohm": "Ω", "hertz": "Hz", "pascal": "Pa", "kelvin": "K",
    "celsius": "°C", "degreeCelsius": "°C", "mole": "mol", "litre": "L", "liter": "L", "coulomb": "C",
    "farad": "F", "henry": "H", "tesla": "T", "weber": "Wb", "siemens": "S", "lumen": "lm", "lux": "lx",
    "candela": "cd", "becquerel": "Bq", "gray": "Gy", "sievert": "Sv", "radian": "rad",
    "steradian": "sr", "degree": "°", "minute": "min", "hour": "h", "day": "d", "electronvolt": "eV",
    "bar": "bar", "percent": "%", "angstrom": "Å", "atm": "atm", "calorie": "cal", "dalton": "Da",
    "km": "km", "cm": "cm", "mm": "mm", "kg": "kg", "mg": "mg", "ms": "ms", "kHz": "kHz", "MHz": "MHz",
    "kV": "kV", "mA": "mA", "kW": "kW", "kJ": "kJ", "mL": "mL", "ml": "mL", "um": "µm", "nm": "nm",
    "centi": "c", "milli": "m", "micro": "µ", "nano": "n", "pico": "p", "kilo": "k", "mega": "M",
    "giga": "G", "tera": "T", "deci": "d", "deca": "da", "hecto": "h", "femto": "f",
}


def _today_str(lang: str) -> str:
    d = _dt.date.today()
    if lang == "vi":
        return f"Ngày {d.day} tháng {d.month} năm {d.year}"
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
              "October", "November", "December"]
    return f"{months[d.month - 1]} {d.day}, {d.year}"


def _roman(n: int) -> str:
    if n <= 0:
        return str(n)
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"),
            (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = []
    for v, r in vals:
        while n >= v:
            out.append(r)
            n -= v
    return "".join(out)


def _alph(n: int) -> str:
    if n <= 0:
        return str(n)
    out = ""
    while n > 0:
        n -= 1
        out = chr(97 + n % 26) + out
        n //= 26
    return out


def _fnsymbol(n: int) -> str:
    syms = ["*", "†", "‡", "§", "¶", "‖", "**", "††", "‡‡"]
    return syms[(n - 1) % len(syms)] if n > 0 else str(n)


NUM_FORMATS: dict[str, Callable[[int], str]] = {
    "arabic": str, "alph": _alph, "Alph": lambda n: _alph(n).upper(), "roman": _roman,
    "Roman": lambda n: _roman(n).upper(), "fnsymbol": _fnsymbol,
}


# =====================================================================================
# 4. CẤU HÌNH & KẾT QUẢ
# =====================================================================================
@dataclass
class PreprocessResult:
    text: str
    para_specs: list[dict[str, Any]] = field(default_factory=list)
    span_specs: list[dict[str, Any]] = field(default_factory=list)
    table_specs: list[dict[str, Any]] = field(default_factory=list)
    footnotes: list[str] = field(default_factory=list)
    page: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    unknown_commands: dict[str, int] = field(default_factory=dict)
    unknown_environments: dict[str, int] = field(default_factory=dict)
    stats: dict[str, int] = field(default_factory=dict)

    def summary(self) -> dict[str, Any]:
        return {
            "stats": dict(self.stats),
            "warnings": list(self.warnings)[:200],
            "unknown_commands_auto_handled": dict(sorted(self.unknown_commands.items(), key=lambda x: -x[1])[:100]),
            "unknown_environments_auto_handled": dict(self.unknown_environments),
            "paragraph_styles": len(self.para_specs),
            "run_styles": len(self.span_specs),
            "styled_tables": len(self.table_specs),
            "footnotes": len(self.footnotes),
        }


@dataclass
class _Macro:
    name: str
    nargs: int = 0
    default: str | None = None
    body: str = ""
    kind: str = "new"  # new | def | xparse | paired | operator
    params: list[str] | None = None  # \def parameter text pieces: [prefix, d1, d2, ...]
    spec: list[tuple[str, str | None]] | None = None  # xparse
    soft: bool = False


@dataclass
class _EnvDef:
    name: str
    nargs: int = 0
    default: str | None = None
    begin: str = ""
    end: str = ""
    spec: list[tuple[str, str | None]] | None = None


# Lệnh cấu trúc mà engine/tiền xử lý dựng bảng chuẩn — không cho định nghĩa người dùng ghi đè.
PROTECTED_MACROS = {
    "choice", "choiceTF", "choiceTFt", "True", "immini", "impicinpar", "heva", "hoac", "boncot",
    "haicot", "motcot", "bonpa", "choiceA", "choiceB", "choiceC", "choiceD",
}
CORE_PROTECTED = {
    "section", "subsection", "subsubsection", "chapter", "part", "paragraph", "subparagraph", "item",
    "caption", "label", "ref", "eqref", "begin", "end", "textbf", "textit", "emph", "underline",
    "footnote", "includegraphics", "maketitle", "tableofcontents", "newpage", "frac", "sqrt", "left",
    "right", "par", "\\", "hline", "cline", "multicolumn", "multirow", "item", "documentclass",
    "usepackage", "newcommand", "renewcommand", "def", "let", "text", "mathrm", "mathbf",
}
# Định nghĩa để lại cho Phase B (đánh số, nhãn, tên mục...)
PASS_TO_B_RE = re.compile(
    r"^(the[a-zA-Z]+|label(enum|item)[iv]+|[a-z]*name|arraystretch|baselinestretch|theenumi|tabcolsep|"
    r"headrulewidth|footrulewidth|labelitemi|labelitemii|labelitemiii|labelitemiv)$"
)

SOFT_MACROS: dict[str, tuple[int, str]] = {
    "R": (0, r"\mathbb{R}"), "N": (0, r"\mathbb{N}"), "Z": (0, r"\mathbb{Z}"), "Q": (0, r"\mathbb{Q}"),
    "C": (0, r"\mathbb{C}"), "RR": (0, r"\mathbb{R}"), "NN": (0, r"\mathbb{N}"), "ZZ": (0, r"\mathbb{Z}"),
    "QQ": (0, r"\mathbb{Q}"), "CC": (0, r"\mathbb{C}"), "abs": (1, r"\left|#1\right|"),
    "norm": (1, r"\left\|#1\right\|"), "set": (1, r"\left\{#1\right\}"),
    "floor": (1, r"\left\lfloor #1\right\rfloor"), "ceil": (1, r"\left\lceil #1\right\rceil"),
    "vv": (1, r"\overrightarrow{#1}"), "overarc": (1, r"\wideparen{#1}"), "widearc": (1, r"\wideparen{#1}"),
    "arc": (1, r"\wideparen{#1}"), "overparen": (1, r"\wideparen{#1}"), "dx": (0, r"\,\mathrm{d}x"),
    "dy": (0, r"\,\mathrm{d}y"), "dt": (0, r"\,\mathrm{d}t"), "du": (0, r"\,\mathrm{d}u"),
    "ra": (0, r"\rightarrow"), "Ra": (0, r"\Rightarrow"), "LRa": (0, r"\Leftrightarrow"),
    "LRA": (0, r"\Leftrightarrow"), "RA": (0, r"\Rightarrow"), "eps": (0, r"\varepsilon"),
    "tron": (1, r"\overset{\frown}{#1}"), "khung": (1, r"\fbox{#1}"), "hide": (1, r""),
    "dangtoan": (1, r"\textbf{#1}"), "dapan": (1, r"\textbf{Đáp án:} #1"),
    "TT": (0, r"\text{TT}"), "SS": (0, r"\text{SS}"), "uuu": (0, r""),
}

CONDITIONAL_PRIMITIVES = {
    "if", "ifx", "ifnum", "ifdim", "ifodd", "ifcase", "ifvmode", "ifhmode", "ifmmode", "ifinner",
    "ifcat", "ifeof", "iftrue", "iffalse", "ifdefined", "ifcsname", "ifvoid", "ifhbox", "ifvbox",
}


# =====================================================================================
# 5. PHASE A — MỞ RỘNG MACRO NGƯỜI DÙNG
# =====================================================================================
class _MacroExpander:
    MAX_EXPANSIONS = 40000
    MAX_PER_MACRO = 6000

    def __init__(self, res: PreprocessResult) -> None:
        self.res = res
        self.macros: dict[str, _Macro] = {}
        self.envs: dict[str, _EnvDef] = {}
        self.bools: dict[str, bool] = {}
        self.coltypes: dict[str, tuple[int, str]] = {}
        self.count = 0
        self.per_macro: dict[str, int] = {}
        self.disabled: set[str] = set()
        for name, (n, body) in SOFT_MACROS.items():
            self.macros[name] = _Macro(name=name, nargs=n, body=body, soft=True)

    # ---------------- definitions ----------------
    def _read_def_name(self, s: str, i: int) -> tuple[str | None, int]:
        j = skip_ws(s, i)
        if j < len(s) and s[j] == "{":
            inner, k = read_group(s, j)
            inner = inner.strip()
            if inner.startswith("\\"):
                return read_cmd(inner, 0)[0], k
            return inner or None, k
        if j < len(s) and s[j] == "\\":
            name, k = read_cmd(s, j)
            return name, k
        return None, i

    def _register(self, m: _Macro, *, provide: bool = False) -> None:
        if m.name in PROTECTED_MACROS:
            self.res.warnings.append(f"Giữ cấu trúc chuẩn cho \\{m.name} (bỏ qua định nghĩa lại của người dùng).")
            return
        if m.name in CORE_PROTECTED:
            return
        if "\\@startsection" in m.body or "\\@" in m.body[:2]:
            return
        if provide and m.name in self.macros and not self.macros[m.name].soft:
            return
        self.macros[m.name] = m
        self.res.stats["macros_defined"] = self.res.stats.get("macros_defined", 0) + 1

    def _parse_newcommand(self, s: str, i: int, j: int, provide: bool) -> int | None:
        star, k = read_star(s, j)
        name, k = self._read_def_name(s, k)
        if not name:
            return None
        nargs_s, k = read_opt(s, k)
        nargs = int(nargs_s.strip()) if nargs_s and nargs_s.strip().isdigit() else 0
        default, k2 = (None, k)
        if nargs > 0:
            default, k2 = read_opt(s, k)
        body, k3 = read_arg(s, k2)
        if body is None:
            return None
        if PASS_TO_B_RE.match(name):
            return None  # để Phase B xử lý
        self._register(_Macro(name=name, nargs=nargs, default=default, body=body), provide=provide)
        return k3

    def _parse_def(self, s: str, i: int, j: int) -> int | None:
        k = skip_ws(s, j)
        if k >= len(s) or s[k] != "\\":
            return None
        name, k = read_cmd(s, k)
        # parameter text up to first '{'
        p_start = k
        n = len(s)
        while k < n and s[k] != "{":
            if s[k] == "\\":
                k = read_cmd(s, k)[1]
                continue
            k += 1
        if k >= n:
            return None
        ptext = s[p_start:k]
        body, k2 = read_group(s, k)
        if PASS_TO_B_RE.match(name):
            return None
        pieces = re.split(r"#[1-9]", ptext)
        nargs = len(pieces) - 1
        params = [p for p in pieces]
        simple = all(p.strip() == "" for p in pieces)
        m = _Macro(name=name, nargs=nargs, body=body, kind="new" if simple else "def",
                   params=None if simple else params)
        self._register(m)
        return k2

    def _parse_let(self, s: str, i: int, j: int) -> int | None:
        k = skip_ws(s, j)
        if k >= len(s) or s[k] != "\\":
            return None
        name, k = read_cmd(s, k)
        k = skip_ws(s, k)
        if k < len(s) and s[k] == "=":
            k = skip_ws(s, k + 1)
        if k >= len(s):
            return None
        if s[k] == "\\":
            target, k2 = read_cmd(s, k)
            if target in self.macros:
                src = self.macros[target]
                self._register(_Macro(name=name, nargs=src.nargs, default=src.default, body=src.body,
                                      kind=src.kind, params=src.params, spec=src.spec))
            elif target != name:
                self._register(_Macro(name=name, body="\\" + target + (" " if target.isalpha() else "")))
            return k2
        self._register(_Macro(name=name, body=s[k]))
        return k + 1

    def _parse_xparse_spec(self, spec: str) -> list[tuple[str, str | None]]:
        out: list[tuple[str, str | None]] = []
        i = 0
        n = len(spec)
        while i < n:
            c = spec[i]
            if c in " \t\n+!>":
                i += 1
                continue
            if c in "mosvlu":
                out.append((c, None))
                i += 1
            elif c in "OD" or c in "rRdDgGtTe":
                if c == "O":
                    j = skip_ws(spec, i + 1)
                    if j < n and spec[j] == "{":
                        d, j2 = read_group(spec, j)
                        out.append(("O", d))
                        i = j2
                    else:
                        out.append(("o", None))
                        i += 1
                elif c in "tT":
                    out.append(("s", None))
                    i += 2
                else:
                    out.append(("m", None))
                    i += 1
                    while i < n and spec[i] not in " mosOvlrRdDgGtTe+":
                        if spec[i] == "{":
                            i = read_group(spec, i)[1]
                        else:
                            i += 1
            else:
                i += 1
        return out

    def _parse_newdocumentcommand(self, s: str, i: int, j: int, provide: bool) -> int | None:
        name, k = self._read_def_name(s, j)
        if not name:
            return None
        spec_s, k = read_braced(s, k)
        body, k2 = read_braced(s, k)
        if spec_s is None or body is None:
            return None
        spec = self._parse_xparse_spec(spec_s)
        self._register(_Macro(name=name, nargs=len(spec), body=body, kind="xparse", spec=spec), provide=provide)
        return k2

    def _parse_newenvironment(self, s: str, i: int, j: int, provide: bool) -> int | None:
        star, k = read_star(s, j)
        name, k = read_braced(s, k)
        if not name:
            return None
        nargs_s, k = read_opt(s, k)
        nargs = int(nargs_s.strip()) if nargs_s and nargs_s.strip().isdigit() else 0
        default = None
        if nargs > 0:
            default, k = read_opt(s, k)
        b, k = read_braced(s, k)
        e, k2 = read_braced(s, k)
        if b is None or e is None:
            return None
        name = name.strip()
        if provide and name in self.envs:
            return k2
        self.envs[name] = _EnvDef(name=name, nargs=nargs, default=default, begin=b, end=e)
        self.res.stats["environments_defined"] = self.res.stats.get("environments_defined", 0) + 1
        return k2

    def _parse_newdocumentenvironment(self, s: str, i: int, j: int) -> int | None:
        name, k = read_braced(s, j)
        spec_s, k = read_braced(s, k)
        b, k = read_braced(s, k)
        e, k2 = read_braced(s, k)
        if not name or spec_s is None or b is None or e is None:
            return None
        spec = self._parse_xparse_spec(spec_s)
        self.envs[name.strip()] = _EnvDef(name=name.strip(), nargs=len(spec), begin=b, end=e, spec=spec)
        return k2

    def _parse_mathop(self, s: str, i: int, j: int) -> int | None:
        star, k = read_star(s, j)
        name, k = self._read_def_name(s, k)
        text, k2 = read_braced(s, k)
        if not name or text is None:
            return None
        self._register(_Macro(name=name, body=r"\operatorname{" + text + "}"))
        return k2

    def _parse_paired(self, s: str, i: int, j: int, x: bool) -> int | None:
        name, k = self._read_def_name(s, j)
        if not name:
            return None
        nargs = 1
        if x:
            ns, k = read_opt(s, k)
            nargs = int(ns) if ns and ns.strip().isdigit() else 1
        l, k = read_arg(s, k)
        r, k2 = read_arg(s, k)
        body = ""
        if x:
            body, k2 = read_arg(s, k2)
        if l is None or r is None:
            return None
        self._register(_Macro(name=name, nargs=nargs, body=body or "#1", kind="paired", params=[l, r]))
        return k2

    def _parse_newcolumntype(self, s: str, i: int, j: int) -> int | None:
        name, k = read_braced(s, j)
        nargs_s, k = read_opt(s, k)
        body, k2 = read_braced(s, k)
        if not name or body is None:
            return None
        self.coltypes[name.strip()] = (int(nargs_s) if nargs_s and nargs_s.strip().isdigit() else 0, body)
        return k2

    # ---------------- expansion ----------------
    def _read_macro_args(self, m: _Macro, s: str, j: int) -> tuple[list[str], int] | None:
        args: list[str] = []
        k = j
        if m.kind == "xparse" and m.spec:
            for typ, dflt in m.spec:
                if typ == "s":
                    st, k = read_star(s, k)
                    args.append(r"\BooleanTrue" if st else r"\BooleanFalse")
                elif typ in ("o", "O"):
                    o, k = read_opt(s, k)
                    args.append(o if o is not None else (dflt if typ == "O" else "-NoValue-"))
                else:
                    a, k = read_arg(s, k)
                    if a is None:
                        return None
                    args.append(a)
            return args, k
        if m.kind == "def" and m.params is not None:
            # delimited parameters
            prefix = m.params[0]
            if prefix.strip():
                kk = k
                if not s.startswith(prefix, kk):
                    kk2 = skip_ws(s, kk)
                    if not s.startswith(prefix.strip(), kk2):
                        return None
                    kk = kk2 + len(prefix.strip())
                else:
                    kk += len(prefix)
                k = kk
            for idx in range(1, len(m.params)):
                delim = m.params[idx]
                if delim == "" or (idx == len(m.params) - 1 and delim == ""):
                    a, k = read_arg(s, k)
                    if a is None:
                        return None
                    args.append(a)
                    continue
                # read until delimiter at depth 0
                depth = 0
                p = k
                n = len(s)
                found = -1
                while p < n:
                    c = s[p]
                    if c == "\\" and not s.startswith(delim, p):
                        p += 2
                        continue
                    if c == "{":
                        depth += 1
                    elif c == "}":
                        if depth == 0:
                            break
                        depth -= 1
                    if depth == 0 and s.startswith(delim, p):
                        found = p
                        break
                    p += 1
                if found < 0:
                    return None
                a = s[k:found]
                if a.startswith("{") and a.endswith("}") and read_group(a, 0)[1] == len(a):
                    a = a[1:-1]
                args.append(a)
                k = found + len(delim)
            return args, k
        if m.kind == "paired":
            st, k = read_star(s, k)
            o, k = read_opt(s, k)
            a, k = read_arg(s, k)
            if a is None:
                return None
            return [a, "*" if st else "", o or ""], k
        if m.default is not None and m.nargs > 0:
            o, k = read_opt(s, k)
            args.append(o if o is not None else m.default)
            for _ in range(m.nargs - 1):
                a, k = read_arg(s, k)
                if a is None:
                    return None
                args.append(a)
            return args, k
        for _ in range(m.nargs):
            a, k = read_arg(s, k)
            if a is None:
                return None
            args.append(a)
        return args, k

    @staticmethod
    def _substitute(body: str, args: list[str]) -> str:
        def rep(mm: re.Match[str]) -> str:
            t = mm.group(1)
            if t == "#":
                return "#"
            idx = int(t) - 1
            return args[idx] if 0 <= idx < len(args) else ""
        return re.sub(r"#(#|[1-9])", rep, body)

    def _expand_macro(self, m: _Macro, args: list[str]) -> str:
        if m.kind == "paired":
            l, r = m.params or ("(", ")")
            inner = self._substitute(m.body, [args[0]]) if m.body else args[0]
            if args[2]:
                return f"{l} {inner} {r} "
            return f"\\left{l} {inner} \\right{r} "
        exp = self._substitute(m.body, args)
        # xparse helpers
        if m.kind == "xparse":
            exp = self._eval_xparse_conditionals(exp)
        # TeX: a control word at the end followed by letters needs a separator
        return exp

    def _eval_xparse_conditionals(self, s: str) -> str:
        pat = re.compile(r"\\(IfNoValueTF|IfNoValueT|IfNoValueF|IfValueTF|IfValueT|IfValueF|IfBooleanTF|IfBooleanT|IfBooleanF)\b")
        for _ in range(200):
            m = pat.search(s)
            if not m:
                break
            name = m.group(1)
            a, k = read_arg(s, m.end())
            if a is None:
                break
            t, k2 = read_arg(s, k) if name.endswith(("TF", "T", "F")) else (None, k)
            f = None
            if name.endswith("TF"):
                f, k2 = read_arg(s, k2)
            novalue = a.strip() == "-NoValue-"
            boolean = a.strip() == r"\BooleanTrue"
            if name.startswith("IfNoValue"):
                cond = novalue
            elif name.startswith("IfValue"):
                cond = not novalue
            else:
                cond = boolean
            if name.endswith("TF"):
                rep = t if cond else f
            elif name.endswith("T"):
                rep = t if cond else ""
            else:
                rep = t if not cond else ""
            s = s[: m.start()] + (rep or "") + s[k2:]
        return s

    def _find_cond_branches(self, s: str, j: int) -> tuple[list[str], int] | None:
        """From j (just after \\ifX), return branches split at \\else/\\or and index after \\fi."""
        depth = 0
        n = len(s)
        k = j
        seg_start = j
        branches: list[str] = []
        while k < n:
            if s[k] == "\\":
                name, k2 = read_cmd(s, k)
                if name in CONDITIONAL_PRIMITIVES or (name.startswith("if") and name[2:] in self.bools):
                    depth += 1
                elif name == "fi":
                    if depth == 0:
                        branches.append(s[seg_start:k])
                        return branches, k2
                    depth -= 1
                elif name in ("else", "or") and depth == 0:
                    branches.append(s[seg_start:k])
                    seg_start = k2
                k = k2
                continue
            k += 1
        return None

    def _eval_condition(self, name: str, s: str, j: int) -> tuple[bool | int, int]:
        if name == "iftrue":
            return True, j
        if name == "iffalse":
            return False, j
        if name[2:] in self.bools:
            return self.bools[name[2:]], j
        if name in ("ifmmode", "ifinner", "ifvmode", "ifvoid", "ifhbox", "ifvbox", "ifeof"):
            return False, j
        if name == "ifhmode":
            return True, j
        if name == "ifdefined":
            k = skip_ws(s, j)
            if k < len(s) and s[k] == "\\":
                nm, k2 = read_cmd(s, k)
                return (nm in self.macros) or nm in TEXT_SYMBOLS, k2
            return True, j
        if name == "ifx":
            k = skip_ws(s, j)
            toks = []
            for _ in range(2):
                k = skip_ws(s, k)
                if k < len(s) and s[k] == "\\":
                    nm, k = read_cmd(s, k)
                    toks.append(nm)
                elif k < len(s):
                    toks.append(s[k])
                    k += 1
            if len(toks) == 2:
                a, b = toks
                if b in ("undefined", "relax") and a not in self.macros:
                    return True, k
                ma, mb = self.macros.get(a), self.macros.get(b)
                if ma and mb:
                    return ma.body == mb.body, k
                return a == b, k
            return True, k
        if name in ("ifnum", "ifdim"):
            m = re.match(r"\s*(-?\d+(?:\.\d+)?)\s*(?:pt|cm|mm)?\s*([<=>])\s*(-?\d+(?:\.\d+)?)\s*(?:pt|cm|mm)?", s[j:])
            if m:
                a, op, b = float(m.group(1)), m.group(2), float(m.group(3))
                val = a < b if op == "<" else (a > b if op == ">" else a == b)
                return val, j + m.end()
            self.res.warnings.append(f"Không đánh giá được \\{name} — chọn nhánh đầu.")
            return True, j
        if name == "ifodd":
            m = re.match(r"\s*(-?\d+)", s[j:])
            if m:
                return int(m.group(1)) % 2 == 1, j + m.end()
            return True, j
        if name == "ifcase":
            m = re.match(r"\s*(\d+)", s[j:])
            if m:
                return int(m.group(1)), j + m.end()
            return 0, j
        self.res.warnings.append(f"Không đánh giá được \\{name} — chọn nhánh đầu.")
        return True, j

    def _ifthenelse(self, s: str, j: int) -> tuple[str, int] | None:
        cond, k = read_braced(s, j)
        a, k = read_braced(s, k)
        b, k2 = read_braced(s, k)
        if cond is None or a is None or b is None:
            return None
        c = cond.strip()
        val = True
        mb = re.fullmatch(r"\\boolean\{(\w+)\}", c)
        me = re.fullmatch(r"\\equal\{(.*)\}\{(.*)\}", c, re.S)
        mn = re.fullmatch(r"(-?\d+)\s*([<=>])\s*(-?\d+)", c)
        if mb:
            val = self.bools.get(mb.group(1), False)
        elif me:
            val = me.group(1).strip() == me.group(2).strip()
        elif mn:
            x, op, y = int(mn.group(1)), mn.group(2), int(mn.group(3))
            val = x < y if op == "<" else (x > y if op == ">" else x == y)
        return (a if val else b), k2

    def expand(self, s: str, *, in_tikz: bool = False, exclude: frozenset[str] = frozenset()) -> str:
        out: list[str] = []
        i = 0
        while True:
            j0 = s.find("\\", i)
            if j0 == -1:
                out.append(s[i:])
                break
            out.append(s[i:j0])
            i = j0
            name, j = read_cmd(s, i)
            if not name:
                out.append(s[i:j])
                i = j
                continue
            if self.count > self.MAX_EXPANSIONS:
                out.append(s[i:])
                self.res.warnings.append("Dừng mở rộng macro: vượt giới hạn (có thể định nghĩa đệ quy vô hạn).")
                break

            new_i: int | None = None
            handled = True
            if name in ("newcommand", "renewcommand", "DeclareRobustCommand", "newrobustcmd", "renewrobustcmd"):
                new_i = self._parse_newcommand(s, i, j, provide=False)
            elif name == "providecommand":
                new_i = self._parse_newcommand(s, i, j, provide=True)
            elif name in ("def", "gdef", "edef", "xdef"):
                new_i = self._parse_def(s, i, j)
            elif name in ("long", "global", "protected", "outer"):
                new_i = j
            elif name in ("let", "global\\let"):
                new_i = self._parse_let(s, i, j)
            elif name in ("NewDocumentCommand", "RenewDocumentCommand", "DeclareDocumentCommand", "NewExpandableDocumentCommand"):
                new_i = self._parse_newdocumentcommand(s, i, j, provide=False)
            elif name == "ProvideDocumentCommand":
                new_i = self._parse_newdocumentcommand(s, i, j, provide=True)
            elif name in ("newenvironment", "renewenvironment"):
                new_i = self._parse_newenvironment(s, i, j, provide=False)
            elif name == "provideenvironment":
                new_i = self._parse_newenvironment(s, i, j, provide=True)
            elif name in ("NewDocumentEnvironment", "RenewDocumentEnvironment", "DeclareDocumentEnvironment"):
                new_i = self._parse_newdocumentenvironment(s, i, j)
            elif name == "DeclareMathOperator":
                new_i = self._parse_mathop(s, i, j)
            elif name in ("DeclarePairedDelimiter", "DeclarePairedDelimiterX"):
                new_i = self._parse_paired(s, i, j, x=name.endswith("X"))
            elif name == "newcolumntype":
                new_i = self._parse_newcolumntype(s, i, j)
            elif name == "newif":
                k = skip_ws(s, j)
                if k < len(s) and s[k] == "\\":
                    nm, k2 = read_cmd(s, k)
                    if nm.startswith("if"):
                        self.bools[nm[2:]] = False
                    new_i = k2
            elif name == "csname":
                m = re.match(r"\s*([A-Za-z@]+)\s*\\endcsname", s[j:])
                if m:
                    s = "\\" + m.group(1) + " " + s[j + m.end():]
                    i = 0
                    continue
                new_i = j
            elif name.endswith("true") and name[:-4] in self.bools:
                self.bools[name[:-4]] = True
                new_i = j
            elif name.endswith("false") and name[:-5] in self.bools:
                self.bools[name[:-5]] = False
                new_i = j
            elif name == "setboolean":
                a, k = read_braced(s, j)
                b, k2 = read_braced(s, k)
                if a and b is not None:
                    self.bools[a.strip()] = b.strip().lower() == "true"
                    new_i = k2
            elif name in ("newboolean", "provideboolean"):
                a, k2 = read_braced(s, j)
                if a:
                    self.bools.setdefault(a.strip(), False)
                    new_i = k2
            elif name == "ifthenelse":
                r = self._ifthenelse(s, j)
                if r is not None:
                    rep, k2 = r
                    s = rep + s[k2:]
                    i = 0
                    self.count += 1
                    continue
            elif (name in CONDITIONAL_PRIMITIVES or (name.startswith("if") and name[2:] in self.bools)) and name not in ("iff",):
                val, k = self._eval_condition(name, s, j)
                br = self._find_cond_branches(s, k)
                if br is not None:
                    branches, k2 = br
                    if isinstance(val, bool):
                        chosen = branches[0] if val else (branches[1] if len(branches) > 1 else "")
                    else:
                        chosen = branches[val] if val < len(branches) else (branches[-1] if len(branches) > 1 else "")
                    s = chosen + s[k2:]
                    i = 0
                    self.count += 1
                    continue
                new_i = k
            elif name == "begin":
                k = skip_spaces_only(s, j)
                if k < len(s) and s[k] == "{":
                    env, k2 = read_group(s, k)
                    env = env.strip()
                    if env in ("tikzpicture", "axis", "pgfpicture", "circuitikz") and not in_tikz:
                        ce, after = find_env_end(s, k2, env)
                        inner = s[k2:ce]
                        local = set(re.findall(r"\\foreach\s*\\([A-Za-z]+)", inner))
                        local |= set(re.findall(r"/\\([A-Za-z]+)", inner))
                        local |= set(re.findall(r"\\pgfmathsetmacro\s*\{?\\([A-Za-z]+)", inner))
                        local |= set(re.findall(r"\\(?:def|newcommand)\s*\{?\\([A-Za-z]+)", inner))
                        out.append(s[i:k2])
                        out.append(self.expand(inner, in_tikz=True, exclude=frozenset(local)))
                        out.append(s[ce:after])
                        i = after
                        continue
                    if env in self.envs and env not in ("document",):
                        ed = self.envs[env]
                        args: list[str] = []
                        kk = k2
                        ok = True
                        if ed.spec:
                            fake = _Macro(name=env, kind="xparse", spec=ed.spec)
                            r = self._read_macro_args(fake, s, kk)
                            if r is None:
                                ok = False
                            else:
                                args, kk = r
                        elif ed.nargs:
                            fake = _Macro(name=env, nargs=ed.nargs, default=ed.default)
                            r = self._read_macro_args(fake, s, kk)
                            if r is None:
                                ok = False
                            else:
                                args, kk = r
                        if ok:
                            ce, after = find_env_end(s, kk, env)
                            content = s[kk:ce]
                            b = self._substitute(ed.begin, args)
                            e = self._substitute(ed.end, args)
                            if ed.spec:
                                b = self._eval_xparse_conditionals(b)
                                e = self._eval_xparse_conditionals(e)
                            sep1 = " " if re.search(r"\\[A-Za-z@]+$", b) and content[:1].isalpha() else ""
                            sep2 = " " if re.search(r"\\[A-Za-z@]+$", content) and e[:1].isalpha() else ""
                            s = "{" + b + sep1 + content + sep2 + e + "}" + s[after:]
                            i = 0
                            self.count += 1
                            self.res.stats["environments_expanded"] = self.res.stats.get("environments_expanded", 0) + 1
                            continue
                handled = False
            elif name in self.macros and name not in exclude and name not in self.disabled:
                m = self.macros[name]
                self.per_macro[name] = self.per_macro.get(name, 0) + 1
                if self.per_macro[name] > self.MAX_PER_MACRO:
                    self.disabled.add(name)
                    self.res.warnings.append(f"Macro \\{name} tự gọi lại vô hạn — đã dừng mở rộng.")
                    handled = False
                elif m.soft and in_tikz:
                    handled = False
                else:
                    r = self._read_macro_args(m, s, j)
                    if r is not None:
                        args, k2 = r
                        exp = self._expand_macro(m, args)
                        # keep a space boundary if the expansion ends with a control word and next is a letter
                        nxt = s[k2:k2 + 1]
                        if re.search(r"\\[A-Za-z@]+$", exp) and nxt.isalpha():
                            exp += " "
                        s = exp + s[k2:]
                        i = 0
                        self.count += 1
                        self.res.stats["macros_expanded"] = self.res.stats.get("macros_expanded", 0) + 1
                        continue
                    handled = False
            else:
                handled = False

            if handled and new_i is not None:
                # drop the definition from output
                i = new_i
                continue
            out.append(s[i:j])
            i = j
            # keep the remaining text compact for splicing performance
            if i > 200000:
                out.append(s[i:])
                s = ""
                i = 0
                break
        return "".join(out)


# =====================================================================================
# 6. NGỮ CẢNH DỰNG KHỐI
# =====================================================================================
@dataclass
class _Pending:
    text: str = ""
    spec: dict[str, Any] = field(default_factory=dict)
    used: bool = True


@dataclass
class _Ctx:
    align: str | None = None
    left: int = 0
    hanging: int = 0
    spans: tuple[int, ...] = ()
    box: dict[str, Any] | None = None
    shd: str | None = None
    pending: _Pending | None = None
    first_para_spec: dict[str, Any] | None = None
    italic: bool = False
    bold: bool = False
    in_immini: bool = False
    list_depth: int = 0
    enum_depth: int = 0
    item_depth: int = 0
    keep_lines: bool = False
    float_kind: str | None = None

    def child(self, **kw: Any) -> "_Ctx":
        d = dict(self.__dict__)
        d.update(kw)
        d["pending"] = kw.get("pending", None)
        d["first_para_spec"] = kw.get("first_para_spec", None)
        return _Ctx(**d)


@dataclass
class _ICtx:
    italic: bool = False
    bold: bool = False
    bold_via_span: bool = False
    in_cell: bool = False
    no_markers: bool = False
    plain: bool = False


BLOCK_ENV_NAMES_FOR_GROUP = re.compile(
    r"\\begin\s*\{(tikzpicture|tabular\*?|tabularx|tabulary|longtable|itemize|enumerate|description|center|"
    r"flushleft|flushright|figure|table|equation\*?|align\*?|gather\*?|multline\*?|minipage|tcolorbox|"
    r"verbatim|lstlisting|multicols|tasks|ex|bt|vd)\}|\n[ \t]*\n|\\par\b|\\\[|\$\$|\\section|\\choice|\\immini"
)

THEOREM_BUILTINS_VI: dict[str, tuple[str, bool, str]] = {
    # name: (title, numbered, style)  style: exam|plain|definition|remark
    "ex": ("Câu", True, "exam"), "bt": ("Bài", True, "exam"), "vd": ("Ví dụ", True, "exam"),
    "vidu": ("Ví dụ", True, "exam"), "baitap": ("Bài tập", True, "exam"), "cauhoi": ("Câu hỏi", True, "exam"),
    "bai": ("Bài", True, "exam"), "cau": ("Câu", True, "exam"), "question": ("Câu", True, "exam"),
    "dn": ("Định nghĩa", True, "definition"), "dinhnghia": ("Định nghĩa", True, "definition"),
    "dl": ("Định lý", True, "definition"), "dinhli": ("Định lí", True, "definition"),
    "dinhly": ("Định lý", True, "definition"), "tc": ("Tính chất", True, "definition"),
    "tinhchat": ("Tính chất", True, "definition"), "hq": ("Hệ quả", True, "definition"),
    "hequa": ("Hệ quả", True, "definition"), "bd": ("Bổ đề", True, "definition"),
    "bode": ("Bổ đề", True, "definition"), "md": ("Mệnh đề", True, "definition"),
    "menhde": ("Mệnh đề", True, "definition"), "nx": ("Nhận xét", False, "remark"),
    "nhanxet": ("Nhận xét", False, "remark"), "chuy": ("Chú ý", False, "remark"),
    "chuyy": ("Chú ý", False, "remark"), "luuy": ("Lưu ý", False, "remark"), "note": ("Chú ý", False, "remark"),
    "ghichu": ("Ghi chú", False, "remark"), "pp": ("Phương pháp", False, "remark"),
    "phuongphap": ("Phương pháp", False, "remark"), "kn": ("Khái niệm", True, "definition"),
    "khainiem": ("Khái niệm", True, "definition"), "vdm": ("Ví dụ mẫu", True, "exam"),
    "hoatdong": ("Hoạt động", True, "exam"), "hd": ("Hoạt động", True, "exam"),
    "luyentap": ("Luyện tập", True, "exam"), "vandung": ("Vận dụng", True, "exam"),
    "giai": ("Giải", False, "solution"), "loigiai": ("Lời giải", False, "solution"),
    "huongdan": ("Hướng dẫn", False, "solution"), "solution": ("Lời giải", False, "solution"),
    "dapan": ("Đáp án", False, "solution"),
}
THEOREM_BUILTINS_EN: dict[str, tuple[str, str, bool, str]] = {
    # name: (en, vi, numbered, style)
    "theorem": ("Theorem", "Định lý", True, "plain"), "thm": ("Theorem", "Định lý", True, "plain"),
    "lemma": ("Lemma", "Bổ đề", True, "plain"), "lem": ("Lemma", "Bổ đề", True, "plain"),
    "corollary": ("Corollary", "Hệ quả", True, "plain"), "cor": ("Corollary", "Hệ quả", True, "plain"),
    "proposition": ("Proposition", "Mệnh đề", True, "plain"), "prop": ("Proposition", "Mệnh đề", True, "plain"),
    "definition": ("Definition", "Định nghĩa", True, "definition"), "defn": ("Definition", "Định nghĩa", True, "definition"),
    "example": ("Example", "Ví dụ", True, "definition"), "exercise": ("Exercise", "Bài tập", True, "definition"),
    "problem": ("Problem", "Bài toán", True, "definition"), "remark": ("Remark", "Nhận xét", True, "remark"),
    "rem": ("Remark", "Nhận xét", True, "remark"), "notation": ("Notation", "Ký hiệu", True, "remark"),
    "claim": ("Claim", "Khẳng định", True, "plain"), "conjecture": ("Conjecture", "Giả thuyết", True, "plain"),
    "fact": ("Fact", "Sự kiện", True, "plain"), "observation": ("Observation", "Quan sát", True, "remark"),
    "assumption": ("Assumption", "Giả thiết", True, "definition"), "case": ("Case", "Trường hợp", True, "remark"),
    "algorithm": ("Algorithm", "Thuật toán", True, "definition"), "axiom": ("Axiom", "Tiên đề", True, "plain"),
}

LIST_ENVS = {"itemize", "enumerate", "description", "compactitem", "compactenum", "compactdesc",
             "asparaenum", "asparaitem", "inparaenum", "inparaitem", "enumerate*", "itemize*",
             "description*", "questions", "parts", "subparts", "subsubparts", "steps", "checklist"}
CHOICE_LIST_ENVS = {"choices", "oneparchoices", "checkboxes", "oneparcheckboxes"}
MATH_DISPLAY_ENVS = {"equation", "equation*", "align", "align*", "gather", "gather*", "multline",
                     "multline*", "flalign", "flalign*", "alignat", "alignat*", "eqnarray", "eqnarray*",
                     "displaymath", "dmath", "dmath*", "dgroup", "dgroup*", "empheq", "IEEEeqnarray",
                     "IEEEeqnarray*", "xalignat", "xalignat*", "xxalignat"}
TABLE_ENVS = {"tabular", "tabular*", "tabularx", "tabulary", "longtable", "longtable*", "tabu", "longtabu",
              "supertabular", "xtabular", "NiceTabular", "NiceTabular*", "tblr", "longtblr", "talltblr",
              "array", "TAB"}
TRANSPARENT_ENVS = {"document", "figure", "figure*", "table", "table*", "wrapfigure", "wraptable",
                    "subfigure", "subtable", "minipage", "adjustbox", "landscape", "sidewaysfigure",
                    "sidewaystable", "spacing", "singlespace", "onehalfspace", "doublespace", "otherlanguage",
                    "otherlanguage*", "samepage", "titlepage", "varwidth", "column", "columns", "appendices",
                    "CJK", "CJK*", "multicols*", "turn", "rotate", "sideways", "flushleft*", "linenumbers",
                    "subequations", "groupplot", "scope", "frame", "block", "alertblock", "exampleblock",
                    "onlyenv", "altenv", "fullwidth", "widetext", "esmcomment", "boxedminipage", "wideflush",
                    "parcolumns", "paracol", "vietnamese", "english", "biblist", "minipage*", "small",
                    "footnotesize", "scriptsize", "tiny", "large", "Large", "LARGE", "huge", "Huge",
                    "normalsize", "bfseries", "itshape", "sffamily", "ttfamily", "em", "it", "bf", "sl",
                    "upshape", "scshape", "slshape", "mdseries", "rmfamily", "color", "picture",
                    "addmargin", "addmargin*", "indentation", "narrow", "changemargin", "adjustwidth",
                    "Form", "multipage", "noindent", "flushleft", "flushright", "center", "centering",
                    "raggedleft", "raggedright", "quote", "quotation", "verse", "abstract"}
BOX_ENVS = {"tcolorbox", "mdframed", "framed", "shaded", "shaded*", "leftbar", "fminipage", "fbox",
            "framedbox", "tcbitemize", "boxed", "colorbox", "snugshade", "oframed", "khung", "khungdl",
            "bclogo", "notebox", "tikzbox", "dangtoan", "boxdn", "boxdl", "boxtc"}
SIZE_SWITCH_PT = {"tiny": 6.0, "scriptsize": 8.0, "footnotesize": 10.0, "small": 10.95, "normalsize": 12.0,
                  "large": 14.4, "Large": 17.28, "LARGE": 20.74, "huge": 24.88, "Huge": 24.88}


# =====================================================================================
# 7. PHASE B — DỰNG CẤU TRÚC
# =====================================================================================
class LatexPreprocessor:
    def __init__(self, *, base_dir: Path | None = None, show_solutions: bool = True) -> None:
        self.base_dir = base_dir
        self.show_solutions = show_solutions
        self.res = PreprocessResult(text="")
        self.expander = _MacroExpander(self.res)
        self.items: list[tuple[str, Any, Any]] = []  # ("para", text, spec) | ("block", raw, None)
        self._para_idx: dict[str, int] = {}
        self._span_idx: dict[str, int] = {}
        self.verbatims: list[tuple[str, str]] = []
        # document configuration
        self.doc_class = "article"
        self.class_opts: set[str] = set()
        self.packages: set[str] = set()
        self.lang = "vi"
        self.title: str | None = None
        self.author: str | None = None
        self.date: str | None = None
        self.colors: dict[str, str] = {}
        self.theorems: dict[str, dict[str, Any]] = {}
        self.theorem_style = "plain"
        self.names: dict[str, str] = {}
        self.secnumdepth = 3
        self.textwidth = TEXTWIDTH_TWIPS_DEFAULT
        self.appendix = False
        # counters
        self.counters: dict[str, int] = {}
        self.counter_parent: dict[str, str] = {}
        self.the_formats: dict[str, str] = {}
        self.label_values: dict[str, tuple[str, str]] = {}
        self.current_label: tuple[str, str] = ("", "")
        self.headings: list[tuple[int, str]] = []
        self.bib_keys: dict[str, int] = {}
        self.enum_labels: dict[int, str] = {}
        self.item_labels: dict[int, str] = {}
        self.tasks_label: str | None = None
        self.tasks_cols = 1
        self.list_resume: dict[str, int] = {}
        self.page: dict[str, Any] = {"pagestyle": "plain", "header": {}, "footer": {}, "first_empty": False}
        self._config_only = False
        self._pending_vspace = 0
        self._init_counters()

    # ---------------- counters ----------------
    def _init_counters(self) -> None:
        for c, parent in [("part", ""), ("chapter", ""), ("section", "chapter"), ("subsection", "section"),
                          ("subsubsection", "subsection"), ("paragraph", "subsubsection"),
                          ("subparagraph", "paragraph"), ("equation", ""), ("figure", ""), ("table", ""),
                          ("footnote", ""), ("enumi", ""), ("enumii", "enumi"), ("enumiii", "enumii"),
                          ("enumiv", "enumiii"), ("ex", ""), ("page", ""), ("mpfootnote", ""),
                          ("question", ""), ("partno", "question")]:
            self.counters[c] = 0
            if parent:
                self.counter_parent[c] = parent

    def _children(self, c: str) -> list[str]:
        return [k for k, v in self.counter_parent.items() if v == c]

    def _reset_children(self, c: str) -> None:
        for ch in self._children(c):
            self.counters[ch] = 0
            self._reset_children(ch)

    def step(self, c: str) -> None:
        self.counters[c] = self.counters.get(c, 0) + 1
        self._reset_children(c)

    def refstep(self, c: str, kind: str = "") -> str:
        self.step(c)
        val = self.the(c)
        self.current_label = (val, kind or c)
        return val

    def the(self, c: str) -> str:
        if c in self.the_formats:
            return self._plain(self.the_formats[c])
        n = self.counters.get(c, 0)
        if c == "part":
            return _roman(n).upper()
        if c == "chapter" and self.appendix:
            return _alph(n).upper()
        if c == "section" and self.appendix and self.doc_class not in ("book", "report"):
            return _alph(n).upper()
        if c == "enumii":
            return _alph(n)
        if c == "enumiii":
            return _roman(n)
        if c == "enumiv":
            return _alph(n).upper()
        if c == "footnote":
            return str(n)
        parent = self.counter_parent.get(c, "")
        if c in ("section",) and self.doc_class not in ("book", "report"):
            parent = ""
        if c in ("equation", "figure", "table") and self.doc_class in ("book", "report") and "chapter" not in self.counter_parent.get(c, "chapter"):
            parent = "chapter"
        if parent and parent in self.counters and not (c in ("enumii", "enumiii", "enumiv", "partno")):
            if parent == "chapter" and self.counters.get("chapter", 0) == 0:
                return str(n)
            return f"{self.the(parent)}.{n}"
        return str(n)

    # ---------------- spec registries ----------------
    def para_marker(self, spec: dict[str, Any]) -> str:
        if not spec:
            return ""
        key = json.dumps(spec, sort_keys=True, ensure_ascii=False)
        idx = self._para_idx.get(key)
        if idx is None:
            idx = len(self.res.para_specs)
            self.res.para_specs.append(dict(spec))
            self._para_idx[key] = idx
        return MK_PARA + idx_char(idx)

    def span(self, spec: dict[str, Any], text: str) -> str:
        if not spec or text == "":
            return text
        key = json.dumps(spec, sort_keys=True, ensure_ascii=False)
        idx = self._span_idx.get(key)
        if idx is None:
            idx = len(self.res.span_specs)
            self.res.span_specs.append(dict(spec))
            self._span_idx[key] = idx
        return MK_SPAN + idx_char(idx) + text + MK_END

    def span_id(self, spec: dict[str, Any]) -> int:
        key = json.dumps(spec, sort_keys=True, ensure_ascii=False)
        idx = self._span_idx.get(key)
        if idx is None:
            idx = len(self.res.span_specs)
            self.res.span_specs.append(dict(spec))
            self._span_idx[key] = idx
        return idx

    def _unknown(self, name: str) -> None:
        self.res.unknown_commands[name] = self.res.unknown_commands.get(name, 0) + 1

    def _stat(self, key: str, n: int = 1) -> None:
        self.res.stats[key] = self.res.stats.get(key, 0) + n

    # ---------------- colors ----------------
    def color_hex(self, spec: str, model: str | None = None) -> str | None:
        spec = (spec or "").strip()
        if model:
            if "," in model:  # mixed model list — rare
                model = model.split(",")[0]
            return _model_to_hex(model.strip(), spec)
        if not spec:
            return None
        if spec.startswith("#") and re.fullmatch(r"#[0-9A-Fa-f]{6}", spec):
            return spec[1:].upper()
        # xcolor expression: red!50!blue, red!30, -red
        parts = spec.split("!")
        base = self._named_color(parts[0].lstrip("-"))
        if base is None:
            return None
        r, g, b = _hex_to_rgb(base)
        idx = 1
        while idx < len(parts):
            try:
                pct = float(parts[idx]) / 100.0
            except ValueError:
                break
            other = "FFFFFF"
            if idx + 1 < len(parts):
                o = self._named_color(parts[idx + 1])
                other = o or "FFFFFF"
                idx += 2
            else:
                idx += 1
            r2, g2, b2 = _hex_to_rgb(other)
            r, g, b = r * pct + r2 * (1 - pct), g * pct + g2 * (1 - pct), b * pct + b2 * (1 - pct)
        if spec.startswith("-"):
            r, g, b = 255 - r, 255 - g, 255 - b
        return _rgb_to_hex(r, g, b)

    def _named_color(self, name: str) -> str | None:
        name = name.strip()
        if name in self.colors:
            return self.colors[name]
        if name in XCOLOR_NAMED:
            return XCOLOR_NAMED[name]
        if name.lower() in XCOLOR_NAMED:
            return XCOLOR_NAMED[name.lower()]
        if re.fullmatch(r"[0-9A-Fa-f]{6}", name):
            return name.upper()
        return None

    # =================================================================================
    # ENTRY
    # =================================================================================
    def run(self, src: str) -> PreprocessResult:
        s = src.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n").replace("\t", "    ")
        s = s.lstrip("\ufeff")
        s = self._expand_inputs(s, depth=0)
        s = self._protect_verbatim(s)
        s = self._strip_comments(s)
        if re.search(r"[ăâđêôơưĂÂĐÊÔƠƯàáảãạằắẳẵặầấẩẫậèéẻẽẹềếểễệìíỉĩịòóỏõọồốổỗộờớởỡợùúủũụừứửữựỳýỷỹỵ]", s) or re.search(
            r"\\usepackage(\[[^\]]*\])?\{(vietnam|vntex)\}|vietnamese", s
        ):
            self.lang = "vi"
        else:
            self.lang = "en"

        if r"\begin{document}" in s:
            pre, body = s.split(r"\begin{document}", 1)
            body = body.split(r"\end{document}", 1)[0]
        else:
            pre, body = "", s
        pre_x = self.expander.expand(pre)
        self._scan_preamble(pre_x)
        geo = compute_page_geometry(self.page)
        self.textwidth = geo["width"] - geo["left"] - geo["right"]
        body_x = self.expander.expand(body)
        self._stat("input_chars", len(src))

        root = _Ctx()
        self._blocks(body_x, root)
        text = self._assemble()
        text = self._resolve_refs(text)
        # safety: engine must never see "$$" from adjacent inline maths
        text = text.replace("$$", "$\u200b$")
        self.res.footnotes = [self._resolve_refs(f).replace("$$", "$\u200b$") for f in self.res.footnotes]
        self.page["headings"] = list(self.headings)
        self.res.text = text
        self.res.page = self.page
        self.res.page["textwidth"] = self.textwidth
        return self.res

    # ---------------- stage 1 helpers ----------------
    def _expand_inputs(self, s: str, depth: int) -> str:
        if depth > 8:
            return s

        def rep(m: re.Match[str]) -> str:
            name = (m.group(2) or m.group(3) or "").strip()
            if not name:
                return ""
            if self.base_dir is not None:
                p = Path(name)
                cands = [p] if p.suffix else [p.with_suffix(".tex"), p]
                for c in cands:
                    f = (self.base_dir / c)
                    try:
                        if f.is_file():
                            child = f.read_text(encoding="utf-8-sig", errors="replace")
                            return "\n" + self._expand_inputs(child, depth + 1) + "\n"
                    except OSError:
                        pass
            self.res.warnings.append(f"Bỏ qua \\{m.group(1)}{{{name}}} (không có file đi kèm khi dán mã).")
            return ""

        return re.sub(r"\\(input|include|subfile|import)\s*(?:\{([^{}]+)\}|\s+([^\s{}\\]+))", rep, s)

    def _protect_verbatim(self, s: str) -> str:
        def rep_env(m: re.Match[str]) -> str:
            env = m.group(1)
            body = m.group(2)
            if env == "comment":
                return ""
            if body.startswith("\n"):
                body = body[1:]
            body = re.sub(r"^\[[^\]\n]*\]\n?", "", body) if env in ("lstlisting", "Verbatim") else body
            if env == "minted":
                body = re.sub(r"^(\[[^\]]*\])?\{[^}]*\}\n?", "", body)
            self.verbatims.append(("block", body.rstrip("\n")))
            return "\n\n" + _VERB_BLOCK + idx_char(len(self.verbatims) - 1) + "\n\n"

        s = re.sub(r"\\begin\{(verbatim\*?|Verbatim|BVerbatim|lstlisting|minted|alltt|comment|filecontents\*?)\}(.*?)\\end\{\1\}",
                   rep_env, s, flags=re.S)

        def rep_inline(m: re.Match[str]) -> str:
            self.verbatims.append(("inline", m.group(4)))
            return _VERB_INLINE + idx_char(len(self.verbatims) - 1)

        s = re.sub(r"\\(verb|lstinline|mintinline\{[^}]*\})(\*?)(?:\[[^\]]*\])?([^\sA-Za-z{*])(.*?)\3",
                   rep_inline, s)
        return s

    @staticmethod
    def _strip_comments(s: str) -> str:
        """Remove TeX comments with TeX line-end semantics (keeps blank lines = paragraph breaks)."""
        out: list[str] = []
        i = 0
        n = len(s)
        while i < n:
            p = s.find("%", i)
            if p == -1:
                out.append(s[i:])
                break
            bs = 0
            q = p - 1
            while q >= 0 and s[q] == "\\":
                bs += 1
                q -= 1
            if bs % 2 == 1:
                out.append(s[i : p + 1])
                i = p + 1
                continue
            seg = s[i:p]
            e = s.find("\n", p)
            line_start = s.rfind("\n", 0, p) + 1
            whole_line = s[line_start:p].strip() == ""
            if whole_line:
                out.append(seg.rstrip(" \t"))
                i = n if e == -1 else e + 1
                continue
            out.append(seg)
            if e == -1:
                break
            k = e + 1
            while k < n and s[k] in " \t":
                k += 1
            i = e if (k < n and s[k] == "\n") or k >= n else k
        return "".join(out)

    # ---------------- preamble ----------------
    def _scan_preamble(self, pre: str) -> None:
        m = re.search(r"\\documentclass\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}", pre)
        if m:
            self.doc_class = m.group(2).strip()
            self.class_opts = {o.strip() for o in (m.group(1) or "").split(",") if o.strip()}
            if self.doc_class in ("book", "report", "scrbook", "scrreprt", "memoir"):
                self.doc_class = "book" if "book" in self.doc_class else "report"
                for c in ("equation", "figure", "table"):
                    self.counter_parent[c] = "chapter"
                self.counter_parent["section"] = "chapter"
            if "twocolumn" in self.class_opts:
                self.page["twocolumn"] = True
            if "landscape" in self.class_opts:
                self.page["landscape"] = True
            for p in ("a4paper", "a5paper", "letterpaper", "b5paper", "legalpaper", "a3paper"):
                if p in self.class_opts:
                    self.page["paper"] = p
        for pm in re.finditer(r"\\usepackage\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}", pre):
            for name in pm.group(2).split(","):
                name = name.strip()
                self.packages.add(name)
                if name == "geometry" and pm.group(1):
                    self._geometry(pm.group(1))
                if name in ("vietnam", "vntex") or (name == "babel" and "vietnam" in (pm.group(1) or "")):
                    self.lang = "vi"
        # run preamble-level config commands through the block scanner in "config only" mode
        self._config_only = True
        self._blocks(pre, _Ctx())
        self._config_only = False
        self.items.clear()

    def _geometry(self, opts: str) -> None:
        g = self.page.setdefault("geometry", {})
        for part in self._split_kv(opts):
            if "=" in part:
                k, v = part.split("=", 1)
                k, v = k.strip(), v.strip().strip("{}")
                g[k] = v
            else:
                g[part.strip()] = True

    @staticmethod
    def _split_kv(s: str) -> list[str]:
        out, depth, cur = [], 0, []
        for ch in s:
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
            if ch == "," and depth == 0:
                out.append("".join(cur))
                cur = []
            else:
                cur.append(ch)
        if "".join(cur).strip():
            out.append("".join(cur))
        return [o.strip() for o in out if o.strip()]

    # =================================================================================
    # OUTPUT ITEMS
    # =================================================================================
    def _emit_para(self, text: str, spec: dict[str, Any] | None, ctx: _Ctx) -> None:
        if getattr(self, "_config_only", False):
            return
        spec = dict(spec or {})
        for sid in reversed(ctx.spans):
            text = MK_SPAN + idx_char(sid) + text + MK_END
        if ctx.pending is not None and not ctx.pending.used:
            pt = ctx.pending.text
            ctx.pending.used = True
            spec = {**ctx.pending.spec, **spec}
            text = pt + text
        if ctx.first_para_spec:
            spec = {**ctx.first_para_spec, **spec}
            ctx.first_para_spec = None
        if ctx.align and "align" not in spec:
            spec["align"] = ctx.align
        if ctx.left and "left" not in spec:
            spec["left"] = ctx.left
        if ctx.box and "box" not in spec:
            spec["box"] = ctx.box
        if ctx.shd and "shd" not in spec:
            spec["shd"] = ctx.shd
        if not text.strip() and not any(k in spec for k in ("pb", "cols", "cols_end", "colbreak", "toc", "hr", "vspace_only")):
            return
        if getattr(self, "_pending_vspace", 0):
            spec["sb"] = max(spec.get("sb", 0), self._pending_vspace)
            self._pending_vspace = 0
        self.items.append(("para", text, spec))

    def _emit_block(self, raw: str, ctx: _Ctx, *, allow_prefix_inject: bool = False) -> None:
        if getattr(self, "_config_only", False):
            return
        if ctx.pending is not None and not ctx.pending.used:
            self._emit_para("", None, ctx)
        ctx.first_para_spec = None
        if getattr(self, "_pending_vspace", 0):
            self.items.append(("para", "", {"vspace_only": True, "sb": self._pending_vspace}))
            self._pending_vspace = 0
        if raw.lstrip().startswith(("\\choice", "\\boncot", "\\haicot", "\\motcot")) and self.items:
            kind, a, b = self.items[-1]
            if kind == "para" and a.strip():
                self.items[-1] = (kind, a, {**(b or {}), "keep": 1})
        self.items.append(("block", raw, None))

    def _assemble(self) -> str:
        parts: list[str] = []
        for kind, a, b in self.items:
            if kind == "para":
                text = a.strip()
                spec = b or {}
                marker = self.para_marker(spec) if spec else ""
                if not text and not marker:
                    continue
                parts.append(marker + text)
            else:
                parts.append(a.strip("\n"))
        return "\n\n".join(p for p in parts if p is not None)

    def _resolve_refs(self, text: str) -> str:
        def rep(m: re.Match[str]) -> str:
            kind, key = m.group(1), m.group(2)
            if kind == "c":
                keys = [k.strip() for k in key.split(",") if k.strip()]
                nums = [str(self.bib_keys.get(k, "?")) for k in keys]
                return ", ".join(nums)
            val = self.label_values.get(key)
            if val is None:
                self.res.warnings.append(f"Không tìm thấy nhãn \\label{{{key}}} cho tham chiếu.")
                return "??"
            v, typ = val
            if kind == "e":
                return f"({v})"
            if kind == "a":
                name = self._ref_type_name(typ)
                return f"{name} {v}".strip()
            return v

        return re.sub(_REF_O + r"([reac])([^" + _REF_C + r"]*)" + _REF_C, rep, text)

    def _ref_type_name(self, typ: str) -> str:
        vi = {"equation": "", "figure": "Hình", "table": "Bảng", "section": "Mục", "subsection": "Mục",
              "subsubsection": "Mục", "chapter": "Chương", "item": "", "footnote": "chú thích"}
        en = {"equation": "Equation", "figure": "Figure", "table": "Table", "section": "Section",
              "subsection": "Section", "subsubsection": "Section", "chapter": "Chapter", "item": "item"}
        if typ in self.theorems:
            return self.theorems[typ]["title"]
        return (vi if self.lang == "vi" else en).get(typ, "")

    # =================================================================================
    # BLOCK SCANNER
    # =================================================================================
    def _flush(self, buf: list[str], ctx: _Ctx, spec: dict[str, Any] | None = None) -> None:
        raw = "".join(buf)
        buf.clear()
        if not raw.strip() or getattr(self, "_config_only", False):
            return
        ictx = _ICtx(italic=ctx.italic, bold=ctx.bold)
        text = self._inline(raw, ictx).strip()
        while text.startswith("\\\\"):
            text = text[2:].strip()
        while text.endswith("\\\\"):
            text = text[:-2].strip()
        if not text:
            return
        self._emit_para(text, spec, ctx)

    def _blocks(self, s: str, ctx: _Ctx) -> None:
        buf: list[str] = []
        i = 0
        n = len(s)
        while i < n:
            c = s[i]
            if c == "\n":
                k = i + 1
                while k < n and s[k] in " \t":
                    k += 1
                if k < n and s[k] == "\n":
                    self._flush(buf, ctx)
                    while k < n and s[k] in " \t\n":
                        k += 1
                    i = k
                    continue
                buf.append(" ")
                i += 1
                continue
            if c == _VERB_BLOCK and i + 1 < n:
                self._flush(buf, ctx)
                self._emit_verbatim_block(self.verbatims[char_idx(s[i + 1])][1], ctx)
                i += 2
                continue
            if c == "$" and s.startswith("$$", i):
                end = s.find("$$", i + 2)
                if end != -1:
                    self._flush(buf, ctx)
                    self._display_math(s[i + 2 : end], ctx, numbered=False)
                    i = end + 2
                    continue
            if c == "$":
                end = self._find_inline_math_end(s, i + 1)
                buf.append(s[i : end + 1])
                i = end + 1
                continue
            if c == "{":
                inner, j = read_group(s, i)
                if BLOCK_ENV_NAMES_FOR_GROUP.search(inner):
                    self._flush(buf, ctx)
                    sub = ctx.child(pending=ctx.pending, first_para_spec=ctx.first_para_spec)
                    self._blocks(inner, sub)
                    if sub.pending is not None:
                        ctx.pending = sub.pending
                    i = j
                    continue
                buf.append(s[i:j])
                i = j
                continue
            if c == "\\":
                name, j = read_cmd(s, i)
                if name == "[":
                    end = self._find_display_end(s, j)
                    self._flush(buf, ctx)
                    self._display_math(s[j:end], ctx, numbered=False)
                    i = min(n, end + 2)
                    continue
                handler = self._BLOCK_HANDLERS.get(name)
                if handler is not None:
                    r = handler(self, s, i, j, ctx, buf)
                    if r is not None:
                        i = r
                        continue
                if name in self._CONFIG_CMDS:
                    r = self._config_cmd(name, s, j, ctx)
                    if r is not None:
                        i = r
                        continue
                if name in ("bfseries", "bf", "itshape", "it", "centering", "raggedleft", "raggedright",
                            "small", "footnotesize", "scriptsize", "tiny", "large", "Large", "LARGE", "huge",
                            "Huge", "normalsize", "color", "sffamily", "ttfamily", "scshape") and not buf_has_text(buf):
                    r = self._block_declaration(name, s, i, j, ctx)
                    if r is not None:
                        i = r
                        continue
                r = self._try_choice_like(name, s, i, j, ctx, buf)
                if r is not None:
                    i = r
                    continue
                # read the whole inline command unit (name + args) so its groups stay together
                k = self._skip_inline_unit(s, i, j, name)
                unit = s[i:k]
                if self._unit_has_block(unit, name):
                    self._flush(buf, ctx)
                    self._block_from_inline_unit(name, s, i, j, k, ctx)
                    i = k
                    continue
                buf.append(unit)
                i = k
                continue
            buf.append(c)
            i += 1
        self._flush(buf, ctx)

    def _find_inline_math_end(self, s: str, i: int) -> int:
        n = len(s)
        k = i
        depth = 0
        while k < n:
            ch = s[k]
            if ch == "\\":
                k += 2
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth = max(0, depth - 1)
            elif ch == "$":
                return k
            elif ch == "\n" and s[k + 1 : k + 2] == "\n" and depth == 0:
                return k  # unterminated math — stop at paragraph end
            k += 1
        return n

    @staticmethod
    def _find_display_end(s: str, j: int) -> int:
        k = j
        n = len(s)
        while k < n:
            p = s.find("\\]", k)
            if p == -1:
                return n
            bs = 0
            q = p - 1
            while q >= j and s[q] == "\\":
                bs += 1
                q -= 1
            if bs % 2 == 0:
                return p
            k = p + 2
        return n

    INLINE_ARG_SPECS: dict[str, str] = {
        "textbf": "m", "textit": "m", "emph": "m", "underline": "m", "texttt": "m", "textsf": "m",
        "textrm": "m", "textsc": "m", "textsl": "m", "textup": "m", "textmd": "m", "textnormal": "m",
        "textcolor": "om m", "colorbox": "om m", "fcolorbox": "om m m", "fbox": "m", "framebox": "oo m",
        "boxed": "m", "mbox": "m", "hbox": "m", "makebox": "oo m", "parbox": "ooo m m", "raisebox": "m oo m",
        "scalebox": "m o m", "resizebox": "s m m m", "rotatebox": "o m m", "reflectbox": "m",
        "adjustbox": "m m", "centerline": "m", "rightline": "m", "leftline": "m", "hl": "m", "sout": "m",
        "uline": "m", "uuline": "m", "uwave": "m", "xout": "m", "st": "m", "ul": "m", "dashuline": "m",
        "dotuline": "m", "footnote": "o m", "href": "m m", "url": "m", "textsuperscript": "m",
        "textsubscript": "m", "highlight": "m", "underbar": "m", "colorbox*": "m m", "tcbox": "o m",
        "shadowbox": "m", "doublebox": "m", "ovalbox": "m", "Ovalbox": "m", "cfbox": "m m",
    }

    def _skip_inline_unit(self, s: str, i: int, j: int, name: str) -> int:
        spec = self.INLINE_ARG_SPECS.get(name)
        n = len(s)
        if spec is not None:
            k = j
            for t in spec.split():
                for ch in t:
                    if ch == "s":
                        _, k = read_star(s, k)
                    elif ch == "o":
                        _, k = read_opt(s, k, allow_newline=False)
                    elif ch == "m":
                        a, k2 = read_arg(s, k)
                        if a is None:
                            return k
                        k = k2
            return k
        if not name.isalpha():
            return j
        # unknown/other: swallow star, optional and braced args (TeX-like greedy)
        k = j
        st, k = read_star(s, k)
        while True:
            kk = skip_spaces_only(s, k)
            if kk < n and s[kk] == "[" and name not in ("\\",):
                _, k = read_bracket(s, kk)
                continue
            if kk < n and s[kk] == "{":
                _, k = read_group(s, kk)
                continue
            break
        return k

    def _unit_has_block(self, unit: str, name: str) -> bool:
        if name in ("textbf", "textit", "emph", "underline", "textcolor", "colorbox", "fcolorbox", "fbox",
                    "framebox", "mbox", "makebox", "parbox", "raisebox", "scalebox", "resizebox", "rotatebox",
                    "adjustbox", "centerline", "rightline", "leftline", "hbox", "boxed", "tcbox", "shadowbox",
                    "doublebox", "ovalbox", "Ovalbox", "cfbox", "texttt", "textsf", "reflectbox"):
            return bool(BLOCK_ENV_NAMES_FOR_GROUP.search(unit[len(name) + 1 :])) or name in ("centerline", "rightline", "leftline")
        if name in self.INLINE_ARG_SPECS or name in TEXT_SYMBOLS or name in TEXT_ACCENTS:
            return False
        if name in self._INLINE_KNOWN:
            return False
        rest = unit[len(name) + 1 :]
        return bool(rest) and bool(BLOCK_ENV_NAMES_FOR_GROUP.search(rest))

    def _block_from_inline_unit(self, name: str, s: str, i: int, j: int, k: int, ctx: _Ctx) -> None:
        """A formatting/wrapper command whose argument contains block material."""
        args: list[str] = []
        opts: list[str] = []
        p = j
        _, p = read_star(s, p)
        while p < k:
            pp = skip_ws(s, p)
            if pp >= k:
                break
            if s[pp] == "[":
                o, p = read_bracket(s, pp)
                opts.append(o)
            elif s[pp] == "{":
                a, p = read_group(s, pp)
                args.append(a)
            else:
                break
        if not args:
            return
        known_wrappers = {"textbf", "textit", "emph", "underline", "textcolor", "colorbox", "fcolorbox", "fbox",
                          "framebox", "mbox", "makebox", "parbox", "raisebox", "scalebox", "resizebox", "rotatebox",
                          "adjustbox", "centerline", "rightline", "leftline", "hbox", "boxed", "tcbox", "shadowbox",
                          "doublebox", "ovalbox", "Ovalbox", "cfbox", "texttt", "textsf", "reflectbox"}
        if name not in known_wrappers:
            self._unknown(name)
            for a in args:
                if re.fullmatch(r"\s*-?[\d.]*\s*(cm|mm|pt|em|ex|in)?\s*(\\(textwidth|linewidth))?\s*", a):
                    continue
                self._blocks(a, ctx.child(pending=ctx.pending))
            return
        content = args[-1]
        sub = ctx.child(pending=ctx.pending)
        span_spec: dict[str, Any] = {}
        if name == "textbf":
            span_spec = {"b": 1}
        elif name in ("textit", "emph"):
            span_spec = {"i": 1}
        elif name == "underline":
            span_spec = {"u": "single"}
        elif name == "textcolor" and len(args) >= 2:
            hx = self.color_hex(args[0], opts[0] if opts else None)
            if hx:
                span_spec = {"c": hx}
        elif name in ("colorbox",) and len(args) >= 2:
            hx = self.color_hex(args[0], opts[0] if opts else None)
            if hx:
                sub.shd = hx
        elif name in ("fbox", "framebox", "boxed", "tcbox", "shadowbox", "doublebox", "ovalbox", "Ovalbox", "cfbox", "fcolorbox"):
            sub.box = {"style": "single"}
        elif name == "centerline":
            sub.align = "c"
        elif name == "rightline":
            sub.align = "r"
        elif name == "leftline":
            sub.align = "l"
        if span_spec:
            sub.spans = ctx.spans + (self.span_id(span_spec),)
        self._blocks(content, sub)

    def _block_declaration(self, name: str, s: str, i: int, j: int, ctx: _Ctx) -> int | None:
        """Declarations at the start of a block scope affect all following paragraphs of that scope."""
        if name == "centering":
            ctx.align = "c"
            return j
        if name == "raggedleft":
            ctx.align = "r"
            return j
        if name == "raggedright":
            ctx.align = "l"
            return j
        if name in ("bfseries", "bf"):
            ctx.bold = True
            ctx.spans = ctx.spans + (self.span_id({"b": 1}),)
            return j
        if name in ("itshape", "it"):
            ctx.spans = ctx.spans + (self.span_id({"i": 1}),)
            return j
        if name in SIZE_SWITCH_PT:
            ctx.spans = ctx.spans + (self.span_id({"sz": SIZE_SWITCH_PT[name]}),)
            return j
        if name == "sffamily":
            ctx.spans = ctx.spans + (self.span_id({"f": "Arial"}),)
            return j
        if name == "ttfamily":
            ctx.spans = ctx.spans + (self.span_id({"f": "Courier New"}),)
            return j
        if name == "scshape":
            ctx.spans = ctx.spans + (self.span_id({"sc": 1}),)
            return j
        if name == "color":
            o, k = read_opt(s, j, allow_newline=False)
            a, k2 = read_braced(s, k)
            if a is None:
                return None
            hx = self.color_hex(a, o)
            if hx:
                ctx.spans = ctx.spans + (self.span_id({"c": hx}),)
            return k2
        return None

    # ---------------- verbatim ----------------
    def _emit_verbatim_block(self, body: str, ctx: _Ctx) -> None:
        lines = body.split("\n")
        rendered = []
        for ln in lines:
            ln = ln.replace("\t", "    ")
            lead = len(ln) - len(ln.lstrip(" "))
            rendered.append("\u00a0" * lead + lit(ln.lstrip(" ")).replace(" ", "\u00a0"))
        text = " \\\\ ".join(rendered) if rendered else ""
        text = self.span({"f": "Courier New", "sz": 10.5}, text)
        self._emit_para(text, {"left": ctx.left + 240, "sb": 4, "sa": 4, "keep": 1}, ctx)
        self._stat("verbatim_blocks")

    # ---------------- display math ----------------
    def _display_math(self, body: str, ctx: _Ctx, *, numbered: bool, env: str = "") -> None:
        body = body.strip()
        if not body:
            return
        if getattr(self, "_config_only", False):
            return
        rows_env = env.rstrip("*") in ("align", "flalign", "alignat", "eqnarray", "gather", "multline",
                                       "xalignat", "xxalignat", "IEEEeqnarray", "dgroup")
        if env.startswith("alignat") or env.startswith("xalignat") or env.startswith("IEEEeqnarray"):
            body = re.sub(r"^\s*\{[^{}]*\}", "", body, count=1)
        if env.startswith("eqnarray") or env.startswith("IEEEeqnarray"):
            body = body.replace("&=&", "&=").replace("& = &", "&=")
        # labels & tags per row
        if rows_env:
            rows = self._split_rows(body)
            out_rows = []
            for r in rows:
                if not r.strip():
                    continue
                has_nonum = bool(re.search(r"\\(nonumber|notag)\b", r))
                mt = re.search(r"\\tag\*?\{([^{}]*)\}", r)
                labels = re.findall(r"\\label\{([^{}]*)\}", r)
                r = re.sub(r"\\(nonumber|notag)\b|\\label\{[^{}]*\}|\\tag\*?\{[^{}]*\}", "", r)
                tag = None
                if mt:
                    tag = mt.group(1)
                    self.current_label = (tag, "equation")
                elif numbered and not has_nonum:
                    tag = self.refstep("equation", "equation")
                for lb in labels:
                    self.label_values[lb] = self.current_label
                m_r = self._math(r)
                if tag is not None:
                    m_r = m_r + r" & \qquad (" + tag + ")"
                out_rows.append(m_r)
            if not out_rows:
                return
            nested_sep = any(_has_nested_separators(r) for r in out_rows)
            if len(out_rows) == 1 or nested_sep:
                # the engine splits aligned rows/cols naively → render each row as its own display
                for r in out_rows:
                    core = r
                    mtag = re.search(r" & \\qquad \((.*)\)$", core)
                    if mtag:
                        core = core[: mtag.start()] + r"\tag{" + mtag.group(1) + "}"
                    if not _top_level_amp(core) or nested_sep:
                        core = _drop_top_level_amp(core)
                    else:
                        core = r"\begin{aligned}" + core + r"\end{aligned}"
                    self._emit_block(r"\[" + core + r"\]", ctx)
                self._stat("display_equations")
                return
            math = r"\begin{aligned}" + r" \\ ".join(out_rows) + r"\end{aligned}"
        else:
            labels = re.findall(r"\\label\{([^{}]*)\}", body)
            mt = re.search(r"\\tag\*?\{([^{}]*)\}", body)
            has_nonum = bool(re.search(r"\\(nonumber|notag)\b", body))
            b2 = re.sub(r"\\(nonumber|notag)\b|\\label\{[^{}]*\}|\\tag\*?\{[^{}]*\}", "", body)
            tag = None
            if mt:
                tag = mt.group(1)
                self.current_label = (tag, "equation")
            elif numbered and not has_nonum:
                tag = self.refstep("equation", "equation")
            for lb in labels:
                self.label_values[lb] = self.current_label
            math = self._math(b2)
            if re.search(r"(?<!\\)\\\\", math) and not re.search(r"\\begin\{", math):
                math = r"\begin{aligned}" + math + r"\end{aligned}"
            if tag is not None:
                math = math + r"\tag{" + tag + "}"
        if not math.strip():
            return
        if ctx.in_immini:
            self._emit_block(r"\[" + math + r"\]", ctx)
        else:
            self._emit_block(r"\[" + math + r"\]", ctx)
        self._stat("display_equations")

    @staticmethod
    def _split_rows(body: str) -> list[str]:
        rows: list[str] = []
        depth = 0
        env_depth = 0
        cur: list[str] = []
        i = 0
        n = len(body)
        while i < n:
            if body.startswith("\\begin{", i):
                env_depth += 1
            elif body.startswith("\\end{", i):
                env_depth -= 1
            c = body[i]
            if c == "\\" and body.startswith("\\\\", i) and depth == 0 and env_depth == 0:
                rows.append("".join(cur))
                cur = []
                i += 2
                # skip optional [dim] / *
                if i < n and body[i] == "*":
                    i += 1
                k = skip_spaces_only(body, i)
                if k < n and body[k] == "[":
                    _, i = read_bracket(body, k)
                continue
            if c == "\\":
                cur.append(body[i : i + 2])
                i += 2
                continue
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
            cur.append(c)
            i += 1
        rows.append("".join(cur))
        return rows


PAPER_SIZES_TWIPS = {
    "a4paper": (11906, 16838), "a5paper": (8391, 11906), "a3paper": (16838, 23811), "b5paper": (9979, 14175),
    "letterpaper": (12240, 15840), "legalpaper": (12240, 20160), "executivepaper": (10440, 15120),
}


def compute_page_geometry(page: dict[str, Any]) -> dict[str, int]:
    """Page size & margins (twips). Engine default: A4, 2 cm margins; geometry options override."""
    w, h = PAPER_SIZES_TWIPS.get(page.get("paper") or "a4paper", PAPER_SIZES_TWIPS["a4paper"])
    m = {"left": 1134, "right": 1134, "top": 1134, "bottom": 1134}
    g = page.get("geometry") or {}
    for key in g:
        if key in PAPER_SIZES_TWIPS:
            w, h = PAPER_SIZES_TWIPS[key]
        if key in ("a4", "a5", "letter", "b5"):
            w, h = PAPER_SIZES_TWIPS.get(key + "paper", (w, h))
    if g.get("paper") in PAPER_SIZES_TWIPS:
        w, h = PAPER_SIZES_TWIPS[g["paper"]]
    if isinstance(g.get("papersize"), str):
        parts = [p for p in re.split(r"[,{}]", g["papersize"]) if p.strip()]
        if len(parts) == 2:
            w = dim_to_twips(parts[0]) or w
            h = dim_to_twips(parts[1]) or h
    if isinstance(g.get("paperwidth"), str):
        w = dim_to_twips(g["paperwidth"]) or w
    if isinstance(g.get("paperheight"), str):
        h = dim_to_twips(g["paperheight"]) or h
    if g.get("landscape") is True or page.get("landscape"):
        w, h = max(w, h), min(w, h)

    def d(v: Any) -> int | None:
        return dim_to_twips(v) if isinstance(v, str) else None

    for key, sides in (("margin", ("left", "right", "top", "bottom")), ("hmargin", ("left", "right")),
                       ("vmargin", ("top", "bottom")), ("lmargin", ("left",)), ("rmargin", ("right",)),
                       ("tmargin", ("top",)), ("bmargin", ("bottom",)), ("left", ("left",)), ("inner", ("left",)),
                       ("right", ("right",)), ("outer", ("right",)), ("top", ("top",)), ("bottom", ("bottom",))):
        val = g.get(key)
        if isinstance(val, str) and "," in val and key in ("hmargin", "vmargin"):
            a, b = val.split(",", 1)
            va, vb = d(a), d(b)
            if va is not None:
                m[sides[0]] = va
            if vb is not None:
                m[sides[1]] = vb
            continue
        v = d(val)
        if v is not None:
            for sd in sides:
                m[sd] = v
    tw = d(g.get("textwidth")) or d(g.get("width"))
    if tw:
        side = max(0, (w - tw) // 2)
        if "left" not in g and "lmargin" not in g:
            m["left"] = side
        if "right" not in g and "rmargin" not in g:
            m["right"] = side
    th = d(g.get("textheight")) or d(g.get("height"))
    if th:
        side = max(0, (h - th) // 2)
        if "top" not in g:
            m["top"] = side
        if "bottom" not in g:
            m["bottom"] = side
    return {"width": int(w), "height": int(h), **{k: int(v) for k, v in m.items()}}


def _top_level_amp(m: str) -> bool:
    depth = 0
    env = 0
    i = 0
    while i < len(m):
        if m.startswith("\\begin{", i):
            env += 1
        elif m.startswith("\\end{", i):
            env -= 1
        c = m[i]
        if c == "\\":
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        elif c == "&" and depth == 0 and env == 0:
            return True
        i += 1
    return False


def _drop_top_level_amp(m: str) -> str:
    out = []
    depth = 0
    env = 0
    i = 0
    while i < len(m):
        if m.startswith("\\begin{", i):
            env += 1
        elif m.startswith("\\end{", i):
            env -= 1
        c = m[i]
        if c == "\\":
            out.append(m[i : i + 2])
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        if c == "&" and depth == 0 and env == 0:
            i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _has_nested_separators(m: str) -> bool:
    """True if an inner environment / \\heva uses & or \\\\ (engine's aligned splitter would cut it)."""
    for mm in re.finditer(r"\\begin\{([a-zA-Z*]+)\}", m):
        env = mm.group(1)
        end = m.find("\\end{" + env + "}", mm.end())
        inner = m[mm.end(): end if end != -1 else len(m)]
        if "&" in inner or "\\\\" in inner:
            return True
    for mm in re.finditer(r"\\(heva|hoac)\s*\{", m):
        return True
    return False


def buf_has_text(buf: list[str]) -> bool:
    return any(x.strip() for x in buf)


from .latex_preprocessor_parts import install_parts  # noqa: E402  (split for readability)

install_parts(LatexPreprocessor)


def preprocess_latex(src: str, *, base_dir: Path | None = None, show_solutions: bool = True) -> PreprocessResult:
    """Public API — convert any pasted LaTeX into engine-ready LaTeX + layout markers."""
    return LatexPreprocessor(base_dir=base_dir, show_solutions=show_solutions).run(src)
