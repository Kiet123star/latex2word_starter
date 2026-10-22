"""`backend/macro_fidelity.py` — TRẢ LẠI ĐÚNG MACRO NGƯỜI DÙNG ĐÃ ĐỊNH NGHĨA.

Vấn đề: engine gốc có khuôn dựng sẵn cho một số lệnh đề thi (`\\choice`, `\\True`,
`\\bonpa`…). Khi người dùng **tự định nghĩa lại** những lệnh đó, engine vẫn dùng khuôn
của nó và bỏ qua định nghĩa của người dùng — báo cáo ghi rõ:

    "Giữ cấu trúc chuẩn cho \\choice (bỏ qua định nghĩa lại của người dùng)."

Hậu quả: bản gốc xếp đáp án **2 cột, đáp án đúng in đậm + gạch chân**, còn Word ra
1x4 / 4x1 tuỳ độ dài, không in đậm. Sai bố cục trên toàn bộ tài liệu.

Module này **nở macro của người dùng ngay trong mã nguồn** trước khi đưa cho engine, nên
engine không còn thấy `\\choice` để mà đè lên — nó chỉ thấy một cái `tabular` bình thường
và dựng thành bảng Word sửa được, **đúng bố cục người dùng viết**.

Bộ nở hiểu đúng thành ngữ `ex_test` kinh điển:

    \\newif\\ifcorrect
    \\newcommand{\\True}{\\global\\correcttrue}
    \\newcommand{\\choice}[4]{… \\global\\correctfalse \\sbox0{#1}
                             \\ifcorrect \\underline{\\textbf{A.}} \\textbf{#1} \\else A. #1 \\fi …}

Nguyên tắc an toàn: gặp bất kỳ cấu trúc nào chưa chắc chắn → **bỏ nở, giữ nguyên mã gốc**,
pipeline cũ chạy y như trước.

Tắt: `L2W_MACRO_FIDELITY=off`.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any

__all__ = ["rewrite_user_macros", "ENGINE_OVERRIDDEN", "MacroFidelityResult"]

#: Những lệnh mà engine gốc có khuôn dựng sẵn và sẽ ĐÈ lên định nghĩa của người dùng.
ENGINE_OVERRIDDEN: set[str] = {
    "choice", "choiceA", "choiceB", "choiceC", "choiceD", "choiceTF", "choicef",
    "bonpa", "boncot", "haicot", "motcot", "cau", "Cau", "True", "immini",
    "loigiai", "shortans", "dapan", "khoanh",
}

MAX_EXPANSIONS = 2000


@dataclass
class MacroDef:
    name: str
    nargs: int
    body: str
    has_optional: bool = False
    optional_default: str = ""


@dataclass
class MacroFidelityResult:
    text: str
    expanded: dict[str, int] = field(default_factory=dict)
    skipped: dict[str, str] = field(default_factory=dict)
    flags: list[str] = field(default_factory=list)
    setters: list[str] = field(default_factory=list)

    def summary(self) -> dict[str, Any]:
        return {
            "macro_da_no": self.expanded,
            "macro_bo_qua": self.skipped,
            "co_dieu_kien": self.flags,
            "lenh_danh_dau": self.setters,
        }


def _mode() -> str:
    """`auto` (mặc định) | `all` | `off`.

    * `auto` — chỉ nở macro có **cờ điều kiện** (thành ngữ `\\True` đánh dấu đáp án đúng).
      Đây đúng là chỗ engine gốc dựng sai, còn những macro khác engine vốn đã dựng tốt
      nên không đụng vào → không gây thay đổi ngoài ý muốn ở các file đang chạy ổn.
    * `all`  — nở mọi macro người dùng định nghĩa lại mà engine có khuôn riêng.
    * `off`  — không nở gì cả."""
    v = os.environ.get("L2W_MACRO_FIDELITY", "auto").strip().lower()
    return v if v in ("auto", "all", "off") else ("off" if v in ("0", "no", "false") else "auto")


def is_enabled() -> bool:
    return _mode() != "off"


# ─────────────────────────── đọc định nghĩa ───────────────────────────
def _match_group(s: str, i: int) -> tuple[str, int]:
    """`i` trỏ vào `{` → trả về (nội dung bên trong, vị trí sau `}`)."""
    if i >= len(s) or s[i] != "{":
        return "", i
    depth, j = 0, i
    while j < len(s):
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return "", i


_NEWCMD_RE = re.compile(r"\\(?:new|renew|provide)command\s*\*?\s*(?:\{\s*\\([A-Za-z@]+)\s*\}|\\([A-Za-z@]+))")
_NEWIF_RE = re.compile(r"\\newif\s*\\if([A-Za-z@]+)")


def parse_definitions(src: str) -> tuple[dict[str, MacroDef], set[str], dict[str, tuple[str, bool]]]:
    """Trả về (macro, tên cờ điều kiện, lệnh đánh dấu → (cờ, giá trị))."""
    macros: dict[str, MacroDef] = {}
    flags: set[str] = set(_NEWIF_RE.findall(src))
    for m in _NEWCMD_RE.finditer(src):
        name = m.group(1) or m.group(2)
        i = m.end()
        nargs, has_opt, opt_default = 0, False, ""
        while i < len(src) and src[i] in " \t\n":
            i += 1
        if i < len(src) and src[i] == "[":
            j = src.find("]", i)
            if j > 0:
                try:
                    nargs = int(src[i + 1:j].strip())
                except ValueError:
                    nargs = 0
                i = j + 1
                while i < len(src) and src[i] in " \t\n":
                    i += 1
                if i < len(src) and src[i] == "[":
                    opt_default, i = _match_group("{" + src[i + 1:], 0)
                    k = src.find("]", i)
                    i = (k + 1) if k > 0 else i
                    has_opt = True
        while i < len(src) and src[i] in " \t\n":
            i += 1
        body, _ = _match_group(src, i)
        if body or nargs:
            macros[name] = MacroDef(name, nargs, body, has_opt, opt_default)
    setters: dict[str, tuple[str, bool]] = {}
    for name, d in macros.items():
        if d.nargs:
            continue
        hit = re.fullmatch(r"\s*(?:\\global\s*)?\\([A-Za-z@]+)(true|false)\s*", d.body)
        if hit and hit.group(1) in flags:
            setters[name] = (hit.group(1), hit.group(2) == "true")
    return macros, flags, setters


# ─────────────────────────── nở một lần gọi ───────────────────────────
class _Unsupported(Exception):
    pass


_SETTER_IN_TEXT = re.compile(r"\\([A-Za-z@]+)")


def _apply_setters(text: str, flags: dict[str, bool], setters: dict[str, tuple[str, bool]]) -> None:
    for m in _SETTER_IN_TEXT.finditer(text):
        nm = m.group(1)
        if nm in setters:
            flag, val = setters[nm]
            flags[flag] = val
        elif nm.endswith("true") and nm[:-4] in flags:
            flags[nm[:-4]] = True
        elif nm.endswith("false") and nm[:-5] in flags:
            flags[nm[:-5]] = False


def _strip_setters(text: str, setters: dict[str, tuple[str, bool]]) -> str:
    if not setters:
        return text
    pat = "|".join(re.escape(n) for n in sorted(setters, key=len, reverse=True))
    return re.sub(r"\\(?:" + pat + r")(?![A-Za-z])\s*", "", text)


def _evaluate(body: str, flags: dict[str, bool], setters: dict[str, tuple[str, bool]]) -> str:
    """Chạy phần thân macro: xử lý cờ, `\\sbox`, `\\ifFLAG…\\else…\\fi`; phần khác giữ nguyên."""
    out: list[str] = []
    i, n = 0, len(body)
    guard = 0
    while i < n:
        guard += 1
        if guard > 200000:
            raise _Unsupported("thân macro quá phức tạp")
        c = body[i]
        if c != "\\":
            out.append(c)
            i += 1
            continue
        m = re.match(r"\\([A-Za-z@]+)", body[i:])
        if not m:
            out.append(body[i:i + 2])
            i += 2
            continue
        name = m.group(1)
        j = i + m.end()

        if name == "global":
            i = j
            continue
        if name in setters:
            flag, val = setters[name]
            flags[flag] = val
            i = j
            continue
        if name.endswith("true") and name[:-4] in flags:
            flags[name[:-4]] = True
            i = j
            continue
        if name.endswith("false") and name[:-5] in flags:
            flags[name[:-5]] = False
            i = j
            continue
        if name in ("sbox", "setbox", "savebox"):
            k = j
            groups = 0
            while k < n and groups < 3:
                while k < n and body[k] in " \t\n":
                    k += 1
                if k < n and body[k] == "0":          # \sbox0{...}
                    k += 1
                    continue
                if k < n and body[k] == "\\":         # \setbox\mybox
                    k += 1 + (re.match(r"[A-Za-z@]+", body[k + 1:]) or re.match(r"", "")).end()
                    continue
                if k < n and body[k] == "{":
                    inner, k = _match_group(body, k)
                    _apply_setters(inner, flags, setters)
                    groups += 1
                    if groups >= 1:
                        break
                    continue
                break
            i = k
            continue
        if name.startswith("if"):
            flag = name[2:]
            if flag not in flags:
                raise _Unsupported(f"điều kiện chưa hỗ trợ: \\{name}")
            then_part, else_part, k = _split_conditional(body, j, flags)
            chosen = then_part if flags[flag] else else_part
            out.append(_evaluate(chosen, flags, setters))
            i = k
            continue
        if name in ("else", "fi"):
            raise _Unsupported(f"\\{name} lạc chỗ")
        out.append(body[i:j])
        i = j
    return "".join(out)


def _split_conditional(body: str, start: int, flags: dict[str, bool]) -> tuple[str, str, int]:
    """Từ ngay sau `\\ifFLAG` → trả về (nhánh đúng, nhánh sai, vị trí sau `\\fi`)."""
    tok_re = re.compile(r"\\(if[A-Za-z@]*|else|fi)(?![A-Za-z])")
    depth = 0
    else_at = -1
    i = start
    while True:
        m = tok_re.search(body, i)
        if not m:
            raise _Unsupported("thiếu \\fi")
        tok = m.group(1)
        if tok.startswith("if"):
            depth += 1
        elif tok == "else":
            if depth == 0 and else_at < 0:
                else_at = m.start()
        else:                                  # \fi
            if depth == 0:
                if else_at >= 0:
                    then_part = body[start:else_at]
                    else_part = body[else_at + len("\\else"):m.start()]
                else:
                    then_part = body[start:m.start()]
                    else_part = ""
                return then_part, else_part, m.end()
            depth -= 1
        i = m.end()


# ─────────────────────────── viết lại mã nguồn ───────────────────────────
def _grab_args(src: str, i: int, nargs: int, has_opt: bool, opt_default: str) -> tuple[list[str], int] | None:
    args: list[str] = []
    if has_opt:
        k = i
        while k < len(src) and src[k] in " \t\n":
            k += 1
        if k < len(src) and src[k] == "[":
            j = src.find("]", k)
            if j < 0:
                return None
            args.append(src[k + 1:j])
            i = j + 1
        else:
            args.append(opt_default)
        nargs -= 1
    for _ in range(nargs):
        while i < len(src) and src[i] in " \t\n":
            i += 1
        if i >= len(src):
            return None
        if src[i] == "{":
            inner, i = _match_group(src, i)
            args.append(inner)
        elif src[i] == "\\":
            m = re.match(r"\\[A-Za-z@]+|\\.", src[i:])
            if not m:
                return None
            args.append(src[i:i + m.end()])
            i += m.end()
        else:
            args.append(src[i])
            i += 1
    return args, i


_TABULARX_BEGIN = re.compile(r"\\begin\{tabularx\}\s*")


def _soften_tabularx(text: str) -> str:
    """`tabularx` cột `X` → `tabular` cột `p{…}`: engine gốc chắc chắn dựng được,
    bề rộng cột giữ nguyên nên nhìn không khác bản gốc."""
    if os.environ.get("L2W_SOFTEN_TABULARX", "1").strip().lower() in ("0", "off", "no", "false"):
        return text
    out = text
    while True:
        m = _TABULARX_BEGIN.search(out)
        if not m:
            break
        i = m.end()
        if i >= len(out) or out[i] != "{":
            break
        _width, i = _match_group(out, i)
        if i >= len(out) or out[i] != "{":
            break
        spec, i = _match_group(out, i)
        ncols = spec.count("X")
        if not ncols:
            break
        w = 0.94 / ncols
        new_spec = spec.replace("X", "p{%.3f\\textwidth}" % w)
        out = out[:m.start()] + "\\begin{tabular}{" + new_spec + "}" + out[i:]
        out = out.replace("\\end{tabularx}", "\\end{tabular}", 1)
    return out


_COMMENT_RE = re.compile(r"(?<!\\)%[^\n]*")


def _strip_comments(text: str) -> str:
    """Bỏ chú thích LaTeX lọt ra từ thân macro, tránh engine in nhầm ra Word."""
    out = []
    for line in text.split("\n"):
        i, n, esc = 0, len(line), False
        cut = None
        while i < n:
            c = line[i]
            if c == "\\":
                i += 2
                continue
            if c == "%":
                cut = i
                break
            i += 1
        out.append(line if cut is None else line[:cut])
    return "\n".join(out)


_TABSTAR_RE = re.compile(r"\\begin\{tabular\*\}\s*")


def _soften_tabular_star(text: str) -> str:
    """`tabular*{\textwidth}{spec}` → `tabular{spec}`: engine gốc dựng chắc tay hơn,
    bề rộng cột đã nằm trong spec nên nhìn không khác."""
    out = text
    while True:
        m = _TABSTAR_RE.search(out)
        if not m:
            break
        i = m.end()
        if i >= len(out) or out[i] != "{":
            break
        _w, i = _match_group(out, i)
        if i >= len(out) or out[i] != "{":
            break
        spec, i = _match_group(out, i)
        out = out[:m.start()] + "\\begin{tabular}{" + spec + "}" + out[i:]
        out = out.replace("\\end{tabular*}", "\\end{tabular}", 1)
    return out


_ROWEND_OPT_RE = re.compile(r"\\\\\s*\[[^\]]*\]")
_TAB_BEGIN_RE = re.compile(r"\\begin\{(tabular\*?|tabularx|array)\}")


def _tidy(text: str) -> str:
    r"""Dọn khoảng trắng sau khi nở macro.

    RẤT QUAN TRỌNG: thân macro của người dùng thường có xuống dòng giữa các ô; sau khi bỏ
    `\global\correctfalse \sbox0{#1}` thì chỗ đó thành **dòng trống** → LaTeX hiểu là
    `\par` và engine cắt nhầm ô bảng (đáp án rơi ra ngoài bảng). Gộp hết về một dòng,
    chỉ xuống dòng sau mỗi dấu kết thúc hàng `\\`."""
    t = _ROWEND_OPT_RE.sub(lambda _m: "\\\\", text)              # \\[1.2em] → \\
    t = re.sub(r"[ \t]*\r?\n[ \t\r\n]*", " ", t)       # bỏ mọi dòng trống
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = t.replace("\\\\", "\\\\\n")                    # xuống dòng sau mỗi hàng bảng
    # xuống dòng sau \begin{tabular}{...}{...} và trước \end{...}
    out, i = [], 0
    while True:
        m = _TAB_BEGIN_RE.search(t, i)
        if not m:
            out.append(t[i:])
            break
        j = m.end()
        while j < len(t) and t[j] in " \t":
            j += 1
        while j < len(t) and t[j] == "{":
            _g, j = _match_group(t, j)
            while j < len(t) and t[j] in " \t":
                j += 1
        out.append(t[i:j])
        out.append("\n")
        i = j
    t = "".join(out)
    t = re.sub(r"\s*(\\end\{(?:tabular\*?|tabularx|array)\})", r"\n\1", t)
    t = re.sub(r"\n{2,}", "\n", t)
    return t.strip()


def rewrite_user_macros(src: str, targets: set[str] | None = None) -> MacroFidelityResult:
    """Nở các macro người dùng tự định nghĩa mà engine gốc sẽ đè lên.

    Không bao giờ ném lỗi: chỗ nào không nở được thì giữ nguyên."""
    res = MacroFidelityResult(text=src)
    if not is_enabled():
        res.skipped["*"] = "tắt bằng L2W_MACRO_FIDELITY=off"
        return res
    try:
        macros, flag_names, setters = parse_definitions(src)
    except Exception as exc:
        res.skipped["*"] = f"không đọc được định nghĩa: {exc}"
        return res
    res.flags = sorted(flag_names)
    res.setters = sorted(setters)
    want = (targets if targets is not None else ENGINE_OVERRIDDEN)
    todo = {n: d for n, d in macros.items() if n in want and d.nargs > 0}
    if _mode() == "auto":
        cond = {n: d for n, d in todo.items()
                if any(re.search(r"\\if" + re.escape(f) + r"(?![A-Za-z])", d.body) for f in flag_names)}
        for n in set(todo) - set(cond):
            res.skipped[n] = "engine gốc vốn dựng tốt macro này (đặt L2W_MACRO_FIDELITY=all để vẫn nở)"
        todo = cond
    if not todo:
        return res

    body_start = 0
    m = re.search(r"\\begin\s*\{document\}", src)
    if m:
        body_start = m.end()
    head, body = src[:body_start], src[body_start:]

    names = sorted(todo, key=len, reverse=True)
    call_re = re.compile(r"\\(" + "|".join(re.escape(n) for n in names) + r")(?![A-Za-z@])")
    out: list[str] = []
    pos = 0
    count = 0
    while True:
        m = call_re.search(body, pos)
        if m is None or count >= MAX_EXPANSIONS:
            break
        d = todo[m.group(1)]
        got = _grab_args(body, m.end(), d.nargs, d.has_optional, d.optional_default)
        if got is None:
            out.append(body[pos:m.end()])
            pos = m.end()
            continue
        args, after = got
        try:
            filled = d.body
            for k in range(len(args), 0, -1):
                filled = filled.replace(f"#{k}", args[k - 1])
            flags = {f: False for f in flag_names}
            expanded = _evaluate(filled, flags, setters)
            expanded = _strip_setters(expanded, setters)
            expanded = _soften_tabularx(expanded)
            expanded = _soften_tabular_star(expanded)
            expanded = _strip_comments(expanded)
            expanded = _tidy(expanded)
        except _Unsupported as exc:
            res.skipped[d.name] = str(exc)
            out.append(body[pos:m.end()])
            pos = m.end()
            continue
        except Exception as exc:                 # an toàn tuyệt đối
            res.skipped[d.name] = f"{type(exc).__name__}: {exc}"
            out.append(body[pos:m.end()])
            pos = m.end()
            continue
        out.append(body[pos:m.start()])
        out.append("\n" + expanded.strip() + "\n")
        pos = after
        count += 1
        res.expanded[d.name] = res.expanded.get(d.name, 0) + 1
    out.append(body[pos:])
    res.text = head + "".join(out)
    return res
