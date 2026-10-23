"""`backend/universal_fallback.py` — LỚP BẢO HIỂM PHỦ 100% cho nhánh **dán mã LaTeX**.

Vấn đề cũ: mỗi khi gặp một cấu trúc LaTeX chưa được lập trình sẵn, engine đoán sai và
Word ra sai (mất nét vẽ, mất nhãn, in nguyên văn `\\tenlenh`…).  Càng thêm cấu trúc thì
càng còn cấu trúc khác chưa thêm — không bao giờ hết.

Cách giải quyết ở đây **đảo ngược bài toán**: thay vì cố đoán mọi cấu trúc, ta chỉ cần
biết *cấu trúc nào chắc chắn đã dựng được*.  Phần còn lại — bất kể là gì — được giao cho
**chính TeX** dựng rồi nhúng ảnh vector vào Word (`backend/tex_truth.py`).  Nhờ vậy:

* cấu trúc đã biết  → Word **sửa trực tiếp** được (chữ, công thức OMML, bảng);
* cấu trúc chưa biết → **giống bản gốc 100%** vì do chính LaTeX vẽ ra;
* **không bao giờ** còn `\\tenlenh` lọt ra Word, cũng không bao giờ mất nội dung.

Ba hàm công khai:

    rep = scan_coverage(src)              # đo độ phủ, trả về % và danh sách cấu trúc lạ
    new_src, ctx = protect(src, …)        # thay khối lạ bằng chỗ đánh dấu, dựng ảnh nền
    stats = restore(docx_path, ctx)       # đổi chỗ đánh dấu thành ảnh vector trong .docx

Tắt hoàn toàn: đặt biến môi trường `L2W_UNIVERSAL=0`.
Chọn mức an toàn:   `L2W_TRUTH_POLICY = safe | balanced | native_first`  (mặc định `balanced`).
"""
from __future__ import annotations

import os
import re
import sys
import traceback
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from . import structure_registry as REG
from .macro_fidelity import rewrite_user_macros
from .omml_fix import fix_math_accents
from .docx_media import MARKER_RE, MediaInjector, make_marker
from .tex_truth import TexTruth, describe_toolchain, extract_preamble, tex_engine_available

__all__ = ["scan_coverage", "protect", "restore", "CoverageReport", "ProtectContext", "is_enabled"]

# ─────────────────────────── cấu hình ───────────────────────────
POLICY_ENV = "L2W_TRUTH_POLICY"
ENABLE_ENV = "L2W_UNIVERSAL"
MAX_UNKNOWN_ENV_CHARS = 900      # môi trường lạ dài hơn mức này thì bóc vỏ, không chụp ảnh
EMU_PER_PT = 12700

_DRAW_PRIMITIVE_RE = re.compile(
    r"\\(?:draw|fill|filldraw|shade|shadedraw|path|node|coordinate|clip|pattern|addplot\d*|"
    r"pgf[a-zA-Z@]*|tkz[A-Za-z]*|ps[a-z]+|chemfig|foreach|matrix|pic|arrow|edge|useasboundingbox)\b")
_TEXTY_UNKNOWN_RE = re.compile(r"\\(?:item|cau|Cau|choice|loigiai|section|subsection)\b")


def _log(msg: str) -> None:
    try:
        print(f"[universal] {msg}", file=sys.stderr, flush=True)
    except Exception:
        pass


def is_enabled() -> bool:
    return os.environ.get(ENABLE_ENV, "1").strip().lower() not in ("0", "false", "off", "no")


def _policy() -> str:
    p = os.environ.get(POLICY_ENV, "balanced").strip().lower()
    return p if p in ("safe", "balanced", "native_first") else "balanced"


# ─────────────────────── che chú thích & verbatim ───────────────────────
def _mask(src: str) -> str:
    """Trả về chuỗi cùng độ dài, phần chú thích / verbatim bị thay bằng khoảng trắng.

    Nhờ giữ nguyên độ dài nên mọi chỉ số tìm được trên bản che đều dùng thẳng cho bản gốc."""
    out = list(src)
    n = len(src)
    i = 0
    while i < n:
        ch = src[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "%":
            j = src.find("\n", i)
            j = n if j < 0 else j
            for k in range(i, j):
                out[k] = " "
            i = j
            continue
        i += 1
    masked = "".join(out)
    for env in REG.VERBATIM_ENVIRONMENTS:
        for m in re.finditer(r"\\begin\{" + re.escape(env) + r"\}(.*?)\\end\{" + re.escape(env) + r"\}",
                             masked, re.S):
            masked = masked[: m.start()] + " " * (m.end() - m.start()) + masked[m.end():]
    for m in re.finditer(r"\\verb\*?(.)(.*?)\1", masked):
        masked = masked[: m.start()] + " " * (m.end() - m.start()) + masked[m.end():]
    return masked


_BEGIN_END_RE = re.compile(r"\\(begin|end)\s*\{([^{}]{1,60})\}")


@dataclass
class EnvSpan:
    name: str
    start: int          # vị trí của \begin
    end: int            # vị trí ngay sau \end{...}
    body_start: int
    body_end: int
    depth: int

    def body(self, src: str) -> str:
        return src[self.body_start:self.body_end]

    def text(self, src: str) -> str:
        return src[self.start:self.end]


def iter_environments(src: str, masked: str | None = None) -> list[EnvSpan]:
    """Liệt kê mọi môi trường, hiểu cả lồng nhau, bỏ qua chú thích/verbatim."""
    masked = masked if masked is not None else _mask(src)
    stack: list[tuple[str, int, int]] = []
    spans: list[EnvSpan] = []
    for m in _BEGIN_END_RE.finditer(masked):
        kind, name = m.group(1), m.group(2).strip()
        if kind == "begin":
            stack.append((name, m.start(), m.end()))
        else:
            for k in range(len(stack) - 1, -1, -1):
                if stack[k][0] == name:
                    bname, bstart, bend = stack[k]
                    spans.append(EnvSpan(bname, bstart, m.end(), bend, m.start(), k))
                    del stack[k:]
                    break
    spans.sort(key=lambda s: s.start)
    return spans


def _user_macros(src: str) -> set[str]:
    """Tên macro do chính người dùng định nghĩa — coi như đã được bộ tiền xử lý lo."""
    names: set[str] = set()
    for pat in (r"\\(?:re)?newcommand\*?\s*\{?\\([A-Za-z@]+)\}?",
                r"\\providecommand\*?\s*\{?\\([A-Za-z@]+)\}?",
                r"\\DeclareRobustCommand\*?\s*\{?\\([A-Za-z@]+)\}?",
                r"\\(?:New|Renew|Provide|Declare)DocumentCommand\s*\{?\\([A-Za-z@]+)\}?",
                r"\\def\s*\\([A-Za-z@]+)", r"\\let\s*\\([A-Za-z@]+)",
                r"\\DeclareMathOperator\*?\s*\{?\\([A-Za-z@]+)\}?",
                r"\\DeclarePairedDelimiter\s*\{?\\([A-Za-z@]+)\}?"):
        names.update(re.findall(pat, src))
    envs = set(re.findall(r"\\(?:re)?newenvironment\*?\s*\{([^{}]+)\}", src))
    names.update(envs)
    return names


def _match_braces(masked: str, pos: int) -> int:
    """`pos` trỏ vào dấu `{` → trả về vị trí ngay sau dấu `}` khớp với nó (-1 nếu hỏng)."""
    if pos >= len(masked) or masked[pos] != "{":
        return -1
    depth, i, n = 0, pos, len(masked)
    while i < n:
        c = masked[i]
        if c == "\\":
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return -1


MATH_ENVIRONMENTS = {"equation", "equation*", "align", "align*", "alignat", "alignat*", "gather",
                     "gather*", "multline", "multline*", "flalign", "flalign*", "displaymath",
                     "eqnarray", "eqnarray*", "math", "split", "aligned", "gathered", "cases"}


def _math_regions(src: str, masked: str) -> list[tuple[int, int]]:
    r"""Vùng công thức: `$…$`, `$$…$$`, `\(…\)`, `\[…\]` và các môi trường toán."""
    regions: list[tuple[int, int]] = []
    i, n = 0, len(masked)
    while i < n:
        c = masked[i]
        if c == "\\":
            if masked[i:i + 2] == "\\(":
                j = masked.find("\\)", i + 2)
                if j > 0:
                    regions.append((i, j + 2)); i = j + 2; continue
            elif masked[i:i + 2] == "\\[":
                j = masked.find("\\]", i + 2)
                if j > 0:
                    regions.append((i, j + 2)); i = j + 2; continue
            i += 2
            continue
        if c == "$":
            dd = masked[i:i + 2] == "$$"
            tok = "$$" if dd else "$"
            j = i + len(tok)
            while j < n:
                if masked[j] == "\\":
                    j += 2
                    continue
                if masked.startswith(tok, j):
                    break
                j += 1
            if j < n:
                regions.append((i, j + len(tok)))
                i = j + len(tok)
                continue
        i += 1
    for sp in iter_environments(src, masked):
        if sp.name in MATH_ENVIRONMENTS:
            regions.append((sp.start, sp.end))
    regions.sort()
    return regions


def _inline_picture_spans(src: str, masked: str, body_start: int,
                          covered: list[tuple[int, int]],
                          math_regions: list[tuple[int, int]] | None = None) -> list[EnvSpan]:
    """Bắt lệnh vẽ / hoá học nằm GIỮA DÒNG: `\\ce{...}`, `\\chemfig{...}`…"""
    out: list[EnvSpan] = []
    names = "|".join(sorted(REG.INLINE_PICTURE_COMMANDS, key=len, reverse=True))
    for m in re.finditer(r"\\(" + names + r")(?![A-Za-z])", masked):
        if m.start() < body_start or any(a <= m.start() < b for a, b in covered):
            continue
        i = m.end()
        while i < len(masked) and masked[i] in " \t":
            i += 1
        if i < len(masked) and masked[i] == "[":
            j = masked.find("]", i)
            i = j + 1 if j > 0 else i
        end = _match_braces(masked, i) if i < len(masked) and masked[i] == "{" else -1
        if end < 0:
            continue
        start = m.start()
        for a, b in (math_regions or []):
            if a <= start and end <= b:
                start, end = a, b      # nằm trong công thức → chụp trọn công thức
                break
        if any(not (end <= a or start >= b) for a, b in covered):
            continue
        out.append(EnvSpan(m.group(1), start, end, start, end, 0))
        covered.append((start, end))
    return out


# ─────────────────────────── đo độ phủ ───────────────────────────
@dataclass
class CoverageReport:
    total: int = 0
    native: int = 0
    picture: int = 0
    textual: int = 0
    ignored: int = 0
    user_macro: int = 0
    unknown: int = 0
    unknown_commands: dict[str, int] = field(default_factory=dict)
    unknown_environments: dict[str, int] = field(default_factory=dict)
    picture_blocks: int = 0

    @property
    def native_percent(self) -> float:
        base = self.total - self.ignored
        return round(100.0 * (self.native + self.user_macro) / base, 1) if base else 100.0

    @property
    def total_percent(self) -> float:
        """Độ phủ tổng: phần được dựng đúng (gốc + ảnh TeX) trên tổng số cấu trúc."""
        base = self.total - self.ignored
        if not base:
            return 100.0
        covered = self.native + self.user_macro + self.picture + self.textual
        return round(100.0 * covered / base, 1)

    def to_dict(self) -> dict[str, Any]:
        return {
            "tong_so_cau_truc": self.total,
            "dung_engine_goc": self.native,
            "macro_nguoi_dung": self.user_macro,
            "dung_anh_tex": self.picture,
            "boc_vo_giu_chu": self.textual,
            "bo_qua_preamble": self.ignored,
            "chua_biet": self.unknown,
            "do_phu_nguyen_ban_phan_tram": self.native_percent,
            "do_phu_tong_phan_tram": self.total_percent,
            "lenh_la": dict(sorted(self.unknown_commands.items(), key=lambda kv: -kv[1])[:40]),
            "moi_truong_la": dict(sorted(self.unknown_environments.items(), key=lambda kv: -kv[1])[:40]),
            "so_khoi_chup_anh": self.picture_blocks,
            "so_cau_truc_trong_so_dang_ky": REG.registry_size(),
        }


_CMD_RE = re.compile(r"\\([A-Za-z@]+)\*?")


def _picture_regions(src: str, masked: str, body_start: int, policy: str,
                     mine: set[str]) -> tuple[list[tuple[int, int]], list[EnvSpan]]:
    """Các vùng sẽ được giao cho TeX dựng — mọi thứ bên trong coi như đã phủ."""
    regions: list[tuple[int, int]] = []
    chosen: list[EnvSpan] = []
    for sp in iter_environments(src, masked):
        if sp.start < body_start or any(a <= sp.start and sp.end <= b for a, b in regions):
            continue
        tier = REG.tier_of_environment(sp.name)
        if tier == "PICTURE" or (tier is None and sp.name not in mine
                                 and _should_picture_unknown_env(sp.name, sp.body(src), policy)):
            regions.append((sp.start, sp.end))
            chosen.append(sp)
    chosen.extend(_standalone_picture_commands(src, masked, body_start, regions))
    chosen.extend(_inline_picture_spans(src, masked, body_start, regions,
                                        _math_regions(src, masked)))
    if policy == "safe":
        chosen.extend(_unknown_command_spans(src, masked, body_start, regions, mine))
    regions.sort()
    chosen.sort(key=lambda s: s.start)
    return regions, chosen


def scan_coverage(src: str, policy: str | None = None) -> CoverageReport:
    """Đo xem bản dán này có bao nhiêu phần trăm cấu trúc nằm trong tầm xử lý.

    Mọi thứ nằm trong một khối do TeX dựng được tính là **đã phủ** (vì cả khối sẽ giống
    bản gốc 100%), nên không đếm lẻ từng `\\draw`, `\\node` bên trong hình."""
    rep = CoverageReport()
    policy = policy or _policy()
    masked = _mask(src)
    body_start = 0
    m = re.search(r"\\begin\s*\{document\}", masked)
    if m:
        body_start = m.end()
    mine = _user_macros(src)
    regions, chosen = _picture_regions(src, masked, body_start, policy, mine)
    rep.picture_blocks = len(chosen)

    def _inside(pos: int) -> bool:
        return any(a <= pos < b for a, b in regions)

    for sp in chosen:
        rep.total += 1
        rep.picture += 1

    for sp in iter_environments(src, masked):
        if sp.start < body_start or _inside(sp.start):
            continue
        rep.total += 1
        tier = REG.tier_of_environment(sp.name)
        if sp.name in mine:
            rep.user_macro += 1
        elif tier == "NATIVE":
            rep.native += 1
        elif tier == "PICTURE":
            rep.picture += 1
        elif tier == "TEXT":
            rep.textual += 1
        elif tier in ("IGNORE", "VERBATIM"):
            rep.ignored += 1
        else:
            rep.unknown += 1
            rep.unknown_environments[sp.name] = rep.unknown_environments.get(sp.name, 0) + 1

    for cm in _CMD_RE.finditer(masked):
        if cm.start() < body_start or _inside(cm.start()):
            continue
        name = cm.group(1)
        if name in ("begin", "end"):
            continue
        rep.total += 1
        tier = REG.tier_of_command(name)
        if name in mine:
            rep.user_macro += 1
        elif tier == "NATIVE":
            rep.native += 1
        elif tier == "PICTURE":
            rep.picture += 1
        elif tier == "TEXT":
            rep.textual += 1
        elif tier == "IGNORE":
            rep.ignored += 1
        else:
            rep.unknown += 1
            rep.unknown_commands[name] = rep.unknown_commands.get(name, 0) + 1
    return rep


def _unknown_command_spans(src: str, masked: str, body_start: int,
                           covered: list[tuple[int, int]], mine: set[str]) -> list[EnvSpan]:
    """Chỉ dùng ở chính sách `safe`: lệnh chưa ai khai báo → chụp trọn đoạn văn chứa nó,
    bảo đảm tuyệt đối không có chữ nào sai so với bản gốc."""
    out: list[EnvSpan] = []
    for cm in _CMD_RE.finditer(masked):
        name = cm.group(1)
        if name in ("begin", "end") or cm.start() < body_start:
            continue
        if name in mine or REG.tier_of_command(name) is not None:
            continue
        if any(a <= cm.start() < b for a, b in covered):
            continue
        start = masked.rfind("\n\n", 0, cm.start())
        start = body_start if start < 0 else max(body_start, start + 2)
        nxt = masked.find("\n\n", cm.end())
        end = len(src) if nxt < 0 else nxt
        if any(not (end <= a or start >= b) for a, b in covered):
            continue
        out.append(EnvSpan("doan-chua-" + name, start, end, start, end, 0))
        covered.append((start, end))
    return out


def _build_stubs(blocks_text: str, mine: set[str]) -> tuple[str, int]:
    r"""Sinh khung rỗng cho mọi tên lạ xuất hiện trong các khối sẽ giao cho TeX.

    Nhờ vậy `\begin{bangxetdauthan}` hay `\hang{...}` không làm TeX báo lỗi nữa mà vẫn
    in ra đúng nội dung bên trong.  Dùng `\@ifundefined` / `\providecommand` nên không
    bao giờ đè lên định nghĩa thật của gói."""
    masked = _mask(blocks_text)
    envs = {n for n in re.findall(r"\\begin\s*\{([^{}]{1,60})\}", masked)}
    envs |= {n for n in re.findall(r"\\end\s*\{([^{}]{1,60})\}", masked)}
    cmds = {m.group(1) for m in _CMD_RE.finditer(masked)}
    lines = ["\\makeatletter"]
    for e in sorted(envs):
        if e in mine or REG.tier_of_environment(e) is not None or not re.fullmatch(r"[A-Za-z@*]+", e):
            continue
        lines.append("\\@ifundefined{%s}{\\newenvironment{%s}[1][]{}{}}{}" % (e, e))
    for c in sorted(cmds):
        if c in ("begin", "end") or c in mine or REG.tier_of_command(c) is not None:
            continue
        if not re.fullmatch(r"[A-Za-z@]+", c):
            continue
        lines.append("\\providecommand{\\%s}{}" % c)
    lines.append("\\makeatother")
    return ("\n".join(lines) + "\n", len(lines) - 2) if len(lines) > 2 else ("", 0)


# ─────────────────────── thay khối lạ bằng chỗ đánh dấu ───────────────────────
@dataclass
class ProtectContext:
    enabled: bool = False
    renderer: TexTruth | None = None
    future: Future | None = None
    blocks: dict[int, dict[str, Any]] = field(default_factory=dict)   # index → {key, inline, label}
    coverage: CoverageReport | None = None
    macro_fidelity: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    reason: str = ""
    loi_khuyen: str = ""          # câu hướng dẫn cụ thể khi thiếu công cụ trên máy
    failed_blocks: list[dict[str, Any]] = field(default_factory=list)
    tong_khoi_can_dung: int = 0   # số khối ĐÁNG LẼ phải dựng (kể cả khối hỏng)

    def summary(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "bat": self.enabled,
            "ly_do": self.reason,
            "so_khoi_giao_cho_tex": len(self.blocks),
            "tong_khoi_can_dung": self.tong_khoi_can_dung,
            "loi_khuyen": self.loi_khuyen,
            "khoi_chua_dung_duoc": self.failed_blocks[:12],
            "canh_bao": self.warnings[:30],
            "ton_trong_macro_nguoi_dung": self.macro_fidelity,
        }
        if self.coverage is not None:
            out["do_phu"] = self.coverage.to_dict()
        if self.renderer is not None:
            out["tex"] = {k: v for k, v in self.renderer.stats.items() if k != "toolchain"}
            out["cong_cu"] = self.renderer.stats.get("toolchain")
        return out


def _should_picture_unknown_env(name: str, body: str, policy: str) -> bool:
    if policy == "native_first":
        return False
    if _DRAW_PRIMITIVE_RE.search(body):
        return True
    if _TEXTY_UNKNOWN_RE.search(body):
        return False          # bên trong là đề bài / danh sách → để engine dựng chữ sửa được
    if len(body) <= MAX_UNKNOWN_ENV_CHARS:
        return True
    return policy == "safe"


def protect(src: str, *, base_dir: Path | None = None, workdir: Path | None = None,
            background: bool = True) -> tuple[str, ProtectContext]:
    """Thay mọi khối "phải để TeX dựng" bằng chỗ đánh dấu và bắt đầu dựng ảnh ở luồng nền.

    Trả về `(mã_nguồn_mới, ngữ_cảnh)`.  Không bao giờ ném lỗi: có trục trặc gì thì trả lại
    đúng mã nguồn ban đầu, pipeline cũ chạy y như trước."""
    ctx = ProtectContext()
    try:
        if not is_enabled():
            ctx.reason = "tắt bằng L2W_UNIVERSAL=0"
            return src, ctx
        # Bước -1 — trả lại đúng macro người dùng tự định nghĩa mà engine gốc sẽ đè lên
        # (\choice, \True…): nở ngay trong mã nguồn để engine không còn thấy mà đè.
        mf = rewrite_user_macros(src)
        if mf.expanded:
            src = mf.text
            ctx.macro_fidelity = mf.summary()
            ctx.warnings.append(
                "Đã nở " + ", ".join(f"\\{k} ({v} lần)" for k, v in mf.expanded.items())
                + " theo đúng định nghĩa của bạn, thay vì để engine dùng khuôn riêng.")
        elif mf.skipped:
            ctx.macro_fidelity = mf.summary()
        ctx.coverage = scan_coverage(src)
        tc = describe_toolchain()
        if not tc.get("tuyen_dung_duoc"):
            ctx.reason = "máy thiếu công cụ dựng ảnh → giữ nguyên engine gốc"
            ctx.loi_khuyen = tc.get("loi_khuyen", "")
            ctx.tong_khoi_can_dung = ctx.coverage.picture_blocks if ctx.coverage else 0
            ctx.warnings.append(ctx.loi_khuyen or "Máy thiếu công cụ dựng ảnh từ LaTeX.")
            return src, ctx
        policy = _policy()
        masked = _mask(src)
        spans = iter_environments(src, masked)
        mine = _user_macros(src)
        body_start = 0
        mdoc = re.search(r"\\begin\s*\{document\}", masked)
        if mdoc:
            body_start = mdoc.end()

        chosen: list[EnvSpan] = []
        covered_ranges: list[tuple[int, int]] = []

        def _inside_chosen(sp: EnvSpan) -> bool:
            return any(a <= sp.start and sp.end <= b for a, b in covered_ranges)

        for sp in spans:
            if sp.start < body_start or _inside_chosen(sp):
                continue
            tier = REG.tier_of_environment(sp.name)
            take = False
            if tier == "PICTURE":
                take = True
            elif tier is None and sp.name not in mine:
                take = _should_picture_unknown_env(sp.name, sp.body(src), policy)
                if not take:
                    ctx.warnings.append(
                        f"Môi trường lạ '{sp.name}' quá dài hoặc chứa đề bài → bóc vỏ giữ chữ "
                        f"(đặt {POLICY_ENV}=safe nếu muốn chụp ảnh nguyên khối).")
            if take:
                chosen.append(sp)
                covered_ranges.append((sp.start, sp.end))

        # lệnh vẽ đứng một mình (\tikz{...};  \tkzTabInit…) → gom cả dòng
        for sp in _standalone_picture_commands(src, masked, body_start, covered_ranges):
            chosen.append(sp)
            covered_ranges.append((sp.start, sp.end))
        # lệnh vẽ/hoá học giữa dòng → ảnh inline, căn đúng đường cơ sở
        inline_spans = set()
        for sp in _inline_picture_spans(src, masked, body_start, covered_ranges,
                                        _math_regions(src, masked)):
            chosen.append(sp)
            inline_spans.add((sp.start, sp.end))
        if policy == "safe":
            for sp in _unknown_command_spans(src, masked, body_start, covered_ranges, mine):
                chosen.append(sp)

        if not chosen:
            ctx.enabled = True
            ctx.reason = "không có khối nào cần TeX dựng"
            return src, ctx

        preamble = extract_preamble(src)
        # Khung rỗng cho tên lạ nay được tạo THEO PHẢN ỨNG trong tex_truth: chỉ khi TeX thật sự
        # báo thiếu. Đoán trước như bản cũ có lúc đè lên \\node/\\draw của TikZ làm hỏng cả hình.
        renderer = TexTruth(preamble, workdir=workdir)
        ctx.renderer = renderer
        chosen.sort(key=lambda s: s.start)
        keys = [renderer.add(sp.text(src), inline=(sp.start, sp.end) in inline_spans, label=sp.name)
                for sp in chosen]

        optimistic = background and os.environ.get("L2W_TRUTH_PARALLEL", "0").strip() in ("1", "true", "on", "yes")
        assets: dict[str, Any] = {}
        if optimistic:
            # Nhanh hơn: dựng ảnh song song với engine. Khối nào hỏng thì mất hình.
            pool = ThreadPoolExecutor(max_workers=1)
            ctx.future = pool.submit(renderer.build)
            pool.shutdown(wait=False)
        else:
            # Mặc định — AN TOÀN: dựng xong mới quyết định.  Khối nào TeX dựng hỏng thì
            # GIỮ NGUYÊN mã gốc để engine cũ xử lý y như trước ⇒ không bao giờ tệ hơn trước.
            assets = renderer.build()
            ctx.future = _DoneFuture(assets)

        pieces: list[str] = []
        prev = 0
        kept = 0
        idx = 0
        for sp, key in zip(chosen, keys):
            asset = assets.get(key) if assets else None
            if assets and (asset is None or not asset.ok):
                head = " ".join(sp.text(src)[:110].split())
                ctx.warnings.append(
                    f"TeX chưa dựng được khối '{sp.name}' → giữ nguyên cách xử lý cũ của engine"
                    + (f" | {asset.error}" if asset is not None and asset.error else "")
                    + f" | mã nguồn: {head}…")
                ctx.failed_blocks.append({
                    "moi_truong": sp.name,
                    "ma_nguon": head,
                    "loi": (asset.error if asset is not None else ""),
                    "tex_bao": (asset.tex_errors if asset is not None else []),
                })
                continue                       # không thay gì cả — hành vi y như trước khi có lớp này
            is_inline = (sp.start, sp.end) in inline_spans
            ctx.blocks[idx] = {"key": key, "inline": is_inline, "label": sp.name, "src": sp.text(src)}
            pieces.append(src[prev:sp.start])
            pieces.append(make_marker(idx) if is_inline
                          else "\n\n\\par\n" + make_marker(idx) + "\n\\par\n\n")
            prev = sp.end
            idx += 1
            kept += 1
        pieces.append(src[prev:])
        new_src = "".join(pieces)

        ctx.enabled = True
        ctx.tong_khoi_can_dung = len(chosen)
        ctx.loi_khuyen = str(renderer.stats.get("loi_khuyen") or "")
        if assets and kept == 0 and chosen:
            ctx.warnings.insert(0, "KHÔNG dựng được khối hình nào. " + (ctx.loi_khuyen or
                                "Xem mục cảnh báo bên dưới để biết TeX báo lỗi gì."))
        ctx.reason = (f"giao {kept}/{len(chosen)} khối cho TeX dựng "
                      f"(chính sách: {policy}{', chạy song song' if optimistic else ''})")
        _log(ctx.reason)
        return new_src, ctx
    except Exception as exc:
        _log(f"LỖI khi bảo vệ khối lạ → bỏ qua lớp này: {type(exc).__name__}: {exc}")
        _log(traceback.format_exc())
        ctx.enabled = False
        ctx.reason = f"lỗi nội bộ: {type(exc).__name__}: {exc}"
        ctx.warnings.append(ctx.reason)
        return src, ctx


class _DoneFuture:
    def __init__(self, value: Any) -> None:
        self._v = value

    def result(self, timeout: float | None = None) -> Any:
        return self._v


_STANDALONE_PICTURE_RE = re.compile(
    r"\\(tikz|chemfig|schemestart|tkzTabInit|xymatrix|Qcircuit|smartdiagram|qrcode|QRCode)\b")


def _standalone_picture_commands(src: str, masked: str, body_start: int,
                                 covered: list[tuple[int, int]]) -> list[EnvSpan]:
    """Bắt các lệnh vẽ không nằm trong môi trường — gom trọn đoạn chứa nó."""
    out: list[EnvSpan] = []
    for m in _STANDALONE_PICTURE_RE.finditer(masked):
        if m.start() < body_start:
            continue
        if any(a <= m.start() < b for a, b in covered):
            continue
        start = masked.rfind("\n\n", 0, m.start())
        start = 0 if start < 0 else start + 2
        start = max(start, body_start)
        nxt = masked.find("\n\n", m.end())
        end = len(src) if nxt < 0 else nxt
        if any(not (end <= a or start >= b) for a, b in covered):
            continue
        out.append(EnvSpan(m.group(1), start, end, m.end(), end, 0))
        covered.append((start, end))
    return out


# ─────────────────────── đổi chỗ đánh dấu thành ảnh ───────────────────────
def restore(docx_path: Path, ctx: ProtectContext, *, timeout: float = 600.0) -> dict[str, Any]:
    """Mở .docx, đổi mọi chỗ đánh dấu thành ảnh vector + sửa dấu trên công thức.

    Không bao giờ ném lỗi ra ngoài.  Khối KHÔNG dựng được ảnh thì chỗ đánh dấu được thay bằng một thông
    báo hiển thị + mã nguồn gốc của khối (không xoá âm thầm), và được liệt kê trong `khoi_chua_dung`."""
    stats: dict[str, Any] = {"bat": ctx.enabled, "so_anh_chen": 0, "so_khoi_hong": 0,
                             "tong_khoi_can_dung": ctx.tong_khoi_can_dung,
                             "loi_khuyen": ctx.loi_khuyen, "canh_bao": list(ctx.warnings[:10]),
                             "khoi_chua_dung": []}
    stats["sua_dau_cong_thuc"] = fix_math_accents(docx_path)
    if not ctx.enabled or not ctx.blocks:
        stats["khoi_chua_dung"] = _mark_unresolved_markers(docx_path, ctx, "không có ảnh cho chỗ đánh dấu")
        stats["so_khoi_hong"] = len(stats["khoi_chua_dung"])
        return stats
    try:
        assets = ctx.future.result(timeout=timeout) if ctx.future is not None else {}
    except Exception as exc:
        stats["canh_bao"].append(f"Dựng ảnh bằng TeX thất bại: {type(exc).__name__}: {exc}")
        stats["khoi_chua_dung"] = _mark_unresolved_markers(docx_path, ctx, f"dựng ảnh thất bại ({type(exc).__name__})")
        stats["so_khoi_hong"] = len(stats["khoi_chua_dung"])
        return stats
    failed: list[dict[str, Any]] = []
    try:
        import docx as _docx
        doc = _docx.Document(str(docx_path))
        max_w_pt = _content_width_pt(doc)
        inj = MediaInjector(doc)

        def resolver(index: int) -> dict[str, Any] | None:
            info = ctx.blocks.get(index)
            if not info:
                failed.append({"marker": index, "ly_do": "chỗ đánh dấu lạ (không có trong ngữ cảnh)"})
                return {"text": f"[L2W: KHỐI #{index} KHÔNG RÕ NGUỒN — xem báo cáo]"}
            asset = assets.get(info["key"])
            if asset is None or not asset.ok:
                err = (asset.error if asset is not None else "") or "TeX không trả về ảnh"
                stats["canh_bao"].append(f"[{info['label']}] {err}")
                failed.append({"marker": index, "moi_truong": info["label"], "ly_do": err})
                return {"text": _unresolved_text(index, info, err)}
            w, h = asset.width_pt, asset.height_pt
            if max_w_pt and w > max_w_pt:
                h = h * (max_w_pt / w)
                w = max_w_pt
            out = {
                "data": asset.data, "ext": asset.ext, "png": asset.png_data,
                "cx": max(1, int(round(w * EMU_PER_PT))), "cy": max(1, int(round(h * EMU_PER_PT))),
                "name": f"LaTeX {info['label']} ",
                "descr": f"l2w:truth:{index}",
            }
            if info.get("inline") and asset.depth_pt:
                out["position_half_pt"] = -int(round(asset.depth_pt * 2))
            return out

        for err in (ctx.renderer.stats.get("tex_errors") or [])[:8] if ctx.renderer else []:
            stats["canh_bao"].append("TeX báo: " + err)
        inj.replace_markers(resolver)
        stats["so_anh_chen"] = inj.stats["markers_replaced"]
        stats["chi_tiet_anh"] = dict(inj.stats)
        missing_in_docx = [i for i in ctx.blocks if i not in inj.found]
        for i in missing_in_docx:
            failed.append({"marker": i, "moi_truong": ctx.blocks[i]["label"],
                           "ly_do": "chỗ đánh dấu không còn trong Word (engine đã làm mất)"})
        doc.save(str(docx_path))
        failed.extend(_mark_unresolved_markers(docx_path, ctx, "chưa thay được"))
        stats["khoi_chua_dung"] = failed
        stats["so_khoi_hong"] = len(failed)
        _log(f"đã chèn {stats['so_anh_chen']}/{len(ctx.blocks)} ảnh vector vào Word; {len(failed)} khối chưa dựng")
    except Exception as exc:
        stats["canh_bao"].append(f"Chèn ảnh vào Word thất bại: {type(exc).__name__}: {exc}")
        _log(traceback.format_exc())
        stats["khoi_chua_dung"] = failed + _mark_unresolved_markers(docx_path, ctx, f"lỗi chèn ảnh ({type(exc).__name__})")
        stats["so_khoi_hong"] = len(stats["khoi_chua_dung"])
    finally:
        if ctx.renderer is not None:
            ctx.renderer.close()
    return stats


def _unresolved_text(index: int, info: dict[str, Any] | None, reason: str) -> str:
    label = (info or {}).get("label", "?")
    src = " ".join(((info or {}).get("src") or "").split())
    return (f"[L2W: CHƯA DỰNG ĐƯỢC KHỐI '{label}' (#{index}) — {reason}. Mã nguồn gốc: {src}]"
            if src else f"[L2W: CHƯA DỰNG ĐƯỢC KHỐI '{label}' (#{index}) — {reason}]")


def _content_width_pt(doc) -> float:
    try:
        sec = doc.sections[0]
        emu = int(sec.page_width) - int(sec.left_margin) - int(sec.right_margin)
        return max(72.0, emu / EMU_PER_PT)
    except Exception:
        return 467.0          # A4, lề 2cm


def _mark_unresolved_markers(docx_path: Path, ctx: ProtectContext | None, reason: str) -> list[dict[str, Any]]:
    """Mọi chỗ đánh dấu còn sót → thay bằng thông báo hiển thị + mã nguồn gốc của khối.

    Thay cho `_strip_leftover_markers` cũ (xoá chỗ đánh dấu ⇒ hình mất mà Word trông "sạch" và trạng thái
    vẫn thành công). Trả về danh sách khối chưa dựng để đưa vào báo cáo."""
    out: list[dict[str, Any]] = []
    try:
        import docx as _docx
        from docx.oxml.ns import qn
        doc = _docx.Document(str(docx_path))
        touched = False
        blocks = ctx.blocks if ctx is not None else {}
        roots = [doc.element.body]
        for sec in doc.sections:
            for hf in (sec.header, sec.footer):
                try:
                    roots.append(hf._element)
                except Exception:
                    pass
        for root in roots:
            for t in root.iter(qn("w:t"), qn("m:t")):
                if t.text and MARKER_RE.search(t.text):
                    def _sub(m: re.Match) -> str:
                        idx = int(m.group(1))
                        info = blocks.get(idx)
                        out.append({"marker": idx, "moi_truong": (info or {}).get("label"), "ly_do": reason})
                        return _unresolved_text(idx, info, reason)
                    t.text = MARKER_RE.sub(_sub, t.text)
                    touched = True
        if touched:
            doc.save(str(docx_path))
    except Exception as exc:
        out.append({"marker": None, "ly_do": f"không mở được tệp để đánh dấu khối lỗi: {type(exc).__name__}: {exc}"})
    return out


def _strip_leftover_markers(docx_path: Path) -> None:
    """(Giữ tên cũ cho tương thích) — nay KHÔNG xoá mà thay bằng thông báo hiển thị."""
    _mark_unresolved_markers(docx_path, None, "không có ảnh cho chỗ đánh dấu")
