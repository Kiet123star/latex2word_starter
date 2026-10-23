"""`backend/pdf_reader.py` — đọc PDF THẬT (không suy LaTeX rồi tìm .tex trùng tên) thành manifest khối.

Backend và vai trò:
* pikepdf      — mở/kiểm tra cấu trúc (mật khẩu, mã hoá, MediaBox/CropBox/Rotate, tạo PDF một vùng để cắt vector);
* pdfplumber   — ký tự (font, cỡ, toạ độ, đã áp CTM/Form XObject), đường/chữ nhật/đường cong, ảnh, bảng kẻ;
* pypdfium2    — render trang/vùng (PNG) và đọc chế độ vẽ chữ (chữ ẩn của lớp OCR), lỗi ánh xạ Unicode;
* pdftocairo   — (tuỳ chọn) xuất vùng thành SVG vector;
* tesseract    — (tuỳ chọn) OCR trang scan khi có gói ngôn ngữ phù hợp.

Hệ toạ độ: xem `pdf_geometry.PageGeom` (pt, gốc trên-trái trang hiển thị, y xuống).

Mỗi khối có block_id, loại, trang, bbox, thứ tự, khối cha, phương thức dựng (text_native, math_omml,
table_native, vector_image, raster_image, page_image, unresolved), trạng thái, cảnh báo, lý do fallback,
tài sản (ảnh cắt) và mức tin cậy (heuristic — KHÔNG phải xác suất đã hiệu chuẩn).

Nhận dạng toán 2D ở đây cố ý hẹp: chỉ số trên/dưới, phân số một tầng, căn một tầng; cấu trúc khác hoặc
glyph không ánh xạ → cắt ẢNH vùng công thức (math_as_image) thay vì đoán.
"""
from __future__ import annotations

import ctypes
import io
import json
import re
import shutil
import statistics
import subprocess
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .pdf_geometry import PageGeom

__all__ = ["analyze_pdf", "PdfInputError", "crop_region"]

Rect = tuple[float, float, float, float]

MATH_FONT_RE = re.compile(r"CMMI|CMSY|CMEX|CMBSY|MSAM|MSBM|LMMath|LatinModernMath|Cambria\s*Math|CambriaMath|"
                          r"STIX\w*Math|XITS\w*Math|Asana|TeXGyre\w*Math|Symbol|MTExtra|MT\s*Extra|Euclid|RSFS|EUFM|"
                          r"EUSM|esint|wasy|txsy|pxsy|txmi|pxmi|NewCMMath|FiraMath|TeXGyrePagellaMath", re.I)
MATH_EXT_FONT_RE = re.compile(r"CMEX|LMMathExtension|txex|pxex|esint|MTExtra", re.I)
BOLD_RE = re.compile(r"Bold|Black|Heavy|Semibold|Demi|CMBX|CMB10|-B\b|,Bold", re.I)
ITALIC_RE = re.compile(r"Italic|Oblique|CMTI|CMSL|-It\b|Slant|,Italic", re.I)
MATH_CHARS = set("=+−×÷±∓≤≥≠≈≡∼≃≅∝∈∉∋⊂⊆⊃⊇∪∩∖∅∀∃∄∧∨¬∂∇∞→←↔⇒⇐⇔↦∑∏∫∬∭∮√∠⊥∥△∘⋅·′″‰<>|^"
                 "αβγδεϵζηθϑικλμνξπρσςτυφϕχψωΓΔΘΛΞΠΣΥΦΨΩ")
OPERATOR_CHARS = set("+−-=×÷±∓≤≥≠≈≡<>/()[]{}|,.;:!′'^_*·⋅")
LIGATURES = {"ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl", "ﬅ": "st", "ﬆ": "st"}
SPACING_ACCENTS = {"´": "́", "`": "̀", "ˆ": "̂", "^": "̂", "˜": "̃", "~": "̃",
                   "¨": "̈", "˙": "̇", "˘": "̆", "ˇ": "̌", "¯": "̄", "ˉ": "̄",
                   "˚": "̊", "̉": "̉"}
UNMAPPED_RE = re.compile(r"\(cid:\d+\)|[-�]")
FUNC_NAMES = {"sin", "cos", "tan", "cot", "ln", "log", "lim", "max", "min", "exp", "sup", "inf", "det", "arcsin",
              "arccos", "arctan", "sinh", "cosh", "tanh", "lg"}
LATEX_OF = {"−": "-", "×": "\\times ", "÷": "\\div ", "±": "\\pm ", "∓": "\\mp ", "≤": "\\le ", "≥": "\\ge ",
            "≠": "\\ne ", "≈": "\\approx ", "≡": "\\equiv ", "∞": "\\infty ", "→": "\\to ", "⇒": "\\Rightarrow ",
            "⇔": "\\Leftrightarrow ", "∈": "\\in ", "∉": "\\notin ", "⊂": "\\subset ", "⊆": "\\subseteq ",
            "∪": "\\cup ", "∩": "\\cap ", "∅": "\\emptyset ", "∀": "\\forall ", "∃": "\\exists ", "∂": "\\partial ",
            "·": "\\cdot ", "⋅": "\\cdot ", "∠": "\\angle ", "⊥": "\\perp ", "∥": "\\parallel ", "′": "'",
            "{": "\\{", "}": "\\}", "%": "\\%", "&": "\\&", "#": "\\#", "$": "\\$", "_": "\\_", "\\": "\\backslash ",
            "^": "\\hat{}", "~": "\\sim ", "∑": "\\sum ", "∏": "\\prod ", "∫": "\\int ", "√": "\\surd "}
GREEK = {"α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "varepsilon", "ϵ": "epsilon", "ζ": "zeta",
         "η": "eta", "θ": "theta", "ϑ": "vartheta", "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "ν": "nu",
         "ξ": "xi", "π": "pi", "ρ": "rho", "σ": "sigma", "τ": "tau", "υ": "upsilon", "φ": "varphi", "ϕ": "phi",
         "χ": "chi", "ψ": "psi", "ω": "omega", "Γ": "Gamma", "Δ": "Delta", "Θ": "Theta", "Λ": "Lambda", "Ξ": "Xi",
         "Π": "Pi", "Σ": "Sigma", "Φ": "Phi", "Ψ": "Psi", "Ω": "Omega"}


class PdfInputError(ValueError):
    def __init__(self, message: str, code: str = "pdf_error") -> None:
        super().__init__(message)
        self.code = code


@dataclass
class Ch:
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    size: float
    font: str
    upright: bool = True
    hidden: bool = False
    unmapped: bool = False
    owner: str | None = None
    bx: float = 0.0          # gốc đường cơ sở (hệ chuẩn)
    by: float = 0.0
    merged: bool = False     # dấu rời đã được gộp vào chữ cơ sở (không tính riêng nữa)

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.top + self.bottom) / 2

    @property
    def h(self) -> float:
        return max(0.1, self.bottom - self.top)

    @property
    def bold(self) -> bool:
        return bool(BOLD_RE.search(self.font))

    @property
    def italic(self) -> bool:
        return bool(ITALIC_RE.search(self.font))

    @property
    def mathfont(self) -> bool:
        return bool(MATH_FONT_RE.search(self.font))


@dataclass
class Obj:
    kind: str           # rect | line | curve | image
    bbox: Rect
    owner: str | None = None
    meta: dict[str, Any] = field(default_factory=dict)

    @property
    def w(self) -> float:
        return self.bbox[2] - self.bbox[0]

    @property
    def h(self) -> float:
        return self.bbox[3] - self.bbox[1]

    @property
    def thin_h(self) -> bool:
        return self.h <= 1.6 and self.w >= 2.0

    @property
    def thin_v(self) -> bool:
        return self.w <= 1.6 and self.h >= 2.0


def _union(rects) -> Rect:
    rects = list(rects)
    return (min(r[0] for r in rects), min(r[1] for r in rects), max(r[2] for r in rects), max(r[3] for r in rects))


def _area(r: Rect) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _inter(a: Rect, b: Rect) -> Rect | None:
    r = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    return r if r[2] > r[0] and r[3] > r[1] else None


def _touch(a: Rect, b: Rect) -> bool:
    """Chạm/giao nhau kể cả hộp suy biến (đường kẻ dày 0)."""
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _expand(r: Rect, m: float) -> Rect:
    return (r[0] - m, r[1] - m, r[2] + m, r[3] + m)


def _contains(r: Rect, x: float, y: float) -> bool:
    return r[0] <= x <= r[2] and r[1] <= y <= r[3]


def _r(rect: Rect) -> list[float]:
    return [round(v, 2) for v in rect]


# ───────────────────────────── cắt vùng ─────────────────────────────
def crop_region(pdf_path: Path, page_index: int, geom: PageGeom, rect: Rect, out_dir: Path, name: str, *,
                dpi: int, want_svg: bool) -> dict[str, Any]:
    """Cắt vùng `rect` (hệ chuẩn) của trang GỐC: PNG bằng pypdfium2; SVG bằng PDF một-vùng + pdftocairo."""
    import pypdfium2 as pdfium
    out_dir.mkdir(parents=True, exist_ok=True)
    rect = geom.clip(rect)
    W, H = geom.width, geom.height
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        page = doc[page_index]
        crop = (rect[0], H - rect[3], W - rect[2], rect[1])
        img = page.render(scale=dpi / 72.0, crop=crop, may_draw_forms=True).to_pil()
    finally:
        doc.close()
    png_path = out_dir / f"{name}.png"
    img.save(png_path, "PNG")
    res: dict[str, Any] = {"png": str(png_path), "width_pt": round(rect[2] - rect[0], 2),
                           "height_pt": round(rect[3] - rect[1], 2), "dpi": dpi, "bbox": _r(rect)}
    if want_svg and shutil.which("pdftocairo"):
        try:
            import pikepdf
            ux0, uy0, ux1, uy1 = geom.display_rect_to_user(rect)
            region_pdf = out_dir / f"{name}.region.pdf"
            with pikepdf.open(str(pdf_path)) as src:
                dst = pikepdf.new()
                dst.pages.append(src.pages[page_index])
                pg = dst.pages[0]
                box = pikepdf.Array([ux0, uy0, ux1, uy1])
                pg.obj["/MediaBox"] = box
                pg.obj["/CropBox"] = box
                if "/TrimBox" in pg.obj:
                    del pg.obj["/TrimBox"]
                if "/BleedBox" in pg.obj:
                    del pg.obj["/BleedBox"]
                pg.obj["/Rotate"] = geom.rotation
                dst.save(str(region_pdf))
            svg_path = out_dir / f"{name}.svg"
            subprocess.run(["pdftocairo", "-svg", str(region_pdf), str(svg_path)], capture_output=True, timeout=120)
            if svg_path.exists() and svg_path.stat().st_size > 0:
                res["svg"] = str(svg_path)
        except Exception as exc:
            res["svg_error"] = f"{type(exc).__name__}: {exc}"
    return res


# ───────────────────────────── ký tự ─────────────────────────────
def _pdfium_char_flags(pdf_path: Path, page_index: int, geom: PageGeom,
                      tight: list[tuple[str, Rect]] | None = None) -> tuple[list[Rect], list[Rect]]:
    """(hộp chữ ẩn — render mode 3/7, hộp ký tự lỗi ánh xạ Unicode) trong hệ chuẩn.
    Nếu truyền `tight`, thêm (ký tự, hộp glyph thật) — pdfminer tính hộp theo cỡ chữ nên sai với
    glyph lớn của font mở rộng toán (∫, ∑, ngoặc lớn CMEX)."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw
    hidden: list[Rect] = []
    bad: list[Rect] = []
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        page = doc[page_index]
        tp = page.get_textpage()
        n = tp.count_chars()
        l, r, b, t = ctypes.c_double(), ctypes.c_double(), ctypes.c_double(), ctypes.c_double()
        for i in range(n):
            if raw.FPDFText_IsGenerated(tp.raw, i):
                continue
            if not raw.FPDFText_GetCharBox(tp.raw, i, ctypes.byref(l), ctypes.byref(r), ctypes.byref(b),
                                           ctypes.byref(t)):
                continue
            rect = geom.user_rect(l.value, b.value, r.value, t.value)
            if tight is not None:
                try:
                    tight.append((chr(raw.FPDFText_GetUnicode(tp.raw, i)), rect))
                except (ValueError, OverflowError):
                    pass
            obj = raw.FPDFText_GetTextObject(tp.raw, i)
            mode = raw.FPDFTextObj_GetTextRenderMode(obj) if obj else 0
            if mode in (3, 7):
                hidden.append(rect)
            if raw.FPDFText_HasUnicodeMapError(tp.raw, i):
                bad.append(rect)
        tp.close()
    finally:
        doc.close()
    return hidden, bad


def _match_any(rect: Rect, boxes: list[Rect], tol: float = 1.5) -> bool:
    cx, cy = (rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2
    for b in boxes:
        if b[0] - tol <= cx <= b[2] + tol and b[1] - tol <= cy <= b[3] + tol:
            return True
    return False


def _fix_extension_glyph(ch: "Ch", tight: list[tuple[str, Rect]]) -> None:
    best = None
    bestd = 1e9
    for u, r in tight:
        if u != ch.text:
            continue
        if r[0] > ch.x1 + 2 or r[2] < ch.x0 - 2:
            continue
        d = abs((r[0] + r[2]) / 2 - ch.cx) + abs(r[1] - ch.top) * 0.2
        if d < bestd:
            best, bestd = r, d
    if best is not None and (best[3] - best[1]) > 0.3:
        ch.x0, ch.top, ch.x1, ch.bottom = best


def _load_chars(page, geom: PageGeom, hidden_boxes: list[Rect], bad_boxes: list[Rect],
                tight: list[tuple[str, Rect]] | None = None) -> list[Ch]:
    out: list[Ch] = []
    for c in page.chars:
        rect = geom.plumber_rect(c["x0"], c["top"], c["x1"], c["bottom"])
        txt = c.get("text") or ""
        txt = LIGATURES.get(txt, txt)
        ch = Ch(txt, *rect, size=float(c.get("size") or (rect[3] - rect[1])), font=str(c.get("fontname") or ""),
                upright=bool(c.get("upright", True)))
        mtx = c.get("matrix")
        if mtx and len(mtx) == 6:
            ch.bx, ch.by = geom.plumber_to_display(float(mtx[4]), float(page.height) - float(mtx[5]))
        else:
            ch.bx, ch.by = rect[0], rect[3] - 0.2 * ch.size
        ch.unmapped = bool(UNMAPPED_RE.search(txt)) or (bool(bad_boxes) and _match_any(rect, bad_boxes, 0.5))
        if hidden_boxes and _match_any(rect, hidden_boxes, 0.8):
            ch.hidden = True
        if tight and (MATH_EXT_FONT_RE.search(ch.font) or txt in "∫∑∏√(){}[]|"):
            _fix_extension_glyph(ch, tight)
        out.append(ch)
    return out


def _load_objects(page, geom: PageGeom) -> list[Obj]:
    objs: list[Obj] = []
    for kind, lst in (("rect", page.rects), ("line", page.lines), ("curve", page.curves)):
        for o in lst:
            rect = geom.plumber_rect(o["x0"], o["top"], o["x1"], o["bottom"])
            meta = {"fill": bool(o.get("fill")), "stroke": bool(o.get("stroke", True)),
                    "linewidth": float(o.get("linewidth") or 0)}
            if kind == "curve":
                pts = o.get("pts") or []
                meta["npts"] = len(pts)
            if kind == "line":
                # đường chéo → không phải thanh ngang/dọc
                meta["diagonal"] = abs(o["x1"] - o["x0"]) > 1.0 and abs(o["bottom"] - o["top"]) > 1.0
            objs.append(Obj(kind, rect, meta=meta))
    for im in page.images:
        rect = geom.plumber_rect(im["x0"], im["top"], im["x1"], im["bottom"])
        src = im.get("srcsize") or (0, 0)
        objs.append(Obj("image", rect, meta={"srcsize": list(src)}))
    return objs


# ───────────────────────────── dòng ─────────────────────────────
@dataclass
class Atom:
    kind: str                     # char | frac | rad | complex
    bbox: Rect
    chars: list[Ch] = field(default_factory=list)
    latex: str | None = None
    script: str | None = None     # sup | sub
    reason: str = ""
    anchor: float | None = None   # đường cơ sở dòng chứa nguyên tử (để gắn đúng dòng)

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def x1(self) -> float:
        return self.bbox[2]


@dataclass
class Line:
    atoms: list[Atom]
    bbox: Rect
    size: float
    baseline: float

    @property
    def text(self) -> str:
        return "".join(a.chars[0].text if a.kind == "char" else " " for a in self.atoms)


def _compose_accents(chars: list[Ch]) -> list[Ch]:
    """Dấu rời (dấu kết hợp hoặc dấu dạng spacing đặt chồng lên chữ) → gộp vào chữ cái gốc, rồi NFC."""
    out: list[Ch] = []
    for c in chars:
        t = c.text
        if t and all(unicodedata.combining(x) for x in t) and out:
            out[-1].text += t
            c.merged = True
            continue
        if t in SPACING_ACCENTS and out:
            base = out[-1]
            overlap = min(base.x1, c.x1) - max(base.x0, c.x0)
            if overlap > 0.3 * min(base.x1 - base.x0, c.x1 - c.x0) and c.cy < base.cy:
                base.text += SPACING_ACCENTS[t]
                c.merged = True
                continue
        out.append(c)
    for c in out:
        c.text = unicodedata.normalize("NFC", c.text)
    return out


ACCENT_GLYPHS = set("−→←↔^ˆ˜~¯˙¨⃗⃗̂̃̄̇̈")


def _is_tall(c: Ch) -> bool:
    return bool(MATH_EXT_FONT_RE.search(c.font)) or c.h > 1.6 * c.size


def _build_lines(chars: list[Ch]) -> list[list[Ch]]:
    """Gom ký tự (upright) thành dòng theo ĐƯỜNG CƠ SỞ (gốc glyph từ ma trận chữ), không theo hộp cỡ chữ.
    Chữ nhỏ (chỉ số) và glyph cao của font mở rộng (∫, ∑, ngoặc lớn) được gắn vào dòng phù hợp."""
    if not chars:
        return []
    med = statistics.median(c.size for c in chars)
    big = sorted([c for c in chars if c.size >= 0.8 * med and not _is_tall(c)], key=lambda c: (c.by, c.bx))
    rest = [c for c in chars if not (c.size >= 0.8 * med and not _is_tall(c))]
    lines: list[dict[str, Any]] = []
    for c in big:
        best = None
        for ln in lines[-12:]:
            if abs(ln["by"] - c.by) <= 0.3 * min(c.size, ln["size"]):
                if best is None or abs(ln["by"] - c.by) < abs(best["by"] - c.by):
                    best = ln
        if best is None:
            lines.append({"chars": [c], "by": c.by, "size": c.size})
        else:
            best["chars"].append(c)
    for c in sorted(rest, key=lambda c: (c.by, c.bx)):
        best = None
        bestd = 1e9
        for ln in lines:
            ls = ln["size"]
            if _is_tall(c):
                axis = ln["by"] - 0.3 * ls
                ok = c.top - 0.2 * ls <= axis <= c.bottom + 0.2 * ls
                d = abs((c.top + c.bottom) / 2 - axis)
            else:
                tall_near = any(_is_tall(x) and x.x0 - 1.5 * ls <= c.cx <= x.x1 + 1.5 * ls for x in ln["chars"])
                lo, hi = (1.5, 1.3) if tall_near else (0.8, 0.5)
                ok = ln["by"] - lo * ls <= c.by <= ln["by"] + hi * ls
                d = abs(c.by - ln["by"])
            if not ok:
                continue
            near = min((min(abs(x.x1 - c.x0), abs(c.x1 - x.x0), 0 if x.x0 <= c.cx <= x.x1 else 1e9)
                        for x in ln["chars"]), default=1e9)
            if near < 3.5 * ls and d < bestd:
                best, bestd = ln, d
        if best is None:
            lines.append({"chars": [c], "by": c.by, "size": c.size})
        else:
            best["chars"].append(c)
    out = [sorted(ln["chars"], key=lambda c: c.x0) for ln in lines]
    out.sort(key=lambda cs: (statistics.median(c.by for c in cs), cs[0].x0))
    return out


def _line_base(chars: list[Ch]) -> tuple[float, float]:
    """(cỡ chữ chính, đường cơ sở chính) của một nhóm ký tự."""
    norm = [c for c in chars if not _is_tall(c)] or chars
    mx = max(c.size for c in norm)
    big = [c for c in norm if c.size >= 0.85 * mx] or norm
    return statistics.median(c.size for c in big), statistics.median(c.by for c in big)


def _script_of(c: Ch, base_size: float, base_by: float) -> str | None:
    if c.size >= 0.9 * base_size or _is_tall(c):
        return None
    if c.by < base_by - 0.18 * base_size:
        return "sup"
    if c.by > base_by + 0.08 * base_size:
        return "sub"
    return None


# ───────────────────────────── toán 2D ─────────────────────────────
def _char_latex(c: Ch) -> str:
    t = c.text
    if t in GREEK:
        return "\\" + GREEK[t] + " "
    if t in LATEX_OF:
        return LATEX_OF[t]
    return t


def _linear_latex(chars: list[Ch], base_size: float | None = None, base_by: float | None = None) -> tuple[str | None, str]:
    """Chuỗi LaTeX tuyến tính từ ký tự một dòng (có chỉ số trên/dưới theo đường cơ sở). None nếu không chắc."""
    if not chars:
        return "", ""
    if any(c.unmapped for c in chars):
        return None, "ký tự không ánh xạ được Unicode"
    talls = [c for c in chars if _is_tall(c)]

    def sort_key(c: Ch) -> float:
        # cận trên/dưới đặt chồng lên toán tử lớn (∑, ∫, lim) → xếp ngay SAU toán tử
        for t in talls:
            if c is not t and c.size < 0.9 * t.size and t.x0 - 1 <= c.cx <= t.x1 + 1:
                return t.x0 + 0.001 + (c.x0 - t.x0) * 1e-4
        return c.x0

    chars = sorted(chars, key=sort_key)
    bs, bby = _line_base(chars)
    if base_size is not None:
        bs = base_size
    if base_by is not None:
        bby = base_by
    out: list[str] = []
    i = 0
    while i < len(chars):
        c = chars[i]
        sc = _script_of(c, bs, bby)
        if sc is None:
            word = ""
            j = i
            while j < len(chars) and chars[j].text.isalpha() and not chars[j].mathfont \
                    and _script_of(chars[j], bs, bby) is None and (j == i or chars[j].x0 - chars[j - 1].x1 < 0.15 * bs):
                word += chars[j].text
                j += 1
            if word in FUNC_NAMES:
                out.append("\\" + word + " ")
                i = j
                continue
            out.append(_char_latex(c))
            i += 1
            continue
        group = []
        while i < len(chars) and _script_of(chars[i], bs, bby) == sc:
            group.append(chars[i])
            i += 1
        if any(not (g.text.isalnum() or g.text in "+−-′'*∗=,") for g in group):
            return None, "chỉ số chứa ký hiệu lạ (có thể là dấu/mũi tên đặt trên)"
        inner, why = _linear_latex(group)
        if inner is None:
            return None, why
        out.append(("^{" if sc == "sup" else "_{") + inner + "}")
    return "".join(out), ""


def _find_radicals(chars: list[Ch], objs: list[Obj], page_id: str) -> list[Atom]:
    atoms: list[Atom] = []
    rules = [o for o in objs if o.owner is None and o.thin_h and o.kind in ("rect", "line")]
    for k, c in enumerate(chars):
        if c.owner is not None:
            continue
        is_rad = c.text == "√" or (MATH_EXT_FONT_RE.search(c.font) and c.unmapped and c.h > 0.8 * c.size)
        if not is_rad:
            continue
        vinc = None
        for r in rules:
            if r.owner is None and abs(r.bbox[0] - c.x1) <= 2.5 and abs(r.bbox[1] - c.top) <= 3.0:
                vinc = r
                break
        if vinc is None:
            continue
        ry = vinc.bbox[1]
        under = [x for x in chars if x.owner is None and x is not c and vinc.bbox[0] - 0.5 <= x.cx <= vinc.bbox[2] + 0.5
                 and ry < x.by <= ry + 1.6 * max(x.size, 6)]
        idx = [x for x in chars if x.owner is None and x is not c and x.x1 <= c.x0 + 0.6 * (c.x1 - c.x0)
               and x.x1 >= c.x0 - 2 * x.size and x.size < 0.8 * c.size and x.by < (c.top + c.bottom) / 2]
        region = _union([vinc.bbox, (c.x0, c.top, c.x1, c.bottom)] + [(x.x0, x.top, x.x1, x.bottom) for x in under + idx])
        atom = Atom("rad", region, chars=[c] + under + idx)
        atom.anchor = statistics.median(x.by for x in under) if under else c.bottom
        tag = f"{page_id}-rad{k}"
        for x in atom.chars:
            x.owner = tag
        vinc.owner = tag
        lines = _build_lines(under)
        if len(lines) != 1 or _is_tall(c) and c.h > 2.2 * max(x.size for x in under or [c]):
            atom.kind, atom.reason = "complex", "căn nhiều tầng/căn cao"
        else:
            body, why = _linear_latex(lines[0])
            ix, why2 = _linear_latex(idx) if idx else ("", "")
            if body is None or ix is None:
                atom.kind, atom.reason = "complex", why or why2
            else:
                atom.latex = "\\sqrt" + (f"[{ix}]" if ix else "") + "{" + body + "}"
        atoms.append(atom)
    return atoms


def _find_fractions(chars: list[Ch], objs: list[Obj], page_id: str) -> tuple[list[Atom], list[str]]:
    """Thanh ngang mảnh, tử số có đường cơ sở NGAY trên thanh và mẫu số ngay dưới, cùng được căn giữa theo
    thanh (như TeX/Word dựng phân số). Lồng nhau / nhiều dòng → complex (ảnh)."""
    atoms: list[Atom] = []
    rules = [o for o in objs if o.owner is None and o.thin_h and o.kind in ("rect", "line") and o.w < 300]
    for i, r in enumerate(rules):
        if r.owner is not None:
            continue
        bar = (r.bbox[1] + r.bbox[3]) / 2
        x0, x1 = r.bbox[0] - 1.0, r.bbox[2] + 1.0
        cands = [c for c in chars if c.owner is None and x0 <= c.cx <= x1]
        above = [c for c in cands if 0.02 * c.size <= bar - c.by <= 1.0 * max(c.size, 6)]
        below = [c for c in cands if 0.25 * c.size <= c.by - bar <= 1.45 * max(c.size, 6)]
        if not above or not below:
            continue
        if any(c.x0 < x0 - 2 or c.x1 > x1 + 2 for c in above + below):
            continue

        def centered(cs: list[Ch]) -> bool:
            a, b = min(c.x0 for c in cs), max(c.x1 for c in cs)
            return abs((a + b) / 2 - (r.bbox[0] + r.bbox[2]) / 2) <= 0.25 * max(r.w, 4) + 1.0

        widest = max(max(c.x1 for c in above) - min(c.x0 for c in above), max(c.x1 for c in below) - min(c.x0 for c in below))
        if not (centered(above) and centered(below)) or r.w > widest + 1.2 * max(c.size for c in above + below):
            continue
        region = _union([r.bbox] + [(c.x0, c.top, c.x1, c.bottom) for c in above + below])
        nested = [o for j, o in enumerate(rules) if j != i and o.owner is None and _touch(o.bbox, region) and o.w < r.w - 1]
        atom = Atom("frac", region, chars=above + below)
        atom.anchor = bar + 0.25 * statistics.median(c.size for c in above + below)
        tag = f"{page_id}-frac{i}"
        for c in above + below:
            c.owner = tag
        r.owner = tag
        num_lines = _build_lines(above)
        den_lines = _build_lines(below)
        if nested or len(num_lines) > 1 or len(den_lines) > 1:
            atom.kind = "complex"
            atom.reason = "phân số lồng/nhiều dòng"
            for o in nested:
                o.owner = tag
        else:
            num, w1 = _linear_latex(num_lines[0])
            den, w2 = _linear_latex(den_lines[0])
            if num is None or den is None:
                atom.kind = "complex"
                atom.reason = w1 or w2
            else:
                atom.latex = "\\frac{" + num + "}{" + den + "}"
        atoms.append(atom)
    return atoms, []


_ACCENT_LATEX = {"→": "\\overrightarrow", "−→": "\\overrightarrow", "←": "\\overleftarrow", "←−": "\\overleftarrow",
                 "↔": "\\overleftrightarrow", "⃗": "\\vec", "⃗": "\\vec", "^": "\\widehat", "ˆ": "\\widehat",
                 "̂": "\\widehat", "˜": "\\widetilde", "~": "\\widetilde", "̃": "\\widetilde",
                 "¯": "\\overline", "−": "\\overline", "̄": "\\overline", "˙": "\\dot", "̇": "\\dot",
                 "¨": "\\ddot", "̈": "\\ddot"}


def _find_accents(chars: list[Ch], objs: list[Obj], page_id: str) -> list[Atom]:
    """Dấu/mũi tên đặt TRÊN chữ (vectơ \\overrightarrow{AB}, mũ, gạch trên): glyph trong ACCENT_GLYPHS có
    đường cơ sở cao hơn chữ bên dưới 0.35–1.4 cỡ chữ và chồng lấn theo chiều ngang; hoặc thanh mảnh ngắn
    ngay trên chữ (\\overline)."""
    atoms: list[Atom] = []
    acc = [c for c in chars if c.owner is None and c.text in ACCENT_GLYPHS]
    used: set[int] = set()
    for i, a in enumerate(sorted(acc, key=lambda c: c.x0)):
        if id(a) in used:
            continue
        group = [a]
        for b in sorted(acc, key=lambda c: c.x0):
            if id(b) in used or b is a or b in group:
                continue
            if abs(b.by - a.by) <= 0.2 * a.size and b.x0 <= group[-1].x1 + 1.0 and b.x0 >= group[-1].x0:
                group.append(b)
        gx0, gx1 = min(c.x0 for c in group), max(c.x1 for c in group)
        base = [c for c in chars if c.owner is None and c not in group and c.text not in ACCENT_GLYPHS
                and gx0 - 1.5 <= c.cx <= gx1 + 1.5 and 0.35 * c.size <= c.by - a.by <= 1.4 * c.size]
        if not base:
            continue
        for c in group:
            used.add(id(c))
        key = "".join(c.text for c in group)
        if key not in _ACCENT_LATEX:
            key = "−→" if group[-1].text == "→" and all(c.text in "−→" for c in group) else key
        cmd = _ACCENT_LATEX.get(key)
        region = _union([(c.x0, c.top, c.x1, c.bottom) for c in group + base])
        atom = Atom("acc", region, chars=group + base)
        atom.anchor = statistics.median(c.by for c in base)
        tag = f"{page_id}-acc{i}"
        for c in group + base:
            c.owner = tag
        inner, why = _linear_latex(base)
        if cmd is None or inner is None:
            atom.kind, atom.reason = "complex", why or f"dấu trên chưa nhận dạng: {key!r}"
        else:
            if cmd == "\\overrightarrow" and len([c for c in base if c.text.strip()]) == 1 and key == "→":
                cmd = "\\vec"
            if cmd == "\\widehat" and len(base) == 1:
                cmd = "\\hat"
            if cmd == "\\widetilde" and len(base) == 1:
                cmd = "\\tilde"
            if cmd == "\\overline" and len(base) == 1:
                cmd = "\\bar"
            atom.latex = cmd + "{" + inner + "}"
        atoms.append(atom)
    # gạch trên bằng thanh mảnh ngắn
    for k, r in enumerate(o for o in objs if o.owner is None and o.thin_h and o.kind in ("rect", "line") and o.w < 150):
        ry = (r.bbox[1] + r.bbox[3]) / 2
        base = [c for c in chars if c.owner is None and r.bbox[0] - 1 <= c.cx <= r.bbox[2] + 1
                and 0.55 * c.size <= c.by - ry <= 1.15 * c.size]
        above = [c for c in chars if c.owner is None and r.bbox[0] - 1 <= c.cx <= r.bbox[2] + 1 and 0 < ry - c.by < 1.2 * c.size]
        if not base or above:
            continue
        bx0, bx1 = min(c.x0 for c in base), max(c.x1 for c in base)
        if abs(bx0 - r.bbox[0]) > 0.6 * base[0].size or abs(bx1 - r.bbox[2]) > 0.6 * base[0].size:
            continue
        inner, why = _linear_latex(base)
        tag = f"{page_id}-ovl{k}"
        region = _union([r.bbox] + [(c.x0, c.top, c.x1, c.bottom) for c in base])
        atom = Atom("acc", region, chars=list(base))
        atom.anchor = statistics.median(c.by for c in base)
        for c in base:
            c.owner = tag
        r.owner = tag
        if inner is None:
            atom.kind, atom.reason = "complex", why
        else:
            atom.latex = "\\overline{" + inner + "}"
        atoms.append(atom)
    return atoms


# ───────────────────────────── phân đoạn chữ/toán trong dòng ─────────────────────────────
def _is_mathish_char(c: Ch) -> bool:
    return c.mathfont or c.text in MATH_CHARS


def _line_atoms(line_chars: list[Ch], struct_atoms: list[Atom]) -> Line:
    """Ghép ký tự dòng + nguyên tử 2D (phân số/căn/dấu trên) nằm trong dải dọc của dòng."""
    size, base = _line_base(line_chars) if line_chars else (10.0, struct_atoms[0].bbox[3] if struct_atoms else 0)
    atoms = [Atom("char", (c.x0, c.top, c.x1, c.bottom), chars=[c]) for c in line_chars]
    atoms += struct_atoms
    atoms.sort(key=lambda a: a.x0)
    bbox = _union([a.bbox for a in atoms])
    return Line(atoms, bbox, size, base)


def _segment_line(line: Line) -> list[dict[str, Any]]:
    """Chia dòng thành đoạn {"kind": "text"|"math", "atoms": [...]}. Heuristic: font toán, ký hiệu toán,
    chỉ số (theo đường cơ sở), cấu trúc 2D; chữ số/dấu câu trung tính theo ngữ cảnh; dấu câu cuối cụm toán
    in bằng font chữ thường được trả về đoạn chữ."""
    size = line.size
    words: list[list[Atom]] = []
    cur: list[Atom] = []
    prev = None
    for a in line.atoms:
        if prev is not None and a.x0 - prev.x1 > 0.22 * size:
            words.append(cur)
            cur = []
        cur.append(a)
        prev = a
    if cur:
        words.append(cur)

    def classify(word: list[Atom]) -> str:
        chars = [a.chars[0] for a in word if a.kind == "char"]
        if any(a.kind != "char" for a in word):
            return "math"
        if any(_script_of(c, size, line.baseline) for c in chars):
            return "math"
        if any(_is_mathish_char(c) and c.text not in "|<>" for c in chars):
            letters = [c for c in chars if c.text.isalpha() and not c.mathfont]
            if len(letters) >= 3:
                return "text"
            return "math"
        txt = "".join(c.text for c in chars)
        if re.fullmatch(r"[0-9.,;:()\[\]/%]+", txt):
            return "neutral"
        return "text"

    kinds = [classify(w) for w in words]

    def ends_with_operator(word: list[Atom]) -> bool:
        last = word[-1]
        return last.kind == "char" and last.chars[0].text in "+−-=×÷±≤≥≠<>(,;:/·⋅∈⊂∪∩→⇒⇔"

    def starts_with_operator(word: list[Atom]) -> bool:
        first = word[0]
        return first.kind == "char" and first.chars[0].text in "+−-=×÷±≤≥≠<>)/·⋅∈⊂∪∩→⇒⇔^"

    for i, k in enumerate(kinds):
        if k == "neutral":
            left = kinds[i - 1] if i > 0 else None
            right = kinds[i + 1] if i + 1 < len(kinds) else None
            if (left == "math" and right == "math") or (left == "math" and ends_with_operator(words[i - 1])) or \
                    (right == "math" and starts_with_operator(words[i + 1])):
                kinds[i] = "math"
            else:
                kinds[i] = "text"
    # dấu câu cuối từ toán (font chữ thường) → chữ
    split_words: list[list[Atom]] = []
    split_kinds: list[str] = []
    for w, k in zip(words, kinds):
        if k == "math" and len(w) > 1 and w[-1].kind == "char" and w[-1].chars[0].text in ".,;:" \
                and not w[-1].chars[0].mathfont:
            split_words += [w[:-1], [w[-1]]]
            split_kinds += ["math", "text"]
        else:
            split_words.append(w)
            split_kinds.append(k)
    segs: list[dict[str, Any]] = []
    for w, k in zip(split_words, split_kinds):
        if segs and segs[-1]["kind"] == k:
            segs[-1]["atoms"].extend(w)
            segs[-1]["words"] += 1
        else:
            segs.append({"kind": k, "atoms": list(w), "words": 1})
    return segs


def _math_seg_latex(atoms: list[Atom], size: float) -> tuple[str | None, str]:
    """LaTeX của một cụm toán: ký tự (tuyến tính, có chỉ số) xen phân số/căn/dấu trên. complex → None."""
    out: list[str] = []
    buf: list[Ch] = []
    char_atoms = [a.chars[0] for a in atoms if a.kind == "char"]
    bs, bby = _line_base(char_atoms) if char_atoms else (size, None)

    def flush() -> str | None:
        nonlocal buf
        if not buf:
            return ""
        s, _why = _linear_latex(buf, bs, bby)
        buf = []
        return s

    for a in atoms:
        if a.kind == "char":
            buf.append(a.chars[0])
            continue
        s = flush()
        if s is None:
            return None, "ký tự không ánh xạ/chỉ số lạ"
        out.append(s)
        if a.kind == "complex" or a.latex is None:
            return None, a.reason or "cấu trúc 2D chưa hỗ trợ"
        out.append(a.latex)
    s = flush()
    if s is None:
        return None, "ký tự không ánh xạ/chỉ số lạ"
    out.append(s)
    latex = "".join(out).strip()
    if latex.count("{") - latex.count("\\{") != latex.count("}") - latex.count("\\}"):
        return None, "ngoặc không cân"
    return latex, ""


def _classify_page(chars: list[Ch], objs: list[Obj], geom: PageGeom) -> dict[str, Any]:
    area = geom.width * geom.height
    vis = [c for c in chars if not c.hidden and c.text.strip()]
    hid = [c for c in chars if c.hidden and c.text.strip()]
    imgs = [o for o in objs if o.kind == "image"]
    img_cover = sum(_area(_inter(o.bbox, (0, 0, geom.width, geom.height)) or (0, 0, 0, 0)) for o in imgs) / max(1, area)
    vec = [o for o in objs if o.kind != "image"]
    unm = [c for c in vis if c.unmapped]
    legacy_vn = sorted({c.font for c in vis if re.search(r"(^|\+)\.?Vn[A-Z]|VNI[-_]", c.font)})
    info = {"visible_chars": len(vis), "hidden_chars": len(hid), "image_cover": round(min(1.0, img_cover), 3),
            "vector_objects": len(vec), "unmapped_chars": len(unm), "legacy_vn_fonts": legacy_vn}
    if img_cover >= 0.6 and len(vis) < 20:
        info["type"] = "scan_with_hidden_text" if len(hid) >= 20 else "scan"
    elif not vis and not hid:
        info["type"] = "figure_only" if (vec or imgs) else "empty"
    elif len(vis) < 20 and (vec or imgs):
        info["type"] = "figure_only"
    elif vis and len(unm) / max(1, len(vis)) > 0.3:
        info["type"] = "broken_text"
    elif img_cover >= 0.2:
        info["type"] = "mixed"
    else:
        info["type"] = "text"
    return info


def _detect_columns(lines: list[Line], width: float) -> dict[str, Any] | None:
    """Tìm khe dọc ở giữa trang mà gần như không dòng nào cắt qua (≥3 dòng mỗi bên)."""
    if len(lines) < 6:
        return None
    step = 2.0
    nb = int(width / step) + 1
    cover = [0] * nb
    for ln in lines:
        for k in range(max(0, int(ln.bbox[0] / step)), min(nb, int(ln.bbox[2] / step) + 1)):
            cover[k] += 1
    lo, hi = int(0.28 * width / step), int(0.72 * width / step)
    best = None
    k = lo
    while k < hi:
        if cover[k] <= max(1, 0.08 * len(lines)):
            j = k
            while j < hi and cover[j] <= max(1, 0.08 * len(lines)):
                j += 1
            if best is None or (j - k) > (best[1] - best[0]):
                best = (k, j)
            k = j
        k += 1
    if not best or (best[1] - best[0]) * step < 6:
        return None
    gx0, gx1 = best[0] * step, best[1] * step
    left = [ln for ln in lines if ln.bbox[2] <= gx0 + 1]
    right = [ln for ln in lines if ln.bbox[0] >= gx1 - 1]
    if len(left) < 3 or len(right) < 3:
        return None
    return {"gutter": [round(gx0, 1), round(gx1, 1)], "left_lines": len(left), "right_lines": len(right)}


_PAGE_NO_RE = re.compile(r"^(?:[-–—]\s*)?(?:(?:Trang|Page|tr\.)\s*)?\d{1,4}(?:\s*/\s*\d{1,4})?(?:\s*[-–—])?$",
                         re.IGNORECASE)


def _split_at_gutter(lines: list[Line], width: float) -> tuple[list[Line], dict[str, Any] | None]:
    """Trang 2 cột thường có ĐƯỜNG CƠ SỞ THẲNG HÀNG giữa hai cột (LaTeX twocolumn) → bộ gom dòng nối chữ hai cột
    thành một dòng. Tìm khe cột trên các ĐOẠN dòng (tách ở khoảng trống lớn), rồi tách những dòng vắt qua khe."""
    segs: list[Line] = []
    cuts: list[list[int]] = []
    for ln in lines:
        atoms = sorted(ln.atoms, key=lambda a: a.bbox[0])
        idx = [k for k in range(1, len(atoms))
               if atoms[k].bbox[0] - atoms[k - 1].bbox[2] > max(2.5 * ln.size, 15.0)]
        cuts.append(idx)
        start = 0
        for k in idx + [len(atoms)]:
            part = atoms[start:k]
            if part:
                segs.append(Line(part, _union([a.bbox for a in part]), ln.size, ln.baseline))
            start = k
    cols = _detect_columns(segs, width)
    if not cols:
        return lines, None
    g0, g1 = cols["gutter"]
    out: list[Line] = []
    split_n = 0
    for ln in lines:
        atoms = sorted(ln.atoms, key=lambda a: a.bbox[0])
        k_cut = None
        for k in range(1, len(atoms)):
            if atoms[k - 1].bbox[2] <= g0 + 1.0 and atoms[k].bbox[0] >= g1 - 1.0:
                k_cut = k
                break
        if k_cut is None:
            out.append(ln)
            continue
        for part in (atoms[:k_cut], atoms[k_cut:]):
            out.append(Line(part, _union([a.bbox for a in part]), ln.size, ln.baseline))
        split_n += 1
    out.sort(key=lambda ln: (ln.bbox[1], ln.bbox[0]))
    cols = dict(cols, lines_split_at_gutter=split_n)
    return out, cols


# ───────────────────────────── chính ─────────────────────────────
def analyze_pdf(pdf_path: Path, work_dir: Path, *, dpi: int = 400, ocr: str = "auto", ocr_lang: str | None = None,
                max_pages: int | None = None) -> dict[str, Any]:
    import pdfplumber
    import pikepdf
    pdf_path = Path(pdf_path)
    work_dir.mkdir(parents=True, exist_ok=True)
    assets = work_dir / "assets"
    max_pages = max_pages or int(__import__("os").environ.get("L2W_MAX_PDF_PAGES", "300"))
    doc_info: dict[str, Any] = {}
    try:
        pk = pikepdf.open(str(pdf_path))
    except pikepdf.PasswordError as exc:
        raise PdfInputError("PDF được bảo vệ bằng mật khẩu — không mở được. Hãy gỡ mật khẩu rồi tải lại.",
                            "pdf_password") from exc
    except Exception as exc:
        raise PdfInputError(f"PDF hỏng hoặc không đọc được: {type(exc).__name__}: {exc}", "pdf_broken") from exc
    with pk:
        doc_info["encrypted"] = bool(pk.is_encrypted)
        doc_info["pdf_version"] = str(pk.pdf_version)
        try:
            doc_info["producer"] = str(pk.docinfo.get("/Producer", ""))[:120]
            doc_info["creator"] = str(pk.docinfo.get("/Creator", ""))[:120]
        except Exception:
            pass
        n_pages = len(pk.pages)
        if n_pages == 0:
            raise PdfInputError("PDF không có trang nào.", "pdf_empty")
        if n_pages > max_pages:
            raise PdfInputError(f"PDF có {n_pages} trang, vượt giới hạn {max_pages} (L2W_MAX_PDF_PAGES).",
                                "pdf_too_many_pages")
        geoms = []
        for pg in pk.pages:
            mb = pg.mediabox
            try:
                cb = pg.cropbox
            except Exception:
                cb = mb
            rot = pg.obj.get("/Rotate", None)
            if rot is None:
                # /Rotate kế thừa
                par = pg.obj.get("/Parent")
                while par is not None and rot is None:
                    rot = par.get("/Rotate")
                    par = par.get("/Parent")
            geoms.append(PageGeom.create([float(v) for v in mb], [float(v) for v in cb], int(rot or 0)))
        doc_info["form_xobjects"] = sum(
            1 for pg in pk.pages for _n, xo in (pg.resources.get("/XObject", {}) or {}).items()
            if getattr(xo, "get", lambda k: None)("/Subtype") == "/Form") if pk.pages else 0
    from .capabilities import probe_capabilities
    caps = probe_capabilities()
    ocr_cap = caps["derived"]["ocr"]
    langs = caps.get("tesseract_languages", [])
    manifest: dict[str, Any] = {
        "schema": "l2w-pdf-manifest/1",
        "coordinate_system": "pt; gốc trên-trái trang HIỂN THỊ (đã áp CropBox và /Rotate); y tăng xuống",
        "source": {"file": pdf_path.name}, "document": doc_info, "page_count": n_pages,
        "pages": [], "page_types": {}, "blocks": [], "issues": [], "summary": {},
    }
    counter = {"order": 0}

    def add_block(b: dict[str, Any]) -> dict[str, Any]:
        counter["order"] += 1
        b.setdefault("order", counter["order"])
        b.setdefault("warnings", [])
        b.setdefault("status", "ok")
        b.setdefault("parent", None)
        b.setdefault("fallback_reason", None)
        b.setdefault("confidence", None)
        b["source"] = {"file": pdf_path.name, "page": b["page"], "bbox_pt": b.get("bbox"),
                       "coords": "display_topleft_pt"}
        manifest["blocks"].append(b)
        return b

    def issue(sev: str, code: str, msg: str, **kw: Any) -> None:
        manifest["issues"].append({"severity": sev, "code": code, "message": msg, **kw})

    with pdfplumber.open(str(pdf_path)) as pl:
        for pi, page in enumerate(pl.pages):
            geom = geoms[pi]
            pno = pi + 1
            pid = f"p{pno}"
            tight: list[tuple[str, Rect]] = []
            try:
                hidden_boxes, bad_boxes = _pdfium_char_flags(pdf_path, pi, geom, tight)
            except Exception as exc:
                hidden_boxes, bad_boxes = [], []
                issue("warning", "pdfium_text", f"Trang {pno}: không đọc được cờ chữ ẩn ({type(exc).__name__}).", page=pno)
            chars = _load_chars(page, geom, hidden_boxes, bad_boxes, tight)
            objs = _load_objects(page, geom)
            # chỉ giữ những gì HIỂN THỊ: ký tự/đối tượng nằm ngoài CropBox (sau xoay) không được đưa vào Word
            n0 = len(chars)
            chars = [c for c in chars if -1.0 <= c.cx <= geom.width + 1.0 and -1.0 <= c.cy <= geom.height + 1.0]
            objs = [o for o in objs if o.bbox[2] >= -1.0 and o.bbox[0] <= geom.width + 1.0
                    and o.bbox[3] >= -1.0 and o.bbox[1] <= geom.height + 1.0]
            if len(chars) < n0:
                issue("info", "outside_cropbox", f"Trang {pno}: bỏ {n0 - len(chars)} ký tự nằm ngoài vùng hiển thị "
                      "(CropBox) — không hiển thị trong trình xem PDF nên không đưa vào Word.", page=pno)
            pinfo = _classify_page(chars, objs, geom)
            pinfo.update({"page": pno, **geom.to_dict()})
            manifest["pages"].append(pinfo)
            manifest["page_types"][str(pno)] = pinfo["type"]
            ptype = pinfo["type"]
            if pinfo["legacy_vn_fonts"]:
                issue("warning", "legacy_vn_font",
                      f"Trang {pno}: font bảng mã cũ ({', '.join(pinfo['legacy_vn_fonts'][:3])}) — chưa giải mã "
                      "VNI/TCVN3 (cần bảng mã đã kiểm chứng); dòng chứa font này giữ ẢNH vùng.", page=pno)
            if ptype == "empty":
                add_block({"block_id": f"{pid}-empty", "type": "page", "page": pno, "bbox": [0, 0, geom.width, geom.height],
                           "method": "text_native", "text": "", "runs": [], "status": "ok",
                           "warnings": ["trang trống"]})
                continue
            if ptype in ("scan", "scan_with_hidden_text"):
                _handle_scan_page(pdf_path, pi, geom, chars, ptype, pid, pno, assets, dpi, ocr, ocr_lang, ocr_cap,
                                  langs, add_block, issue)
                continue
            # chữ ẩn trùng chữ hiện → bỏ bản ẩn (tránh lặp chữ)
            vis = [c for c in chars if not c.hidden]
            hid = [c for c in chars if c.hidden]
            dup = 0
            if hid and vis:
                vis_boxes = [(c.x0, c.top, c.x1, c.bottom) for c in vis]
                for c in hid:
                    if _match_any((c.x0, c.top, c.x1, c.bottom), vis_boxes, 1.0):
                        c.owner = "duplicate-hidden"
                        dup += 1
                if dup:
                    issue("info", "hidden_text_duplicate", f"Trang {pno}: bỏ {dup} ký tự chữ ẩn trùng chữ hiển thị.",
                          page=pno)
            hid_left = [c for c in hid if c.owner is None]
            if hid_left:
                issue("warning", "hidden_text", f"Trang {pno}: {len(hid_left)} ký tự chữ ẩn (lớp OCR) không trùng chữ "
                      "hiển thị → không đưa vào Word (nội dung đã có ở ảnh/chữ hiển thị).", page=pno)
                for c in hid_left:
                    c.owner = "hidden"
            _analyze_text_page(pdf_path, pi, geom, page, chars, objs, pid, pno, assets, dpi, add_block, issue, pinfo)
    # tổng hợp
    by_method = Counter(b["method"] for b in manifest["blocks"])
    by_type = Counter(b["type"] for b in manifest["blocks"])
    manifest["summary"] = {
        "blocks": len(manifest["blocks"]), "by_method": dict(by_method), "by_type": dict(by_type),
        "page_types": dict(Counter(manifest["page_types"].values())),
        "unresolved": sum(1 for b in manifest["blocks"] if b["method"] == "unresolved"),
        "inline_math_omml": sum(len(b.get("inline_math") or []) for b in manifest["blocks"]),
        "inline_math_images": sum(len(b.get("inline_math_images") or []) for b in manifest["blocks"]),
        "confidence_note": "confidence là heuristic nội bộ (0–1), chưa hiệu chuẩn trên tập dữ liệu thật.",
    }
    (work_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1, default=str),
                                            encoding="utf-8")
    return manifest


def _handle_scan_page(pdf_path, pi, geom, chars, ptype, pid, pno, assets, dpi, ocr, ocr_lang, ocr_cap, langs,
                      add_block, issue) -> None:
    full = (0.0, 0.0, geom.width, geom.height)
    page_dpi = min(dpi, 300)
    use_ocr = ocr != "off" and ocr_cap.get("available")
    lang = ocr_lang
    if use_ocr and not lang:
        if "vie" in langs:
            lang = "vie+eng" if "eng" in langs else "vie"
        else:
            use_ocr = False
            issue("warning", "ocr_lang_missing",
                  f"Trang {pno} là trang scan; máy có tesseract nhưng KHÔNG có gói tiếng Việt (vie) → không OCR "
                  "(giữ ảnh trang). Đặt ocr_lang=eng nếu tài liệu tiếng Anh.", page=pno)
    if ocr != "off" and not ocr_cap.get("available"):
        issue("warning", "ocr_unavailable", f"Trang {pno} là trang scan; máy không có OCR → giữ ảnh trang.", page=pno)
    if ptype == "scan_with_hidden_text":
        issue("info", "scan_has_ocr_layer", f"Trang {pno}: scan có lớp chữ OCR ẩn sẵn — không dùng làm chữ chính "
              "(chất lượng không kiểm chứng được); giữ ảnh trang.", page=pno)
    if use_ocr and lang:
        try:
            res = _ocr_page(pdf_path, pi, geom, lang, pid, pno)
        except Exception as exc:
            res = None
            issue("warning", "ocr_failed", f"Trang {pno}: OCR lỗi ({type(exc).__name__}: {exc}) → giữ ảnh trang.", page=pno)
        if res and res["mean_conf"] >= 75 and res["words"] >= 5:
            for b in res["blocks"]:
                b["status"] = "review"
                b["warnings"].append(f"chữ OCR ({lang}), độ tin cậy trung bình {res['mean_conf']:.0f} — cần soát")
                b["confidence"] = {"value": round(res["mean_conf"] / 100, 2), "kind": "tesseract_conf_mean"}
                add_block(b)
            return
        if res:
            issue("warning", "ocr_low_confidence",
                  f"Trang {pno}: OCR ({lang}) tin cậy thấp ({res['mean_conf']:.0f}) → giữ ảnh trang.", page=pno)
    crop = crop_region(pdf_path, pi, geom, full, assets, f"{pid}-page", dpi=page_dpi, want_svg=False)
    add_block({"block_id": f"{pid}-page", "type": "page", "page": pno, "bbox": _r(full), "method": "page_image",
               "status": "fallback", "assets": crop, "text": "",
               "fallback_reason": "trang scan, không có OCR đủ tin cậy → giữ ảnh trang (khả năng sửa giảm)"})


def _ocr_page(pdf_path, pi, geom, lang, pid, pno) -> dict[str, Any]:
    import pypdfium2 as pdfium
    import pytesseract
    from .capabilities import find_tool
    _tess = find_tool("tesseract")
    if _tess:
        pytesseract.pytesseract.tesseract_cmd = _tess
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        img = doc[pi].render(scale=300 / 72).to_pil().convert("L")
    finally:
        doc.close()
    data = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DICT, config="--psm 3")
    k = 72 / 300
    groups: dict[tuple, list[int]] = defaultdict(list)
    confs = []
    for i, w in enumerate(data["text"]):
        if not (w or "").strip():
            continue
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if conf < 0:
            continue
        confs.append(conf)
        groups[(data["block_num"][i], data["par_num"][i])].append(i)
    blocks = []
    for n, (key, idxs) in enumerate(sorted(groups.items(), key=lambda kv: min(data["top"][i] for i in kv[1]))):
        lines: dict[int, list[int]] = defaultdict(list)
        for i in idxs:
            lines[data["line_num"][i]].append(i)
        text = " ".join(" ".join(data["text"][i] for i in sorted(v, key=lambda i: data["left"][i]))
                        for _ln, v in sorted(lines.items()))
        x0 = min(data["left"][i] for i in idxs) * k
        y0 = min(data["top"][i] for i in idxs) * k
        x1 = max(data["left"][i] + data["width"][i] for i in idxs) * k
        y1 = max(data["top"][i] + data["height"][i] for i in idxs) * k
        text = unicodedata.normalize("NFC", text)
        blocks.append({"block_id": f"{pid}-ocr{n + 1}", "type": "paragraph", "page": pno, "bbox": _r((x0, y0, x1, y1)),
                       "method": "text_native", "text": text, "runs": [{"text": text}], "ocr": True, "warnings": []})
    return {"blocks": blocks, "mean_conf": (sum(confs) / len(confs)) if confs else 0.0, "words": len(confs)}


def _analyze_text_page(pdf_path, pi, geom, page, chars, objs, pid, pno, assets, dpi, add_block, issue, pinfo) -> None:
    W, H = geom.width, geom.height
    live = [c for c in chars if c.owner is None and c.text.strip()]
    blocks: list[dict[str, Any]] = []
    # 1) bảng có đường kẻ (pdfplumber), loại bỏ khung trang
    tables = []
    try:
        for t in page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"}):
            rect = geom.plumber_rect(*t.bbox)
            if _area(rect) > 0.85 * W * H:
                continue
            cells = [geom.plumber_rect(*c) for c in t.cells if c]
            if len(cells) < 2:
                continue
            tables.append((rect, cells))
    except Exception as exc:
        issue("warning", "table_detect", f"Trang {pno}: dò bảng lỗi ({type(exc).__name__}).", page=pno)
    ti = 0
    for rect, cells in tables:
        inside = [c for c in live if c.owner is None and _contains(_expand(rect, 1), c.cx, c.cy)]
        objs_in = [o for o in objs if o.owner is None and _touch(o.bbox, _expand(rect, 1))]
        diag = [o for o in objs_in if o.kind == "curve" or o.meta.get("diagonal")]
        filled = sum(1 for cell in cells if any(_contains(cell, c.cx, c.cy) for c in inside))
        if filled < 0.4 * len(cells):
            # lưới gần như không có chữ (vd. lưới đồ thị) → không phải bảng; để bước hình xử lý
            continue
        ti += 1
        bid = f"{pid}-t{ti}"
        for c in inside:
            c.owner = bid
        for o in objs_in:
            inter = _inter(o.bbox, rect)
            if o.thin_h or o.thin_v or o.w < 0.5 or o.h < 0.5 or (
                    inter and _area(inter) >= 0.8 * max(_area(o.bbox), 0.01)):
                o.owner = bid
        grid = _table_grid(cells, inside)
        if diag or grid is None or any(c.unmapped for c in inside) or any(not c.upright for c in inside):
            reason = ("bảng có mũi tên/đường cong (vd. bảng biến thiên)" if diag else
                      "lưới ô không đều" if grid is None else "ký tự không ánh xạ/chữ xoay trong bảng")
            crop = crop_region(pdf_path, pi, geom, _expand(rect, 2), assets, bid, dpi=dpi, want_svg=True)
            blocks.append({"block_id": bid, "type": "table", "page": pno, "bbox": _r(rect),
                           "method": "vector_image" if crop.get("svg") else "raster_image", "status": "fallback",
                           "assets": crop, "fallback_reason": "table_as_image: " + reason,
                           "text": " ".join(c.text for c in sorted(inside, key=lambda c: (c.top, c.x0)))})
            continue
        rows = []
        flat_text = []
        for r_cells in grid["rows"]:
            row = []
            for cell in r_cells:
                cch = [c for c in inside if _contains(cell["bbox"], c.cx, c.cy)]
                paras = _cell_paragraphs(cch)
                if cch:
                    lg = min(c.x0 for c in cch) - cell["bbox"][0]
                    rg = cell["bbox"][2] - max(c.x1 for c in cch)
                    cell["align"] = ("center" if abs(lg - rg) <= max(2.0, 0.25 * (lg + rg)) else
                                     "right" if rg < lg else "left")
                row.append({**cell, "bbox": _r(cell["bbox"]), "paragraphs": paras})
                flat_text += [p["text"] for p in paras]
            rows.append(row)
        blocks.append({"block_id": bid, "type": "table", "page": pno, "bbox": _r(rect), "method": "table_native",
                       "rows": rows, "n_rows": grid["n_rows"], "n_cols": grid["n_cols"], "text": " ".join(flat_text),
                       "col_widths_pt": [round(grid["xs"][k + 1] - grid["xs"][k], 2) for k in range(grid["n_cols"])],
                       "confidence": {"value": 0.8, "kind": "heuristic"},
                       "inline_math": [m for r in rows for c in r for p in c["paragraphs"] for m in p.get("math", [])]})
    # 2) gạch chân: thanh mảnh ngay dưới chữ, không có chữ bên dưới
    for o in objs:
        if o.owner is None and o.thin_h:
            above = [c for c in live if c.owner is None and o.bbox[0] - 1 <= c.cx <= o.bbox[2] + 1
                     and 0 <= o.bbox[1] - c.bottom <= 0.45 * c.size + 1.5]
            below = [c for c in live if c.owner is None and o.bbox[0] - 1 <= c.cx <= o.bbox[2] + 1
                     and 0 <= c.top - o.bbox[3] <= 1.2 * c.size]
            if above and not below:
                o.owner = "underline"
                o.meta["underline_of"] = len(above)
    # 3) phân số & căn (trước khi gom dòng)
    rads = _find_radicals(live, objs, pid)
    fracs, _ = _find_fractions(live, objs, pid)
    accs = _find_accents(live, objs, pid)
    struct_atoms = rads + fracs + accs
    # 4) hình: đối tượng vector/ảnh còn lại → cụm
    fig_objs = [o for o in objs if o.owner is None and not (o.thin_h and o.w > 0.5 * W)]
    clusters = _cluster(fig_objs, gap=6.0)
    figures = []
    for cl in clusters:
        rect = _union(o.bbox for o in cl)
        if _area(rect) < 40 and len(cl) < 3:
            for o in cl:
                o.owner = "decoration"
            continue
        inside_chars = [c for c in live if c.owner is None and _contains(_expand(rect, 1), c.cx, c.cy)]
        # khung bao chữ (tcolorbox, viền trang) — ít đối tượng, viền trùng khung, nhiều dòng chữ bên trong
        border_like = [o for o in cl if o.kind in ("rect", "curve", "line") and
                       (abs(o.bbox[0] - rect[0]) < 3 or abs(o.bbox[2] - rect[2]) < 3)]
        if len(inside_chars) > 40 and len(cl) <= 8 and len(border_like) >= 1 and \
                not any(o.kind == "image" for o in cl):
            n_lines = len(_build_lines(inside_chars))
            if n_lines >= 2:
                for o in cl:
                    o.owner = "frame"
                issue("info", "frame_ignored", f"Trang {pno}: khung trang trí quanh chữ (vd. tcolorbox) — giữ chữ, bỏ khung.",
                      page=pno)
                continue
        figures.append((rect, cl))
    for fi, (rect, cl) in enumerate(figures):
        bid = f"{pid}-f{fi + 1}"
        # nhãn: chữ nhỏ trong/sát hình mà dòng của nó không chạy dài ra ngoài hình
        region = rect
        for _pass in range(3):
            near_sz = statistics.median([c.size for c in live if c.owner is None] or [10.0])
            reach = _expand(region, max(8.0, 1.6 * near_sz))
            cand = [c for c in live if c.owner is None and _contains(reach, c.cx, c.cy)]
            if not cand:
                break
            grown = False
            for c in [c for c in cand if not c.upright]:
                c.owner = bid
                region = _union([region, (c.x0, c.top, c.x1, c.bottom)])
                grown = True
            for ln in _build_lines([c for c in cand if c.upright]):
                lb = _union((c.x0, c.top, c.x1, c.bottom) for c in ln)
                inside_fig = _contains(_expand(rect, 1), (lb[0] + lb[2]) / 2, (lb[1] + lb[3]) / 2)
                short = (lb[2] - lb[0]) < max(40.0, 0.35 * (rect[2] - rect[0]))
                fully_near = _contains(reach, lb[0], lb[1]) and _contains(reach, lb[2], lb[3])
                within_x = lb[0] >= reach[0] and lb[2] <= reach[2]
                ink_w = sum(c.x1 - c.x0 for c in ln)
                sparse = ink_w < 0.5 * max(1.0, lb[2] - lb[0])     # vd. hàng số chia trục
                # dòng dài ngoài hình (đề bài/chú thích cạnh hình) → không phải nhãn
                if (inside_fig or short or (within_x and sparse)) and fully_near:
                    for c in ln:
                        c.owner = bid
                    region = _union([region, lb])
                    grown = True
            if not grown:
                break
        for a in struct_atoms:
            if _inter(a.bbox, rect):
                for c in a.chars:
                    c.owner = bid
                a.kind = "owned"
                region = _union([region, a.bbox])
        for o in cl:
            o.owner = bid
        has_raster = any(o.kind == "image" for o in cl)
        has_vector = any(o.kind != "image" for o in cl)
        eff_dpi = dpi
        if has_raster and not has_vector:
            im = max((o for o in cl if o.kind == "image"), key=lambda o: _area(o.bbox))
            sw = (im.meta.get("srcsize") or [0])[0] or 0
            if sw and im.w > 0:
                eff = sw / (im.w / 72.0)
                eff_dpi = int(max(96, min(dpi, eff * 1.05)))
        crop = crop_region(pdf_path, pi, geom, _expand(region, 2.5), assets, bid, dpi=eff_dpi,
                           want_svg=has_vector and not has_raster)
        method = "vector_image" if crop.get("svg") else "raster_image"
        warnings = []
        if has_raster and eff_dpi < dpi:
            warnings.append(f"ảnh gốc độ phân giải thấp (~{eff_dpi} dpi hiệu dụng) — không tăng DPI giả")
        blocks.append({"block_id": bid, "type": "figure" if has_vector else "image", "page": pno,
                       "bbox": _r(region), "method": method, "status": "fallback", "assets": crop,
                       "warnings": warnings,
                       "fallback_reason": "hình giữ bằng ảnh vùng gốc (chưa dựng lại DrawingML từ PDF)",
                       "objects": dict(Counter(o.kind for o in cl))})
    # 5) chữ xoay ngoài hình → ảnh vùng
    rot_chars = [c for c in live if c.owner is None and not c.upright]
    if rot_chars:
        for gi, grp in enumerate(_cluster_chars(rot_chars, 8.0)):
            rect = _union((c.x0, c.top, c.x1, c.bottom) for c in grp)
            bid = f"{pid}-rot{gi + 1}"
            for c in grp:
                c.owner = bid
            crop = crop_region(pdf_path, pi, geom, _expand(rect, 2), assets, bid, dpi=dpi, want_svg=True)
            blocks.append({"block_id": bid, "type": "image", "page": pno, "bbox": _r(rect),
                           "method": "vector_image" if crop.get("svg") else "raster_image", "status": "fallback",
                           "assets": crop, "fallback_reason": "chữ xoay (không nằm ngang) → ảnh vùng",
                           "text": "".join(c.text for c in grp)})
    # 6) dòng chữ
    text_chars = _compose_accents(sorted([c for c in live if c.owner is None], key=lambda c: (c.top, c.x0)))
    raw_lines = _build_lines(text_chars)
    lines: list[Line] = []
    bases = [(_line_base(lc), lc) for lc in raw_lines]
    assigned: dict[int, list[Atom]] = defaultdict(list)
    orphans: list[Atom] = []
    for a in struct_atoms:
        if a.kind not in ("frac", "rad", "acc", "complex"):
            continue
        ay = a.anchor if a.anchor is not None else (a.bbox[1] + a.bbox[3]) / 2
        best, bestd = None, 1e9
        for li, ((ls, lby), lc) in enumerate(bases):
            lx0, lx1 = min(c.x0 for c in lc), max(c.x1 for c in lc)
            hdist = max(0.0, lx0 - a.bbox[2], a.bbox[0] - lx1)
            d = abs(lby - ay)
            if d <= 0.6 * ls and hdist < 6 * ls and d < bestd:
                best, bestd = li, d
        if best is None:
            orphans.append(a)
        else:
            assigned[best].append(a)
    for li, ((_ls, _lby), lc) in enumerate(bases):
        lines.append(_line_atoms(lc, assigned.get(li, [])))
    for a in orphans:
        lines.append(Line([a], a.bbox, 10.0, a.anchor or a.bbox[3]))
    lines.sort(key=lambda ln: (ln.bbox[1], ln.bbox[0]))
    # Số trang ở lề trên/dưới (dòng chỉ có số, "Trang n", "n/m", "- n -") là phần trình bày của bản in: không đưa vào
    # thân văn bản Word (Word tự đánh số trang qua trường PAGE ở chân trang); ghi lại thành khối page_number.
    H = geom.height
    kept: list[Line] = []
    for ln in lines:
        txt = _line_plain(ln)[0].strip()
        near_edge = ln.bbox[1] > 0.90 * H or ln.bbox[3] < 0.07 * H
        if near_edge and len(txt) <= 16 and _PAGE_NO_RE.match(txt) and all(a.kind == "char" for a in ln.atoms):
            for a in ln.atoms:
                for c in a.chars:
                    c.owner = "artifact:page_number"
            blocks.append({"block_id": f"{pid}-pageno", "type": "page_number", "page": pno, "bbox": _r(ln.bbox),
                           "method": "text_native", "status": "artifact", "text": txt,
                           "warnings": ["số trang của bản in → trường PAGE ở chân trang Word, không chép vào thân"]})
            continue
        kept.append(ln)
    lines = kept
    lines, cols = _split_at_gutter(lines, W)
    pinfo["columns"] = cols
    # 7) dòng → khối đoạn / toán hiển thị / trắc nghiệm
    blocks += _lines_to_blocks(lines, pid, pno, geom, pdf_path, pi, assets, dpi, cols, issue)
    # 8) thứ tự đọc + bố cục hình cạnh chữ (immini)
    ordered = _reading_order(blocks, cols, W)
    ordered = _side_by_side_groups(ordered, pid, pno)
    for b in ordered:
        add_block(b)
    # Chữ tham chiếu ĐỘC LẬP cho kiểm định: mọi ký tự hiển thị của trang, trừ ký tự nằm trong vùng đã giữ
    # bằng ảnh (hình/bảng-ảnh/công thức-ảnh/chữ xoay) — các ký tự đó được kiểm ở nhóm hình/hiển thị.
    image_ids = {b["block_id"] for b in blocks if b.get("method") in ("vector_image", "raster_image", "page_image")}
    image_regions = [tuple(b["bbox"]) for b in blocks if b.get("method") in ("vector_image", "raster_image")]
    for b in blocks:
        for part in (b.get("lines") or []) + (b.get("options") or []):
            for r in part.get("runs", []):
                if r.get("image") and r.get("bbox"):
                    image_regions.append(tuple(r["bbox"]))
    # glyph không có ánh xạ Unicode được giữ bằng ẢNH glyph tại chỗ → không thuộc chữ tham chiếu
    ref_chars = [c for c in chars if not c.hidden and c.owner != "duplicate-hidden" and c.text.strip()
                 and not c.unmapped and not c.merged and not (c.owner or "").startswith("artifact")
                 and c.owner not in image_ids and not any(_contains(_expand(rg, 0.5), c.cx, c.cy) for rg in image_regions)]
    pinfo["reference_text"] = _chars_to_text(ref_chars)
    pinfo["reference_excluded_chars"] = sum(1 for c in chars if c.text.strip() and not c.hidden) - len(ref_chars)
    n_unm = sum(1 for c in chars if c.unmapped and not c.hidden and c.text.strip())
    if n_unm:
        issue("warning", "unmapped_glyphs",
              f"Trang {pno}: {n_unm} glyph không có ánh xạ Unicode (ToUnicode thiếu/hỏng) → giữ bằng ẢNH glyph tại chỗ; "
              "chữ ở các vị trí đó không tìm kiếm/sửa được trong Word.", page=pno)
    plaus = _text_plausibility(_chars_to_text([c for c in chars if not c.hidden and c.text.strip() and not c.unmapped
                                               and not c.merged]))
    pinfo["text_plausibility"] = plaus
    if plaus.get("suspect"):
        issue("warning", "text_mapping_suspect",
              f"Trang {pno}: chữ trích từ PDF không giống văn bản tự nhiên (heuristic: cụm dài {plaus['long_token_frac']:.0%}, "
              f"ký hiệu {plaus['punct_ratio']:.0%}, từ không nguyên âm {plaus['vowelless_word_ratio']:.0%}) — có thể bảng "
              "ánh xạ ToUnicode của phông SAI; chữ trong Word có thể khác chữ hiển thị. Cần đối chiếu bằng mắt.", page=pno)


def _chars_to_text(chars: list[Ch]) -> str:
    """Văn bản đơn giản từ ký tự (theo dòng cơ sở, chèn khoảng trắng theo khe hở) — dùng làm tham chiếu."""
    out = []
    for ln in _build_lines([c for c in chars if c.upright]):
        s = ""
        prev = None
        for c in ln:
            combining = bool(c.text) and all(unicodedata.combining(ch) for ch in c.text)
            if prev is not None and not combining and c.x0 - prev.x1 > 0.2 * min(c.size, prev.size):
                s += " "
            s += c.text
            if not combining:
                prev = c
        out.append(s)
    rot = "".join(c.text for c in chars if not c.upright)
    if rot:
        out.append(rot)
    # dấu tổ hợp (NFD) ghép vào chữ cơ sở như khi hiển thị
    return unicodedata.normalize("NFC", "\n".join(out))


_VOWELS = set("aeiouyAEIOUY")


def _text_plausibility(text: str) -> dict[str, Any]:
    """Heuristic phát hiện ÁNH XẠ CHỮ SAI (ToUnicode hỏng nhưng vẫn "hợp lệ"): chữ trích ra không giống văn bản tự
    nhiên — gần như không có khoảng trắng (cụm rất dài), tỉ lệ ký hiệu cao, hoặc nhiều "từ" không có nguyên âm.
    Chỉ là tín hiệu (không chứng minh); ngưỡng chưa hiệu chuẩn trên dữ liệu thật."""
    nonspace = [ch for ch in text if not ch.isspace()]
    if len(nonspace) < 60:
        return {"checked": False}
    toks = text.split()
    long_chars = sum(len(t) for t in toks if len(t) > 30)
    punct = sum(1 for ch in nonspace if not ch.isalnum())
    words = [t for t in re.findall(r"[^\W\d_]{3,}", text)]

    def has_vowel(w: str) -> bool:
        base = unicodedata.normalize("NFD", w)
        return any(ch in _VOWELS for ch in base)

    vowelless = sum(1 for w in words if not has_vowel(w))
    m = {"checked": True, "long_token_frac": round(long_chars / len(nonspace), 3),
         "punct_ratio": round(punct / len(nonspace), 3),
         "vowelless_word_ratio": round(vowelless / max(1, len(words)), 3), "words": len(words)}
    m["suspect"] = (m["long_token_frac"] > 0.5 or m["punct_ratio"] > 0.35
                    or (len(words) >= 8 and m["vowelless_word_ratio"] > 0.4))
    return m


def _cluster(objs: list[Obj], gap: float) -> list[list[Obj]]:
    n = len(objs)
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    boxes = [_expand(o.bbox, gap / 2) for o in objs]
    order = sorted(range(n), key=lambda i: boxes[i][0])
    for ii, i in enumerate(order):
        for j in order[ii + 1:]:
            if boxes[j][0] > boxes[i][2]:
                break
            if _inter(boxes[i], boxes[j]):
                parent[find(i)] = find(j)
    groups: dict[int, list[Obj]] = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(objs[i])
    return list(groups.values())


def _cluster_chars(chars: list[Ch], gap: float) -> list[list[Ch]]:
    objs = [Obj("char", (c.x0, c.top, c.x1, c.bottom), meta={"c": c}) for c in chars]
    return [[o.meta["c"] for o in g] for g in _cluster(objs, gap)]


def _table_grid(cells: list[Rect], chars: list[Ch]) -> dict[str, Any] | None:
    xs = sorted({round(v, 0) for c in cells for v in (c[0], c[2])})
    ys = sorted({round(v, 0) for c in cells for v in (c[1], c[3])})

    def merge(vals: list[float]) -> list[float]:
        out: list[float] = []
        for v in vals:
            if out and v - out[-1] <= 2.0:
                continue
            out.append(v)
        return out

    xs, ys = merge(xs), merge(ys)
    if len(xs) < 2 or len(ys) < 2:
        return None

    def idx(vals: list[float], v: float) -> int:
        return min(range(len(vals)), key=lambda i: abs(vals[i] - v))

    n_cols, n_rows = len(xs) - 1, len(ys) - 1
    occupied = [[False] * n_cols for _ in range(n_rows)]
    rows: list[list[dict[str, Any]]] = [[] for _ in range(n_rows)]
    for c in sorted(cells, key=lambda r: (r[1], r[0])):
        c0, c1 = idx(xs, c[0]), idx(xs, c[2])
        r0, r1 = idx(ys, c[1]), idx(ys, c[3])
        if c1 <= c0 or r1 <= r0:
            continue
        if any(occupied[r][k] for r in range(r0, r1) for k in range(c0, c1)):
            continue
        for r in range(r0, r1):
            for k in range(c0, c1):
                occupied[r][k] = True
        rows[r0].append({"row": r0, "col": c0, "rowspan": r1 - r0, "colspan": c1 - c0, "bbox": c})
    if not all(all(r) for r in occupied):
        return None
    for r in rows:
        r.sort(key=lambda c: c["col"])
    return {"rows": rows, "n_rows": n_rows, "n_cols": n_cols, "xs": xs, "ys": ys}


def _runs_from_atoms(atoms: list[Atom], size: float) -> tuple[list[dict[str, Any]], list[str], list[str], bool]:
    """(runs, latex các công thức inline native, lý do toán phải dùng ảnh, có_toán_phải_ảnh)."""
    line = Line(atoms, _union(a.bbox for a in atoms), size, 0)
    segs = _segment_line(line)
    runs: list[dict[str, Any]] = []
    maths: list[str] = []
    img_reasons: list[str] = []
    prev_x1 = None
    for s in segs:
        if prev_x1 is not None and s["atoms"][0].x0 - prev_x1 > 0.22 * size:
            runs.append({"text": " "})
        prev_x1 = s["atoms"][-1].x1
        if s["kind"] == "math":
            latex, why = _math_seg_latex(s["atoms"], size)
            if latex is None:
                runs.append({"math_image": True, "bbox": _union(a.bbox for a in s["atoms"]), "reason": why})
                img_reasons.append(why)
            else:
                runs.append({"math": latex, "bbox": _r(_union(a.bbox for a in s["atoms"]))})
                maths.append(latex)
            continue
        cur: dict[str, Any] | None = None
        pa = None
        for a in s["atoms"]:
            c = a.chars[0]
            if c.unmapped:
                runs.append({"glyph_image": True, "bbox": (c.x0, c.top, c.x1, c.bottom), "reason": "glyph không ánh xạ"})
                img_reasons.append("glyph không ánh xạ")
                cur = None
                pa = a
                continue
            space = pa is not None and a.x0 - pa.x1 > 0.22 * size
            style = (c.bold, c.italic, round(c.size * 2) / 2)
            if cur is None or (cur["bold"], cur["italic"], cur["size"]) != style:
                cur = {"text": (" " if space else "") + c.text, "bold": c.bold, "italic": c.italic, "size": style[2]}
                runs.append(cur)
            else:
                cur["text"] += (" " if space else "") + c.text
            pa = a
    return runs, maths, img_reasons, bool(img_reasons)


def _cell_paragraphs(chars: list[Ch]) -> list[dict[str, Any]]:
    chars = _compose_accents(sorted(chars, key=lambda c: (c.top, c.x0)))
    out = []
    for lc in _build_lines(chars):
        ln = _line_atoms(lc, [])
        runs, maths, _r_, has_img = _runs_from_atoms(ln.atoms, ln.size)
        text = "".join(r.get("text", "") for r in runs)
        out.append({"runs": [r for r in runs if not r.get("math_image") and not r.get("glyph_image")],
                    "math": maths, "text": text, "has_unconverted": has_img})
    return out


_MCQ_LABEL_RE = re.compile(r"(?:(?<=\s)|^)([A-D])\.(?=\s|$)")


def _line_plain(ln: Line) -> tuple[str, list[int]]:
    """Văn bản của dòng + vị trí atom tương ứng từng ký tự (để cắt theo nhãn phương án)."""
    s = ""
    pos: list[int] = []
    prev = None
    for k, a in enumerate(ln.atoms):
        if prev is not None and a.x0 - prev.x1 > 0.22 * ln.size:
            s += " "
            pos.append(k)
        s += a.chars[0].text if a.kind == "char" else "⁣"
        pos.append(k)
        prev = a
    return s, pos


def _mcq_options_in_line(ln: Line) -> list[tuple[str, int]]:
    s, pos = _line_plain(ln)
    found = []
    for m in _MCQ_LABEL_RE.finditer(s):
        found.append((m.group(1), pos[m.start(1)]))
    return found


def _lines_to_blocks(lines: list[Line], pid: str, pno: int, geom: PageGeom, pdf_path, pi, assets, dpi, cols,
                     issue) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    if not lines:
        return blocks
    body_size = statistics.median(ln.size for ln in lines)
    gaps = [lines[i + 1].bbox[1] - lines[i].bbox[3] for i in range(len(lines) - 1)
            if 0 <= lines[i + 1].bbox[1] - lines[i].bbox[3] < 3 * body_size]
    typical_gap = statistics.median(gaps) if gaps else 0.3 * body_size
    left_margin = min(ln.bbox[0] for ln in lines)
    right_margin = max(ln.bbox[2] for ln in lines)
    k = 0
    pcount = 0
    used = set()
    # trắc nghiệm: A./B./C./D. trên một hoặc nhiều dòng liên tiếp
    while k < len(lines):
        if k in used:
            k += 1
            continue
        labels = _mcq_options_in_line(lines[k])
        if labels and labels[0][0] == "A":
            seq = []
            j = k
            expect = "ABCD"
            while j < len(lines) and sum(len(l) for _, l in seq) < 4:
                lab = _mcq_options_in_line(lines[j])
                if not lab:
                    break
                letters = "".join(l for l, _ in lab)
                if not expect[sum(len(l) for _, l in seq):].startswith(letters):
                    break
                seq.append((j, lab))
                j += 1
            n_opts = sum(len(l) for _, l in seq)
            if n_opts == 4:
                opts = []
                per_row = [len(l) for _, l in seq]
                for (li, lab) in seq:
                    ln = lines[li]
                    for q, (letter, ai) in enumerate(lab):
                        end = lab[q + 1][1] if q + 1 < len(lab) else len(ln.atoms)
                        atoms = ln.atoms[ai + 2: end] if ai + 2 <= end else []
                        lab_atoms = ln.atoms[ai: ai + 2]
                        runs, maths, imgr, has_img = _runs_from_atoms(atoms, ln.size) if atoms else ([], [], [], False)
                        lb = all(a.chars[0].bold for a in lab_atoms if a.kind == "char")
                        opts.append({"label": letter + ".", "label_bold": lb, "runs": runs, "math": maths,
                                     "text": letter + ". " + "".join(r.get("text", "") for r in runs),
                                     "bbox": _r(_union([a.bbox for a in ln.atoms[ai:end]])),
                                     "has_unconverted": has_img})
                    used.add(li)
                rect = _union(lines[li].bbox for li, _ in seq)
                pcount += 1
                bid = f"{pid}-mcq{pcount}"
                blk = {"block_id": bid, "type": "mcq", "page": pno, "bbox": _r(rect), "method": "text_native",
                       "options": opts, "layout_cols": max(per_row), "answer_key": "answer_key_unknown",
                       "text": " ".join(o["text"] for o in opts),
                       "inline_math": [m for o in opts for m in o["math"]],
                       "confidence": {"value": 0.7, "kind": "heuristic"},
                       "warnings": ["đáp án đúng KHÔNG suy từ định dạng (in đậm/gạch chân) → answer_key_unknown"]}
                _attach_inline_images(blk, opts, pdf_path, pi, geom, assets, dpi)
                blocks.append(blk)
                k = j
                continue
        k += 1
    # đoạn văn
    cur: list[Line] = []

    def flush() -> None:
        nonlocal cur, pcount
        if not cur:
            return
        pcount += 1
        blocks.append(_paragraph_block(cur, f"{pid}-b{pcount}", pno, body_size, left_margin, right_margin, geom,
                                       pdf_path, pi, assets, dpi))
        cur = []

    for i, ln in enumerate(lines):
        if i in used:
            flush()
            continue
        disp = _display_math_line(ln, left_margin, right_margin, body_size)
        if disp is not None:
            flush()
            pcount += 1
            blocks.append(_display_math_block(ln, disp, f"{pid}-b{pcount}", pno, geom, pdf_path, pi, assets, dpi))
            continue
        if cur:
            prev = cur[-1]
            gap = ln.bbox[1] - prev.bbox[3]
            same_col = True
            if cols:
                g0, g1 = cols["gutter"]
                same_col = (prev.bbox[2] <= g0 + 1) == (ln.bbox[2] <= g0 + 1) and (prev.bbox[0] >= g1 - 1) == (ln.bbox[0] >= g1 - 1)
            new_para = (gap > max(1.6 * typical_gap, 0.55 * body_size) or not same_col
                        or abs(ln.size - prev.size) > 0.15 * body_size
                        or re.match(r"^(Câu|Bài|Ví dụ|Question|Problem)\s*\d+", _line_plain(ln)[0])
                        or ln.bbox[1] < prev.bbox[1])
            if new_para:
                flush()
        cur.append(ln)
    flush()
    return blocks


_EQNUM_RE = re.compile(r"^\(\s*[0-9]+(?:\.[0-9]+)*[a-z]?\s*\)$")


def _display_math_line(ln: Line, lm: float, rm: float, body_size: float) -> dict[str, Any] | None:
    """Dòng công thức đứng riêng (căn giữa/thụt sâu, toàn toán, có thể kèm số hiệu (n) ở lề phải)."""
    segs = _segment_line(ln)
    tag = None
    if segs and segs[-1]["kind"] != "math":
        last_txt = "".join(a.chars[0].text for a in segs[-1]["atoms"] if a.kind == "char").strip()
        if _EQNUM_RE.match(last_txt) and segs[-1]["atoms"][0].x0 > (lm + rm) / 2:
            tag = last_txt.strip("() ")
            segs = segs[:-1]
    if not segs or any(sg["kind"] != "math" for sg in segs):
        return None
    atoms = [a for sg in segs for a in sg["atoms"]]
    x0, x1 = atoms[0].x0, atoms[-1].x1
    center = (x0 + x1) / 2
    centered = abs(center - (lm + rm) / 2) < 0.12 * (rm - lm)
    indented = x0 > lm + 3 * body_size
    if not (centered or indented):
        return None
    if len([a for a in atoms if a.kind == "char"]) < 2 and not any(a.kind != "char" for a in atoms):
        return None
    return {"atoms": atoms, "tag": tag}


def _display_math_block(ln: Line, disp: dict[str, Any], bid: str, pno: int, geom, pdf_path, pi, assets,
                        dpi) -> dict[str, Any]:
    latex, why = _math_seg_latex(disp["atoms"], ln.size)
    rect = ln.bbox
    if latex is None:
        crop = crop_region(pdf_path, pi, geom, _expand(rect, 2), assets, bid, dpi=dpi, want_svg=True)
        return {"block_id": bid, "type": "math", "page": pno, "bbox": _r(rect),
                "method": "vector_image" if crop.get("svg") else "raster_image", "assets": crop, "status": "fallback",
                "fallback_reason": "math_as_image: " + (why or "cấu trúc chưa nhận dạng chắc"), "text": "",
                "warnings": [], "display": True}
    if disp.get("tag"):
        latex = latex + "\\tag{" + disp["tag"] + "}"
    return {"block_id": bid, "type": "math", "page": pno, "bbox": _r(rect), "method": "math_omml", "latex": [latex],
            "inline_math": [latex], "text": "", "align": "center", "display": True, "warnings": [],
            "confidence": {"value": 0.6, "kind": "heuristic"}, "status": "ok"}


def _attach_inline_images(blk: dict[str, Any], parts: list[dict[str, Any]], pdf_path, pi, geom, assets, dpi) -> None:
    imgs = []
    n = 0
    for part in parts:
        for r in part["runs"]:
            if r.get("math_image") or r.get("glyph_image"):
                n += 1
                name = f"{blk['block_id']}-m{n}"
                crop = crop_region(pdf_path, pi, geom, _expand(r["bbox"], 1.5), assets, name, dpi=dpi,
                                   want_svg=bool(r.get("math_image")))  # glyph lẻ: PNG là đủ, tránh gọi pdftocairo
                r["image"] = crop
                r["image_id"] = name
                r["bbox"] = _r(r["bbox"])
                imgs.append({"id": name, "reason": r.get("reason"), "kind": "math" if r.get("math_image") else "glyph"})
    if imgs:
        blk["inline_math_images"] = imgs
        blk["status"] = "review"
        blk["warnings"].append(f"{len(imgs)} cụm toán/glyph giữ bằng ảnh vùng (math_as_image)")


def _paragraph_block(lines: list[Line], bid: str, pno: int, body_size: float, lm: float, rm: float, geom,
                     pdf_path, pi, assets, dpi) -> dict[str, Any]:
    parts = []
    texts = []
    maths_all = []
    hyph = 0
    for idx, ln in enumerate(lines):
        runs, maths, imgr, has_img = _runs_from_atoms(ln.atoms, ln.size)
        if idx < len(lines) - 1 and runs and "text" in runs[-1] and runs[-1]["text"].endswith("-") \
                and not runs[-1].get("math") and ln.bbox[2] >= rm - 0.04 * (rm - lm):
            nxt = lines[idx + 1]
            first = nxt.atoms[0].chars[0].text if nxt.atoms and nxt.atoms[0].kind == "char" else ""
            if first.islower():
                runs[-1]["text"] = runs[-1]["text"][:-1]
                runs.append({"join": True})
                hyph += 1
        parts.append({"runs": runs, "math": maths})
        maths_all += maths
        texts.append("".join(r.get("text", "") for r in runs))
    text = ""
    for i_t, t in enumerate(texts):
        if i_t > 0 and not (parts[i_t - 1]["runs"] and parts[i_t - 1]["runs"][-1].get("join")):
            text += " "
        text += t
    rect = _union(ln.bbox for ln in lines)
    size = statistics.median(ln.size for ln in lines)
    all_math = all(all(r.get("math") or r.get("math_image") or not r.get("text", "").strip() for r in p["runs"])
                   for p in parts)
    width = rect[2] - rect[0]
    centered = abs(((rect[0] + rect[2]) / 2) - ((lm + rm) / 2)) < 0.06 * (rm - lm) and width < 0.8 * (rm - lm)
    bold_all = all(r.get("bold") for p in parts for r in p["runs"] if r.get("text", "").strip())
    btype = "paragraph"
    if all_math and maths_all and len(lines) <= 3:
        btype = "math"
    elif size >= 1.15 * body_size and len(lines) <= 3:
        btype = "heading"
    elif bold_all and len(lines) == 1 and len(text) < 90 and text.strip():
        btype = "heading" if size >= body_size else "paragraph"
    blk = {"block_id": bid, "type": btype, "page": pno, "bbox": _r(rect), "method": "math_omml" if btype == "math" else "text_native",
           "lines": parts, "text": text, "inline_math": maths_all, "align": "center" if centered else "left",
           "size": round(size * 2) / 2, "warnings": [], "status": "ok"}
    if hyph:
        blk["warnings"].append(f"nối {hyph} từ bị ngắt bằng gạch nối cuối dòng (heuristic)")
    if btype == "math":
        blk["confidence"] = {"value": 0.6, "kind": "heuristic"}
        blk["latex"] = maths_all
    _attach_inline_images(blk, parts, pdf_path, pi, geom, assets, dpi)
    if btype == "math" and blk.get("inline_math_images"):
        # toàn khối là toán và có phần không dựng được → cả khối thành ảnh (tránh nửa OMML nửa ảnh)
        crop = crop_region(pdf_path, pi, geom, _expand(rect, 2), assets, bid, dpi=dpi, want_svg=True)
        blk.update({"method": "vector_image" if crop.get("svg") else "raster_image", "assets": crop,
                    "status": "fallback", "fallback_reason": "math_as_image: cấu trúc 2D chưa nhận dạng chắc",
                    "inline_math": [], "lines": []})
    return blk


def _reading_order(blocks: list[dict[str, Any]], cols: dict[str, Any] | None, W: float) -> list[dict[str, Any]]:
    def top(b):
        return b["bbox"][1]

    if not cols:
        return sorted(blocks, key=lambda b: (round(b["bbox"][1] / 4), b["bbox"][0]))
    g0, g1 = cols["gutter"]

    def side(b) -> str:
        x0, _, x1, _ = b["bbox"]
        if x1 <= g0 + 2:
            return "L"
        if x0 >= g1 - 2:
            return "R"
        return "S"

    out: list[dict[str, Any]] = []
    band: list[dict[str, Any]] = []
    for b in sorted(blocks, key=top):
        if side(b) == "S":
            out += sorted([x for x in band if side(x) == "L"], key=top) + sorted([x for x in band if side(x) == "R"], key=top)
            band = []
            out.append(b)
        else:
            band.append(b)
    out += sorted([x for x in band if side(x) == "L"], key=top) + sorted([x for x in band if side(x) == "R"], key=top)
    for b in out:
        b["column"] = {"L": 0, "R": 1, "S": None}[side(b)]
    return out


def _side_by_side_groups(blocks: list[dict[str, Any]], pid: str, pno: int) -> list[dict[str, Any]]:
    """Hình đặt CẠNH chữ (trái/phải) — theo quan hệ vùng (chồng dọc, tách ngang), không theo ngưỡng bề rộng trang."""
    used: set[str] = set()
    layouts: dict[str, dict[str, Any]] = {}
    figs = [b for b in blocks if b["type"] in ("figure", "image")]
    n = 0
    for f in figs:
        if f["block_id"] in used:
            continue
        fx0, fy0, fx1, fy1 = f["bbox"]
        partners = [b for b in blocks if b is not f and b["block_id"] not in used
                    and b["type"] in ("paragraph", "mcq", "math", "heading")
                    and min(b["bbox"][3], fy1) - max(b["bbox"][1], fy0) >= 0.5 * min(b["bbox"][3] - b["bbox"][1], fy1 - fy0)
                    and (b["bbox"][2] <= fx0 + 2 or b["bbox"][0] >= fx1 - 2)
                    and b.get("column") == f.get("column")]
        if not partners:
            continue
        left_side = all(b["bbox"][2] <= fx0 + 2 for b in partners)
        right_side = all(b["bbox"][0] >= fx1 - 2 for b in partners)
        if not (left_side or right_side):
            continue
        n += 1
        gid = f"{pid}-lay{n}"
        texts = sorted(partners, key=lambda b: b["bbox"][1])
        for b in texts + [f]:
            b["parent"] = gid
            used.add(b["block_id"])
        rect = _union([tuple(f["bbox"])] + [tuple(b["bbox"]) for b in texts])
        layouts[gid] = {"block_id": gid, "type": "layout", "page": pno, "bbox": _r(rect), "method": "layout_table",
                        "children_left": [b["block_id"] for b in texts] if left_side else [f["block_id"]],
                        "children_right": [f["block_id"]] if left_side else [b["block_id"] for b in texts],
                        "warnings": [], "status": "ok", "text": "",
                        "note": "bố cục hình cạnh chữ (immini) suy từ quan hệ vùng"}
    out: list[dict[str, Any]] = []
    placed: set[str] = set()
    for b in blocks:
        g = b.get("parent")
        if g in layouts and g not in placed:
            out.append(layouts[g])
            placed.add(g)
        out.append(b)
    return out
