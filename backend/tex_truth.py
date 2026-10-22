"""`backend/tex_truth.py` — Bộ render "SỰ THẬT TỪ TeX" cho nhánh **dán mã LaTeX**.

Ý tưởng: với bất kỳ khối LaTeX nào mà engine gốc *chưa* hiểu (TikZ vẽ tay, pgfplots,
môi trường lạ, gói lạ, macro lạ…), ta không đoán nữa — ta để **chính TeX dựng nó**,
rồi lấy kết quả đó nhúng vào Word.  Vì ảnh được dựng bởi đúng trình biên dịch và đúng
preamble của người dùng nên nó **giống tài liệu gốc 100%** theo đúng nghĩa đen.

Tối ưu tốc độ:
  * **Gộp 1 lần biên dịch**: mọi khối lạ trong cả tài liệu được xếp vào MỘT file .tex,
    mỗi khối một trang (gói `preview` + `tightpage`), nên chỉ chạy TeX đúng 1 lần.
  * **Cache theo nội dung** (sha1 của preamble + khối + định dạng): lần sau dán lại
    cùng một khối là lấy ngay từ cache, không biên dịch lại.
  * **Tự vá gói thiếu**: nếu TeX báo thiếu `foo.sty`, tự sinh file `foo.sty` rỗng và
    biên dịch lại (tối đa `MAX_SHIMS` lần) thay vì bỏ cuộc.

Đầu ra ưu tiên **EMF (vector)** → co giãn, in nét như LaTeX; nếu máy không có công cụ
chuyển EMF thì rơi xuống **SVG** rồi **PNG DPI cao**.  Ảnh inline còn được căn đúng
đường cơ sở (baseline) nhờ số liệu `Preview: Snippet` mà gói `preview` ghi ra log.

Module này **độc lập hoàn toàn**: không import gì từ engine gốc, nên nhánh tải file
(.pdf/.tex/.zip) không hề bị ảnh hưởng.
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

__all__ = [
    "TexTruth", "RenderedAsset", "tex_engine_available", "describe_toolchain",
]

# ───────────────────────────── cấu hình ─────────────────────────────
CACHE_VERSION = "3"             # đổi số này là bỏ toàn bộ cache cũ
MAX_SHIMS = 12                  # số lần tự vá gói thiếu trước khi bỏ cuộc
DEFAULT_TIMEOUT = 240           # giây cho một lần chạy TeX
SP_PER_PT = 65536               # 1pt = 65536sp (đơn vị preview ghi ra log)
EMU_PER_PT = 12700              # 1pt = 12700 EMU (đơn vị DrawingML)

_ENGINE_ENV = "L2W_TEX_ENGINE"
_FORMAT_ENV = "L2W_TEX_TRUTH_FORMAT"    # emf | svg | png
_DPI_ENV = "L2W_TEX_TRUTH_DPI"

_PREVIEW_LOG_RE = re.compile(r"^Preview:\s+Snippet\s+(\d+)\s+(-?\d+)\s+(-?\d+)\s+(-?\d+)", re.M)
_MISSING_FILE_RE = re.compile(r"(?:File|file)\s+[`'\"]([A-Za-z0-9@_\-.]+?)['\"`]\s+not found")
_MISSING_ANY_RE = re.compile(r"[`'\"]([A-Za-z0-9@_\-]+\.(?:sty|ldf|def|cfg|clo|cls))['\"`]\s+not found")
_MISSING_PKG_RE = re.compile(r"LaTeX Error: File [`'\"]([^'\"`]+)['\"`] not found")


def _log(msg: str) -> None:
    try:
        print(f"[tex_truth] {msg}", file=sys.stderr, flush=True)
    except Exception:
        pass


# ───────────────────────── dò tìm công cụ ngoài ─────────────────────────
_MIKTEX_CACHE: dict[str, bool] = {}


def _is_miktex(exe: str) -> bool:
    """MiKTeX nhận cờ --enable-installer; TeX Live thì không, truyền vào là lỗi."""
    if exe in _MIKTEX_CACHE:
        return _MIKTEX_CACHE[exe]
    ok = False
    try:
        out = subprocess.run([exe, "--version"], timeout=30,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT).stdout
        ok = b"MiKTeX" in out or b"miktex" in out
    except Exception:
        ok = False
    _MIKTEX_CACHE[exe] = ok
    return ok


def _which(*names: str) -> str | None:
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def tex_engine_available() -> str | None:
    """Trả về đường dẫn trình biên dịch TeX đang có (ưu tiên biến môi trường)."""
    forced = os.environ.get(_ENGINE_ENV, "").strip()
    if forced:
        return _which(forced) or None
    return _which("pdflatex", "xelatex", "lualatex", "latex")


_UNICODE_ENGINE_RE = re.compile(r"\\usepackage\s*(?:\[[^\]]*\])?\s*\{[^}]*"
                                r"(?:fontspec|polyglossia|unicode-math|xeCJK|luatexja)[^}]*\}")


def _pick_engine(preamble: str) -> str | None:
    """Gói kiểu fontspec/polyglossia bắt buộc XeLaTeX/LuaLaTeX — tự chọn giúp người dùng."""
    forced = os.environ.get(_ENGINE_ENV, "").strip()
    if forced:
        return _which(forced) or tex_engine_available()
    if preamble and _UNICODE_ENGINE_RE.search(preamble):
        return _which("xelatex", "lualatex") or tex_engine_available()
    return tex_engine_available()


def _py_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False


def describe_toolchain() -> dict[str, Any]:
    """Báo cáo mọi công cụ có thể dùng để biến TeX thành ảnh — dùng cho chẩn đoán."""
    tc: dict[str, Any] = {
        "tex_engine": tex_engine_available(),
        "latex_dvi": _which("latex"),
        "dvisvgm": _which("dvisvgm"),
        "dvipng": _which("dvipng"),
        "pdftocairo": _which("pdftocairo"),
        "pdftoppm": _which("pdftoppm"),
        "inkscape": _which("inkscape"),
        "soffice": _which("soffice", "libreoffice"),
        "magick": _which("magick", "convert"),
        "pypdfium2": _py_module("pypdfium2"),
        "pymupdf": _py_module("fitz"),
    }
    tc["co_the_ra_png"] = bool(tc["pypdfium2"] or tc["pdftocairo"] or tc["pdftoppm"]
                               or tc["pymupdf"] or tc["magick"] or (tc["latex_dvi"] and tc["dvipng"]))
    tc["co_the_ra_svg"] = bool(tc["pymupdf"] or tc["dvisvgm"] or tc["pdftocairo"] or tc["inkscape"])
    tc["co_the_ra_emf"] = bool(tc["inkscape"] or tc["soffice"])
    if tc["co_the_ra_svg"]:
        best = "svg"
    elif tc["co_the_ra_emf"]:
        best = "emf"
    elif tc["co_the_ra_png"]:
        best = "png"
    else:
        best = "none"
    tc["best_vector_format"] = best
    tc["tuyen_dung_duoc"] = [r for r in ("dvi", "pdf") if _route_available(r, tc)]
    tc["ready"] = bool(tc["tex_engine"] or tc["latex_dvi"]) and bool(tc["tuyen_dung_duoc"])
    tc["loi_khuyen"] = _toolchain_advice(tc)
    return tc


def _route_available(route: str, tc: dict[str, Any]) -> bool:
    if route == "dvi":
        return bool(tc["latex_dvi"]) and bool(tc["dvipng"] or tc["dvisvgm"])
    return bool(tc["tex_engine"]) and bool(tc["co_the_ra_png"] or tc["co_the_ra_svg"] or tc["co_the_ra_emf"])


def _toolchain_advice(tc: dict[str, Any]) -> str:
    """Câu hướng dẫn NGẮN, NÓI THẲNG phải làm gì — hiện lên giao diện khi thiếu đồ."""
    if not (tc["tex_engine"] or tc["latex_dvi"]):
        return ("Máy chưa có LaTeX. Cài MiKTeX (miktex.org) hoặc TeX Live, mở lại ứng dụng. "
                "Thiếu nó thì hình lạ vẫn chuyển như cũ.")
    if not tc["tuyen_dung_duoc"]:
        return ("Có LaTeX nhưng chưa có công cụ nào biến PDF/DVI thành ảnh. "
                "Cách nhanh nhất — chạy một lệnh duy nhất:  pip install pypdfium2   "
                "(hoặc cài gói 'dvipng' trong MiKTeX Console).")
    return ""


# ───────────────────────────── mô hình dữ liệu ─────────────────────────────
@dataclass
class RenderedAsset:
    """Một khối LaTeX đã được TeX dựng xong."""
    key: str
    data: bytes = b""
    ext: str = ""                   # emf | svg | png
    png_data: bytes = b""           # bản PNG kèm theo (bắt buộc khi ext == 'svg')
    width_pt: float = 0.0
    height_pt: float = 0.0
    depth_pt: float = 0.0           # phần chìm dưới đường cơ sở (ảnh inline)
    inline: bool = False
    ok: bool = False
    error: str = ""
    tex_errors: list[str] = field(default_factory=list)

    @property
    def width_emu(self) -> int:
        return max(1, int(round(self.width_pt * EMU_PER_PT)))

    @property
    def height_emu(self) -> int:
        return max(1, int(round(self.height_pt * EMU_PER_PT)))


_PURE_PICTURE_RE = re.compile(
    r"^\s*\\begin\{(tikzpicture|pgfpicture|tikzcd|circuitikz|chemfig|forest|pspicture)\}"
    r".*\\end\{\1\}\s*$", re.S)


@dataclass
class _Snippet:
    key: str
    body: str
    inline: bool
    label: str = ""

    @property
    def needs_box(self) -> bool:
        """Hình thuần TikZ không cần khung bọc — bọc thêm chỉ rước rủi ro."""
        return not (self.inline or _PURE_PICTURE_RE.match(self.body))


# ───────────────────────────── bộ render chính ─────────────────────────────
class TexTruth:
    """Gom khối → biên dịch 1 lần → cắt trang → xuất EMF/SVG/PNG → trả về bytes."""

    def __init__(
        self,
        preamble: str = "",
        *,
        workdir: Path | None = None,
        cache_dir: Path | None = None,
        engine: str | None = None,
        fmt: str | None = None,
        dpi: int | None = None,
        timeout: int = DEFAULT_TIMEOUT,
        stubs: str = "",
    ) -> None:
        self.preamble = sanitize_preamble(preamble) + (("\n" + stubs) if stubs else "")
        self._tmp: tempfile.TemporaryDirectory | None = None
        if workdir is None:
            self._tmp = tempfile.TemporaryDirectory(prefix="l2w_textruth_")
            workdir = Path(self._tmp.name)
        self.workdir = Path(workdir)
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.cache_dir = Path(cache_dir) if cache_dir else _default_cache_dir()
        try:
            self.cache_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            self.cache_dir = self.workdir / "cache"
            self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.engine = engine or _pick_engine(self.preamble)
        tools = describe_toolchain()
        self.fmt = (fmt or os.environ.get(_FORMAT_ENV) or tools["best_vector_format"]).lower()
        if self.fmt not in ("emf", "svg", "png"):
            self.fmt = "png"
        try:
            self.dpi = int(dpi or os.environ.get(_DPI_ENV) or 600)
        except ValueError:
            self.dpi = 600
        self.timeout = timeout
        self.last_errors: list[str] = []
        self._reactive: list[str] = []      # khung rỗng CHỈ tạo khi TeX thật sự báo thiếu
        self._snippets: list[_Snippet] = []
        self._by_key: dict[str, _Snippet] = {}
        self.stats: dict[str, Any] = {
            "requested": 0, "from_cache": 0, "compiled": 0, "failed": 0,
            "tex_runs": 0, "shims_created": [], "format": self.fmt,
            "engine": self.engine, "toolchain": tools,
        }

    # ---------- nạp khối ----------
    def add(self, body: str, *, inline: bool = False, label: str = "") -> str:
        """Đăng ký một khối LaTeX cần dựng. Trả về `key` để tra cứu kết quả sau."""
        norm = body.strip()
        raw = f"{CACHE_VERSION}|{self.fmt}|{self.dpi}|{int(inline)}|{self.engine}|{self.preamble}|{norm}"
        key = hashlib.sha1(raw.encode("utf-8", "replace")).hexdigest()[:20]
        if key not in self._by_key:
            sn = _Snippet(key=key, body=norm, inline=inline, label=label or norm[:40])
            self._by_key[key] = sn
            self._snippets.append(sn)
        self.stats["requested"] += 1
        return key

    def __len__(self) -> int:
        return len(self._snippets)

    # ---------- dựng ----------
    def build(self) -> dict[str, RenderedAsset]:
        """Biên dịch tất cả khối đã đăng ký. Không bao giờ ném lỗi ra ngoài.

        Thử lần lượt nhiều TUYẾN cho tới khi ra được ảnh:

        * tuyến ``pdf`` — pdflatex/xelatex → PDF → SVG/EMF/PNG (cần 1 công cụ chuyển);
        * tuyến ``dvi`` — latex → DVI → ``dvisvgm`` / ``dvipng`` (**chỉ cần bản TeX**,
          không cần Poppler, Ghostscript, Inkscape hay LibreOffice).

        Khối nào tuyến đầu không ra thì tự động thử tuyến sau, nên máy chỉ cài MiKTeX
        trần vẫn chạy được."""
        out: dict[str, RenderedAsset] = {}
        if not self._snippets:
            return out
        pending: list[_Snippet] = []
        for sn in self._snippets:
            cached = self._cache_load(sn)
            if cached is not None:
                out[sn.key] = cached
                self.stats["from_cache"] += 1
            else:
                pending.append(sn)
        if not pending:
            return out

        tc = describe_toolchain()
        self.stats["toolchain"] = tc
        routes = self._route_order(tc)
        self.stats["tuyen_da_thu"] = []
        if not routes:
            for sn in pending:
                out[sn.key] = RenderedAsset(key=sn.key, inline=sn.inline, ok=False,
                                            error=tc.get("loi_khuyen") or "Máy thiếu công cụ dựng ảnh từ TeX.")
                self.stats["failed"] += 1
            self.stats["loi_khuyen"] = tc.get("loi_khuyen", "")
            return out

        base_preamble = self.preamble
        self.stats["preamble_da_thu"] = []
        for variant in range(3):
            if not pending:
                break
            if variant == 1:                 # preamble tối giản, giữ cấu hình vẽ của người dùng
                self.preamble = minimal_preamble("\n".join(s.body for s in pending), base_preamble)
            elif variant == 2:               # tối giản tuyệt đối, bỏ cả macro người dùng
                self.preamble = minimal_preamble("\n".join(s.body for s in pending), base_preamble,
                                                 keep_user_lines=False)
            if variant:
                self._reactive = []
                self.stats["preamble_da_thu"].append(f"toi_gian_{variant}")
                _log(f"còn {len(pending)} khối chưa dựng được → thử lại với preamble tối giản ({variant})")
            for route in routes:
                if not pending:
                    break
                tag = f"{route}" if not variant else f"{route}#{variant}"
                if tag not in self.stats["tuyen_da_thu"]:
                    self.stats["tuyen_da_thu"].append(tag)
                try:
                    got = self._render_route(route, pending, variant=variant)
                except Exception as exc:
                    _log(f"tuyến {route} lỗi: {type(exc).__name__}: {exc}")
                    got = {}
                for sn in list(pending):
                    a = got.get(sn.key)
                    if a is not None and a.ok:
                        if variant:
                            a.error = ""
                        out[sn.key] = a
                        pending.remove(sn)
                    elif a is not None:
                        out[sn.key] = a      # giữ lại thông báo lỗi để báo cho người dùng
        self.preamble = base_preamble
        for sn in pending:
            out.setdefault(sn.key, RenderedAsset(key=sn.key, inline=sn.inline, ok=False,
                                                 error="Mọi tuyến dựng ảnh đều thất bại."))
            self.stats["failed"] += 1
        if pending:
            self.stats["loi_khuyen"] = tc.get("loi_khuyen", "") or (
                "TeX chạy được nhưng không tạo được ảnh. Thử:  pip install pypdfium2")
        for asset in out.values():
            if asset.ok:
                self._cache_store(asset)
        return out

    def _route_order(self, tc: dict[str, Any]) -> list[str]:
        forced = os.environ.get("L2W_TEX_ROUTE", "").strip().lower()
        if forced in ("pdf", "dvi"):
            return [forced] if _route_available(forced, tc) else []
        order = []
        if _route_available("pdf", tc):
            order.append("pdf")
        if _route_available("dvi", tc):
            order.append("dvi")
        return order

    def _render_route(self, route: str, snippets: list[_Snippet], variant: int = 0) -> dict[str, RenderedAsset]:
        if route == "dvi":
            return self._render_dvi(snippets, variant=variant)
        return self._render_pdf(snippets, variant=variant)

    # ---------- tuyến PDF ----------
    def _render_pdf(self, snippets: list[_Snippet], variant: int = 0) -> dict[str, RenderedAsset]:
        out: dict[str, RenderedAsset] = {}
        batch = self._compile_batch(snippets, tag=f"batch{variant or ''}")
        if batch is not None:
            pdf, boxes = batch
            out.update(self._slice_pdf(pdf, snippets, boxes))
        missing = [sn for sn in snippets if not out.get(sn.key, RenderedAsset(key=sn.key)).ok]
        for sn in missing:                       # biên dịch riêng từng khối hỏng
            single = self._compile_batch([sn], tag=f"solo{variant or ''}_{sn.key}")
            errs = list(self.last_errors)
            if single is not None:
                pdf, boxes = single
                got = self._slice_pdf(pdf, [sn], boxes)
                if got.get(sn.key) is not None:
                    out[sn.key] = got[sn.key]
            a = out.setdefault(sn.key, RenderedAsset(key=sn.key, inline=sn.inline, ok=False,
                                                     error="Biên dịch TeX (tuyến PDF) thất bại."))
            if not a.ok:
                a.tex_errors = errs[:4]
                if errs:
                    a.error = (a.error + " TeX báo: " + " | ".join(errs[:2]))[:400]
        self.stats["compiled"] += sum(1 for a in out.values() if a.ok)
        return out

    # ---------- tuyến DVI (chỉ cần bản TeX) ----------
    def _render_dvi(self, snippets: list[_Snippet], variant: int = 0) -> dict[str, RenderedAsset]:
        """latex → .dvi → dvisvgm (vector) và/hoặc dvipng (ảnh điểm).

        Hai công cụ này đi kèm sẵn trong MiKTeX và TeX Live nên tuyến này chạy được
        trên máy hoàn toàn không có Poppler / Ghostscript / Inkscape / LibreOffice."""
        out: dict[str, RenderedAsset] = {}
        latex = _which("latex")
        if not latex:
            return out
        d = self.workdir / f"dvi{variant or ''}"
        d.mkdir(parents=True, exist_ok=True)
        has_svg, has_png = bool(_which("dvisvgm")), bool(_which("dvipng"))

        svgs: dict[int, bytes] = {}
        sizes: dict[int, tuple[float, float]] = {}
        if has_svg:
            dvi = self._compile_dvi(latex, d / "svg", snippets, driver="dvisvgm")
            if dvi is not None:
                svgs, sizes = _dvi_to_svgs(dvi, len(snippets))
        pngs: dict[int, bytes] = {}
        png_sizes: dict[int, tuple[float, float]] = {}
        if has_png and (not svgs or True):       # luôn lấy PNG để Word đời cũ vẫn thấy hình
            dvi = self._compile_dvi(latex, d / "png", snippets, driver="")
            if dvi is not None:
                pngs, png_sizes = _dvi_to_pngs(dvi, len(snippets), self.dpi)

        for i, sn in enumerate(snippets, start=1):
            svg, png = svgs.get(i, b""), pngs.get(i, b"")
            w, h = sizes.get(i) or png_sizes.get(i) or (0.0, 0.0)
            asset = RenderedAsset(key=sn.key, inline=sn.inline, width_pt=w, height_pt=h)
            if svg and png:
                asset.data, asset.ext, asset.png_data = svg, "svg", png
            elif png:
                asset.data, asset.ext = png, "png"
            elif svg:
                asset.data, asset.ext, asset.png_data = svg, "svg", b""
            else:
                asset.error = "Tuyến DVI không tạo được ảnh cho khối này."
                out[sn.key] = asset
                continue
            if _is_degenerate(w, h, sn.body, sn.inline):
                asset.ok = False
                asset.error = f"Tuyến DVI dựng ra khối rỗng/hỏng ({w:.0f}x{h:.0f}pt)."
            else:
                asset.ok = True
            out[sn.key] = asset
        self.stats["compiled"] += sum(1 for a in out.values() if a.ok)
        return out

    def _compile_dvi(self, latex: str, d: Path, snippets: list[_Snippet], driver: str) -> Path | None:
        d.mkdir(parents=True, exist_ok=True)
        tex_path = d / "l2w_truth.tex"
        tex_path.write_text(self._build_source(snippets, driver=driver), encoding="utf-8")
        saved, self.engine = self.engine, latex
        try:
            shims: set[str] = set()
            for _ in range(MAX_SHIMS + 1):
                log = self._run_engine(tex_path, d)
                dvi = tex_path.with_suffix(".dvi")
                if dvi.exists() and dvi.stat().st_size > 0:
                    if self._add_reactive_stubs(log):
                        tex_path.write_text(self._build_source(snippets, driver=driver), encoding="utf-8")
                        continue
                    return dvi
                if self._add_reactive_stubs(log):
                    tex_path.write_text(self._build_source(snippets, driver=driver), encoding="utf-8")
                    continue
                missing = _first_missing_package(log)
                if not missing or missing in shims:
                    _log(f"tuyến DVI không biên dịch được; cuối log: {log[-200:].strip()}")
                    return None
                shims.add(missing)
                _write_shim(d, missing, engine=latex)
        finally:
            self.engine = saved
        return None

    # ---------- biên dịch ----------
    def _compile_batch(self, snippets: list[_Snippet], tag: str = "batch") -> tuple[Path, dict[int, tuple[float, float, float]]] | None:
        d = self.workdir / tag
        d.mkdir(parents=True, exist_ok=True)
        tex_path = d / "l2w_truth.tex"
        tex_path.write_text(self._build_source(snippets), encoding="utf-8")
        shims: set[str] = set()
        repairs: set[str] = set()
        for _ in range(MAX_SHIMS + 1):
            log = self._run_engine(tex_path, d)
            pdf = tex_path.with_suffix(".pdf")
            boxes = _parse_preview_boxes(log)
            self.last_errors = extract_tex_errors(log)
            if self.last_errors:
                self.stats.setdefault("tex_errors", [])
                for e in self.last_errors:
                    if e not in self.stats["tex_errors"]:
                        self.stats["tex_errors"].append(e)
            if pdf.exists() and pdf.stat().st_size > 0 and _pdf_page_count(pdf) >= len(snippets):
                # Dựng được rồi, nhưng nếu log vẫn báo tên chưa định nghĩa thì ảnh đang lẫn chữ
                # thừa — vá khung rỗng đúng những tên đó rồi dựng lại một lần cho sạch.
                if self._add_reactive_stubs(log):
                    tex_path.write_text(self._build_source(snippets), encoding="utf-8")
                    continue
                self.stats["compiled"] += len(snippets)
                return pdf, boxes
            if self._add_reactive_stubs(log):
                tex_path.write_text(self._build_source(snippets), encoding="utf-8")
                continue
            if self._repair_preamble(log, repairs):
                tex_path.write_text(self._build_source(snippets), encoding="utf-8")
                continue
            missing = _first_missing_package(log)
            if not missing or missing in shims:
                _log(f"biên dịch thất bại ({tag}); 200 ký tự cuối log: {log[-200:].strip()}")
                return None
            shims.add(missing)
            # Gói tiếng Việt / Unicode thiếu hẳn trên máy → chuyển sang XeLaTeX + fontspec,
            # nhờ đó chữ trong ảnh vẫn đủ dấu thay vì rơi mất dấu.
            if missing.lower() in UNICODE_FALLBACK_PACKAGES:
                xe = _which("xelatex", "lualatex")
                if xe and xe != self.engine:
                    self.engine = xe
                    self.stats["engine"] = xe
                    _log(f"gói '{missing}' không có → chuyển sang {Path(xe).name} để giữ dấu tiếng Việt")
            _write_shim(d, missing, engine=self.engine)
            self.stats["shims_created"].append(missing)
            _log(f"thiếu gói '{missing}' → tự tạo bản vá rỗng rồi biên dịch lại")
        return None

    def _add_reactive_stubs(self, log: str) -> bool:
        """Tạo khung rỗng cho đúng những tên TeX vừa báo là chưa định nghĩa."""
        cmds, envs = undefined_names(log)
        added = False
        have = "".join(self._reactive)
        for c in cmds:
            if c in _NEVER_STUB or f"\\providecommand{{\\{c}}}" in have:
                continue
            self._reactive.append("\\providecommand{\\%s}{}\n" % c)
            added = True
        for e in envs:
            if f"newenvironment{{{e}}}" in have:
                continue
            self._reactive.append(
                "\\makeatletter\\@ifundefined{%s}{\\newenvironment{%s}[1][]{}{}}{}\\makeatother\n" % (e, e))
            added = True
        if added:
            self.stats.setdefault("khung_rong_tu_tao", [])
            self.stats["khung_rong_tu_tao"] = sorted(set(self.stats["khung_rong_tu_tao"]) | set(cmds) | set(envs))
            _log("TeX báo thiếu " + ", ".join(cmds + envs) + " → tạo khung rỗng rồi dựng lại")
        return added

    def _repair_preamble(self, log: str, done: set[str]) -> bool:
        """Tự gỡ gói gây lỗi ra khỏi preamble rồi dựng lại.

        Hay gặp nhất: máy chưa cài gói ngôn ngữ của babel (`vietnamese.ldf`). Bỏ `babel`
        đi thì chữ vẫn đúng, chỉ mất quy tắc ngắt từ — tốt hơn nhiều so với mất cả hình."""
        got = _first_missing_file(log)
        if not got:
            return False
        name, ext = got
        targets: list[str] = []
        if ext in ("ldf", "cfg", "def"):
            targets = ["babel", "polyglossia"]
        elif ext == "sty" and name.lower() in NEVER_SHIM and name.lower() not in _CORE_FOR_MECHANISM:
            targets = [name]
        for t in targets:
            tag = f"{t}:{name}.{ext}"
            if tag in done:
                continue
            new, n = re.subn(r"\\usepackage\s*(?:\[[^\]]*\])?\s*\{[^}]*\b" + re.escape(t) + r"\b[^}]*\}[^\n]*\n?",
                             "", self.preamble)
            if n:
                done.add(tag)
                self.preamble = new
                self.stats.setdefault("goi_da_go", []).append(t)
                _log(f"thiếu {name}.{ext} → tự gỡ gói '{t}' khỏi preamble rồi dựng lại")
                return True
        return False

    def _run_engine(self, tex_path: Path, cwd: Path) -> str:
        self.stats["tex_runs"] += 1
        env = dict(os.environ)
        env.setdefault("TEXINPUTS", f"{cwd}{os.pathsep}{os.pathsep}")
        env["max_print_line"] = "1000"
        exe = self.engine or "pdflatex"
        cmd = [exe, "-interaction=nonstopmode", "-file-line-error"]
        if _is_miktex(exe):
            cmd.append("--enable-installer")   # MiKTeX: tự tải gói thiếu, không hiện hộp thoại
        cmd.append(tex_path.name)
        if (self.engine or "").lower().endswith("latex") and "pdflatex" not in (self.engine or "").lower():
            pass  # xelatex/lualatex đều xuất PDF trực tiếp
        try:
            proc = subprocess.run(cmd, cwd=str(cwd), env=env, timeout=self.timeout,
                                  stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            out = proc.stdout.decode("utf-8", "replace")
        except subprocess.TimeoutExpired:
            out = "TIMEOUT"
        except Exception as exc:  # engine không chạy được
            out = f"ENGINE ERROR: {exc}"
        logf = tex_path.with_suffix(".log")
        if logf.exists():
            try:
                out += "\n" + logf.read_text(encoding="utf-8", errors="replace")
            except Exception:
                pass
        return out

    def _build_source(self, snippets: list[_Snippet], driver: str = "") -> str:
        parts = [
            "\\documentclass[12pt]{article}\n",
            ("\\def\\pgfsysdriver{pgfsys-dvisvgm.def}\n" if driver == "dvisvgm" else ""),
            "\\usepackage[active,tightpage]{preview}\n",
            "\\setlength\\PreviewBorder{1pt}\n",
            "\\IfFileExists{varwidth.sty}"
            "{\\usepackage{varwidth}\\newenvironment{l2wbox}[1]{\\begin{varwidth}{#1}}{\\end{varwidth}}}"
            "{\\newenvironment{l2wbox}[1]{\\begin{minipage}{#1}}{\\end{minipage}}}\n",
            self.preamble,
            ("".join(self._reactive) if self._reactive else ""),
            "\n\\pagestyle{empty}\n",
            "\\begin{document}\n",
        ]
        for sn in snippets:
            parts.append("\\begin{preview}%\n")
            if not sn.needs_box:
                parts.append(sn.body + "%\n")
            else:
                parts.append("\\begin{l2wbox}{\\linewidth}\n" + sn.body + "\n\\end{l2wbox}%\n")
            parts.append("\\end{preview}%\n")
        parts.append("\\end{document}\n")
        return "".join(parts)

    # ---------- cắt trang & chuyển định dạng ----------
    def _slice_pdf(self, pdf: Path, snippets: list[_Snippet],
                   boxes: dict[int, tuple[float, float, float]]) -> dict[str, RenderedAsset]:
        res: dict[str, RenderedAsset] = {}
        d = pdf.parent
        for i, sn in enumerate(snippets, start=1):
            asset = RenderedAsset(key=sn.key, inline=sn.inline)
            page = d / f"page_{i}.pdf"
            if not _extract_page(pdf, i, page):
                asset.error = f"Không tách được trang {i}."
                res[sn.key] = asset
                continue
            w, h = _pdf_page_size_pt(page)
            asset.width_pt, asset.height_pt = w, h
            if _is_degenerate(w, h, sn.body, sn.inline):
                asset.error = (f"TeX dựng ra khối rỗng/hỏng ({w:.0f}x{h:.0f}pt) — "
                               "giữ nguyên cách xử lý cũ của engine.")
                res[sn.key] = asset
                continue
            hgt, dep, _wd = boxes.get(i, (0.0, 0.0, 0.0))
            if dep:
                asset.depth_pt = dep / SP_PER_PT
            data, ext, png = self._convert(page)
            if not data:
                asset.error = "Không chuyển được sang EMF/SVG/PNG."
                res[sn.key] = asset
                continue
            asset.data, asset.ext, asset.png_data = data, ext, png
            asset.ok = True
            res[sn.key] = asset
        return res

    def _convert(self, page_pdf: Path) -> tuple[bytes, str, bytes]:
        """PDF 1 trang → (bytes, phần mở rộng, bytes PNG kèm theo)."""
        order = {"emf": ["emf", "svg", "png"], "svg": ["svg", "emf", "png"], "png": ["png", "emf", "svg"]}[self.fmt]
        png_cache = b""
        for want in order:
            try:
                if want == "emf":
                    data = _pdf_to_emf(page_pdf)
                elif want == "svg":
                    data = _pdf_to_svg(page_pdf)
                else:
                    data = png_cache or _pdf_to_png(page_pdf, self.dpi)
            except Exception:
                data = b""
            if want == "png" and data:
                png_cache = data
            if data:
                if want == "svg":
                    png_cache = png_cache or _pdf_to_png(page_pdf, min(self.dpi, 300))
                return data, want, png_cache
        return b"", "", b""

    # ---------- cache ----------
    def _cache_paths(self, key: str) -> tuple[Path, Path, Path]:
        base = self.cache_dir / key[:2]
        return base / f"{key}.bin", base / f"{key}.meta", base / f"{key}.png"

    def _cache_load(self, sn: _Snippet) -> RenderedAsset | None:
        binp, metap, pngp = self._cache_paths(sn.key)
        if not (binp.exists() and metap.exists()):
            return None
        try:
            meta = dict(line.split("=", 1) for line in metap.read_text(encoding="utf-8").splitlines() if "=" in line)
            return RenderedAsset(
                key=sn.key, data=binp.read_bytes(), ext=meta.get("ext", "png"),
                png_data=pngp.read_bytes() if pngp.exists() else b"",
                width_pt=float(meta.get("w", 0)), height_pt=float(meta.get("h", 0)),
                depth_pt=float(meta.get("d", 0)), inline=sn.inline, ok=True,
            )
        except Exception:
            return None

    def _cache_store(self, asset: RenderedAsset) -> None:
        try:
            binp, metap, pngp = self._cache_paths(asset.key)
            binp.parent.mkdir(parents=True, exist_ok=True)
            binp.write_bytes(asset.data)
            metap.write_text(f"ext={asset.ext}\nw={asset.width_pt}\nh={asset.height_pt}\nd={asset.depth_pt}\n",
                             encoding="utf-8")
            if asset.png_data:
                pngp.write_bytes(asset.png_data)
        except Exception:
            pass

    def close(self) -> None:
        if self._tmp is not None:
            try:
                self._tmp.cleanup()
            except Exception:
                pass
            self._tmp = None


# ───────────────────────── hàm phụ trợ cấp module ─────────────────────────
def _default_cache_dir() -> Path:
    env = os.environ.get("L2W_TEX_TRUTH_CACHE")
    if env:
        return Path(env)
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
    return Path(base) / "latex2word" / "tex_truth"


_PREAMBLE_DROP_RE = re.compile(
    r"\\(?:documentclass|begin\{document\}|end\{document\}|pagestyle|thispagestyle|"
    r"fancyhf|fancyhead|fancyfoot|headrulewidth|footrulewidth|maketitle|tableofcontents)\b[^\n]*\n?")
_GEOMETRY_RE = re.compile(r"\\usepackage\s*(\[[^\]]*\])?\s*\{geometry\}[^\n]*\n?")
_HDR_RENEW_RE = re.compile(r"\\renewcommand\s*\{?\\(?:headrulewidth|footrulewidth)\}?[^\n]*\n?")


def extract_preamble(src: str) -> str:
    """Lấy preamble người dùng (giữa `\\documentclass` và `\\begin{document}`)."""
    m = re.search(r"\\begin\s*\{document\}", src)
    head = src[: m.start()] if m else src
    m2 = re.search(r"\\documentclass[^\n]*\n", head)
    return head[m2.end():] if m2 else head


def sanitize_preamble(preamble: str) -> str:
    """Bỏ những lệnh chỉ có nghĩa với tài liệu đầy đủ (lề trang, header/footer…)."""
    if not preamble:
        return ""
    p = _GEOMETRY_RE.sub("", preamble)
    p = _HDR_RENEW_RE.sub("", p)
    p = _PREAMBLE_DROP_RE.sub("", p)
    p = re.sub(r"\\usepackage\s*(\[[^\]]*\])?\s*\{(?:fancyhdr|titlesec|titletoc|hyperref)\}[^\n]*\n?", "", p)
    return p.strip() + "\n"


MIN_BLOCK_W_PT = 10.0
MIN_BLOCK_H_PT = 8.0
MIN_AREA_FOR_LONG_BODY = 400.0


def _is_degenerate(w: float, h: float, body: str, inline: bool) -> bool:
    """Trang quá nhỏ so với lượng mã nguồn ⇒ gần như chắc chắn TeX đã dựng hỏng."""
    if w <= 0 or h <= 0:
        return True
    if inline:
        return w < 2.0 or h < 2.0
    if w < MIN_BLOCK_W_PT or h < MIN_BLOCK_H_PT:
        return True
    if len(body) > 200 and w * h < MIN_AREA_FOR_LONG_BODY:
        return True
    return False


_TEX_ERROR_RE = re.compile(r"^(?:! .*|.*?:\d+: (?:LaTeX|Package|Class)? ?Error.*|.*?:\d+: Undefined control sequence.*)$",
                           re.M)


_UNDEF_CS_RE = re.compile(r"Undefined control sequence\.?\s*\n(?:[^\n]*\n){0,2}?[^\n]*?\\([A-Za-z@]+)\s*$",
                          re.M)
_UNDEF_CS_SIMPLE_RE = re.compile(r"^l\.\d+[^\n]*?\\([A-Za-z@]+)\s*$", re.M)
_UNDEF_ENV_RE = re.compile(r"Environment\s+([A-Za-z@*]+)\s+undefined")


def undefined_names(log: str) -> tuple[list[str], list[str]]:
    """Trả về (lệnh chưa định nghĩa, môi trường chưa định nghĩa) đúng như TeX báo.

    Đây là cách AN TOÀN để vá: chỉ tạo khung rỗng cho thứ TeX thật sự kêu thiếu, thay vì
    đoán trước rồi vô tình đè lên lệnh của TikZ / pgfplots (đó là lỗi đã gây hỏng 3 hình)."""
    log = log or ""
    cmds: list[str] = []
    for m in re.finditer(r"Undefined control sequence", log):
        seg = log[m.end():m.end() + 400]
        hit = _UNDEF_CS_SIMPLE_RE.search(seg)
        if hit and hit.group(1) not in cmds:
            cmds.append(hit.group(1))
    envs = []
    for m in _UNDEF_ENV_RE.finditer(log):
        if m.group(1) not in envs:
            envs.append(m.group(1))
    return cmds, envs


def extract_tex_errors(log: str, limit: int = 8) -> list[str]:
    seen: list[str] = []
    for m in _TEX_ERROR_RE.finditer(log or ""):
        line = m.group(0).strip()
        if line and line not in seen:
            seen.append(line)
        if len(seen) >= limit:
            break
    return seen


_CARRY_NAMES = ("usetikzlibrary", "usepgfplotslibrary", "pgfplotsset", "tikzset", "tikzstyle",
                "definecolor", "colorlet", "newcommand", "renewcommand", "providecommand",
                "DeclareMathOperator", "DeclarePairedDelimiter", "def", "newif",
                "newlength", "setlength", "newcounter", "pgfkeys", "pgfdeclareimage",
                "pgfdeclareplotmark", "SetKwFunction", "newtcolorbox", "tcbset")
_CARRY_START_RE = re.compile(r"\\(" + "|".join(_CARRY_NAMES) + r")(?![A-Za-z])")
_FONT_PKG_RE = re.compile(
    r"^[ \t]*\\usepackage\s*(?:\[[^\]]*\])?\s*\{[^}]*"
    r"(?:vietnam|vntex|inputenc|fontenc|babel|polyglossia|fontspec|unicode-math)[^}]*\}[^\n]*$", re.M)


def _balanced(src: str, i: int) -> int:
    """`i` trỏ vào `{` → vị trí ngay sau `}` khớp; -1 nếu không khớp."""
    if i >= len(src) or src[i] != "{":
        return -1
    depth, j = 0, i
    while j < len(src):
        c = src[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return j + 1
        j += 1
    return -1


def _carry_lines(preamble: str) -> list[str]:
    r"""Lấy TRỌN VẸN từng lệnh cấu hình/định nghĩa của người dùng, kể cả thân nhiều dòng.

    Quy tắc dừng: sau tên lệnh cho phép **một** tên lệnh nữa (để bắt `\def\foo`,
    `\newif\ifcorrect`, `\newcommand\foo`), rồi chỉ nuốt các nhóm `{…}` / `[…]`.
    Gặp tên lệnh khác là dừng — nếu không sẽ nuốt luôn lệnh kế tiếp và sinh ra
    định nghĩa trùng (`\True already defined`) hoặc kéo theo cả `\usepackage` lạ."""
    out: list[str] = []
    n = len(preamble)
    for m in _CARRY_START_RE.finditer(preamble):
        i = m.end()
        seen_cs = False
        seen_group = False
        while i < n:
            c = preamble[i]
            if c in " \t\r\n":
                i += 1
                continue
            if c == "*" and not seen_group:
                i += 1
                continue
            if c == "\\" and not seen_cs and not seen_group:
                mm = re.match(r"\\[A-Za-z@]+", preamble[i:])
                if not mm:
                    break
                i += mm.end()
                seen_cs = True
                continue
            if c == "{":
                j = _balanced(preamble, i)
                if j < 0:
                    break
                i = j
                seen_group = True
                continue
            if c == "[":
                j = preamble.find("]", i)
                if j < 0:
                    break
                i = j + 1
                continue
            break
        piece = preamble[m.start():i].strip()
        if piece and piece.count("{") == piece.count("}") and piece not in out:
            out.append(piece)
    return out


#: dấu hiệu trong khối → gói cần nạp
_PKG_HINTS: list[tuple[str, tuple[str, ...]]] = [
    (r"\\begin\{axis\}|\\addplot|semilogxaxis|semilogyaxis|loglogaxis|groupplot|polaraxis", ("pgfplots",)),
    (r"\\tkzTab|\\tkzInit|\\tkzAxe|\\tkzFct", ("tkz-tab",)),
    (r"\\begin\{circuitikz\}|\\ctikzset", ("circuitikz",)),
    (r"\\chemfig|\\schemestart", ("chemfig",)),
    (r"\\ce\{|\\cee\{", ("mhchem",)),
    (r"\\SI\{|\\si\{|\\num\{|\\qty\{", ("siunitx",)),
    (r"\\begin\{tcolorbox\}|\\tcbset|\\newtcolorbox", ("tcolorbox",)),
    (r"\\begin\{forest\}", ("forest",)),
    (r"\\begin\{tikzcd\}", ("tikz-cd",)),
    (r"\\begin\{bytefield\}", ("bytefield",)),
    (r"\\begin\{ganttchart\}", ("pgfgantt",)),
    (r"\\multirow", ("multirow",)),
    (r"\\toprule|\\midrule|\\bottomrule", ("booktabs",)),
    (r"\\begin\{tabularx\}", ("tabularx",)),
    (r"\\begin\{tabular\*?\}|\\begin\{array\}", ("array",)),
]


def minimal_preamble(block: str, user_preamble: str, *, keep_user_lines: bool = True) -> str:
    """Preamble TỐI GIẢN: chỉ những gói khối này thật sự cần, cộng các dòng cấu hình vẽ.

    Dùng khi preamble đầy đủ của người dùng làm TeX gãy (gói ngôn ngữ, font, lớp tài liệu
    lạ…). Bỏ hết thứ không liên quan đến việc vẽ thì khối thường dựng được ngay."""
    pkgs: list[str] = ["amsmath", "amssymb", "amsfonts"]
    if re.search(r"\\begin\{tikzpicture\}|\\tikz|\\begin\{scope\}", block) or True:
        pkgs.append("tikz")
    for pat, names in _PKG_HINTS:
        if re.search(pat, block):
            pkgs.extend(n for n in names if n not in pkgs)
    lines = ["\\usepackage{%s}" % p for p in dict.fromkeys(pkgs)]
    if any(ord(c) > 127 for c in block):
        lines.extend(m.group(0).strip() for m in _FONT_PKG_RE.finditer(user_preamble))
    carried: list[str] = []
    if keep_user_lines:
        carried = _carry_lines(user_preamble)
        lines.extend(carried)
    if "pgfplots" in pkgs and not any("compat" in c for c in carried):
        lines.append("\\pgfplotsset{compat=newest}")
    return "\n".join(x for x in lines if x) + "\n"


def _parse_preview_boxes(log: str) -> dict[int, tuple[float, float, float]]:
    out: dict[int, tuple[float, float, float]] = {}
    for m in _PREVIEW_LOG_RE.finditer(log):
        out[int(m.group(1))] = (float(m.group(2)), float(m.group(3)), float(m.group(4)))
    return out


#: Vá rỗng những gói này thì hỏng luôn cơ chế — thà báo lỗi rõ ràng còn hơn.
#: Những gói mà chính cơ chế này cần — thiếu là phải báo người dùng cài, không gỡ được.
_CORE_FOR_MECHANISM = {"preview", "varwidth", "standalone"}
#: Tuyệt đối không tạo khung rỗng cho những lệnh này — chúng do TikZ/pgfplots định nghĩa
#: ngay trong môi trường vẽ; đè lên là hỏng cả hình.
_NEVER_STUB = {"draw", "node", "fill", "filldraw", "path", "clip", "shade", "shadedraw",
               "coordinate", "addplot", "addplot3", "matrix", "pic", "foreach", "scope",
               "useasboundingbox", "pattern", "begin", "end", "item", "hline", "\\"}
NEVER_SHIM = {"preview", "varwidth", "standalone", "graphicx", "graphics", "tikz", "pgf",
              "pgfplots", "pgfcore", "xcolor", "color", "amsmath", "amssymb", "fontspec",
              "inputenc", "fontenc", "babel", "geometry", "array", "longtable", "tabularx"}


def _first_missing_file(log: str) -> tuple[str, str] | None:
    """Trả về (tên, phần mở rộng) của file TeX đầu tiên bị thiếu."""
    m = _MISSING_ANY_RE.search(log or "") or _MISSING_PKG_RE.search(log or "") or _MISSING_FILE_RE.search(log or "")
    if not m:
        return None
    raw = m.group(1).strip()
    if "." in raw:
        name, _, ext = raw.rpartition(".")
    else:
        name, ext = raw, "sty"
    if not re.fullmatch(r"[A-Za-z0-9@_\-]+", name):
        return None
    return name, ext.lower()


def _first_missing_package(log: str) -> str | None:
    got = _first_missing_file(log)
    if not got:
        return None
    name, ext = got
    if ext != "sty":
        return None
    if not re.fullmatch(r"[A-Za-z0-9@_\-]+", name):
        return None
    if name.lower() in NEVER_SHIM:
        return None            # gói nền tảng — vá rỗng là hỏng, phải báo cho người dùng cài
    return name


#: Bản vá *có nội dung* cho các gói hay thiếu — đặc biệt là gói tiếng Việt, nếu vá rỗng
#: thì chữ trong ảnh sẽ mất dấu.
SPECIAL_SHIMS: dict[str, str] = {
    "vietnam": "\\RequirePackage[T5]{fontenc}\\RequirePackage[utf8]{inputenc}\n",
    "vntex": "\\RequirePackage[T5]{fontenc}\\RequirePackage[utf8]{inputenc}\n",
    "utf8vn": "\\RequirePackage[utf8]{inputenc}\n",
    "vietnamese": "\\RequirePackage[T5]{fontenc}\\RequirePackage[utf8]{inputenc}\n",
}

UNICODE_FALLBACK_PACKAGES = {"vietnam", "vntex", "vietnamese", "utf8vn"}
UNICODE_SHIM = (
    "\\RequirePackage{fontspec}\n"
    "\\IfFontExistsTF{TeX Gyre Termes}{\\setmainfont{TeX Gyre Termes}}"
    "{\\IfFontExistsTF{Latin Modern Roman}{\\setmainfont{Latin Modern Roman}}{}}\n"
)


def _is_unicode_engine(engine: str | None) -> bool:
    name = Path(engine).name.lower() if engine else ""
    return "xelatex" in name or "lualatex" in name


_SHIM_TEMPLATE = (
    "\\NeedsTeXFormat{{LaTeX2e}}\n"
    "\\ProvidesPackage{{{name}}}[2000/01/01 ban va tu dong cua latex2word]\n"
    "\\RequirePackage{{xkeyval}}\n"
    "\\DeclareOption*{{}}\\ProcessOptions\\relax\n"
)


def _write_shim(d: Path, name: str, engine: str | None = None) -> None:
    try:
        extra = SPECIAL_SHIMS.get(name.lower(), "")
        if name.lower() in UNICODE_FALLBACK_PACKAGES and _is_unicode_engine(engine):
            extra = UNICODE_SHIM
        body = _SHIM_TEMPLATE.format(name=name) + extra
        (d / f"{name}.sty").write_text(body, encoding="utf-8")
    except Exception:
        pass


# ---------- tuyến DVI: dvisvgm / dvipng (đi kèm sẵn trong MiKTeX & TeX Live) ----------
_SVG_SIZE_RE = re.compile(rb'<svg[^>]*?width=["\']([0-9.]+)(pt|px|mm|cm|in)?["\'][^>]*?height=["\']([0-9.]+)(pt|px|mm|cm|in)?["\']', re.S)
_UNIT_PT = {b"pt": 1.0, None: 1.0, b"px": 0.75, b"mm": 72 / 25.4, b"cm": 72 / 2.54, b"in": 72.0}


def _svg_size_pt(data: bytes) -> tuple[float, float]:
    m = _SVG_SIZE_RE.search(data[:4000])
    if not m:
        return 0.0, 0.0
    w = float(m.group(1)) * _UNIT_PT.get(m.group(2), 1.0)
    h = float(m.group(3)) * _UNIT_PT.get(m.group(4), 1.0)
    return w, h


def _png_size_px(data: bytes) -> tuple[int, int]:
    """Đọc kích thước PNG từ khối IHDR — thuần thư viện chuẩn, không cần Pillow."""
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return 0, 0
    import struct
    w, h = struct.unpack(">II", data[16:24])
    return int(w), int(h)


def _dvi_to_svgs(dvi: Path, n_pages: int) -> tuple[dict[int, bytes], dict[int, tuple[float, float]]]:
    exe = _which("dvisvgm")
    out: dict[int, bytes] = {}
    sizes: dict[int, tuple[float, float]] = {}
    if not exe:
        return out, sizes
    d = dvi.parent
    for bbox in ("preview", "min"):
        for f in d.glob("pg*.svg"):
            try:
                f.unlink()
            except Exception:
                pass
        try:
            subprocess.run([exe, "--page=1-", "--no-fonts", f"--bbox={bbox}",
                            "--output=pg%p", dvi.name],
                           cwd=str(d), timeout=600,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            continue
        for i in range(1, n_pages + 1):
            f = d / f"pg{i}.svg"
            if f.exists() and f.stat().st_size > 200:
                data = f.read_bytes()
                out[i] = data
                sizes[i] = _svg_size_pt(data)
        if len(out) >= n_pages:
            break
    return out, sizes


def _dvi_to_pngs(dvi: Path, n_pages: int, dpi: int) -> tuple[dict[int, bytes], dict[int, tuple[float, float]]]:
    exe = _which("dvipng")
    out: dict[int, bytes] = {}
    sizes: dict[int, tuple[float, float]] = {}
    if not exe:
        return out, sizes
    d = dvi.parent
    try:
        subprocess.run([exe, "-D", str(dpi), "-T", "tight", "-bg", "Transparent",
                        "-q", "-o", "pg%d.png", dvi.name],
                       cwd=str(d), timeout=600,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        return out, sizes
    for i in range(1, n_pages + 1):
        f = d / f"pg{i}.png"
        if f.exists() and f.stat().st_size > 100:
            data = f.read_bytes()
            out[i] = data
            px, py = _png_size_px(data)
            if px and py:
                sizes[i] = (px * 72.0 / dpi, py * 72.0 / dpi)
    return out, sizes


# ---------- thao tác PDF (ưu tiên thư viện Python, không bắt buộc công cụ ngoài) ----------
def _pdf_page_count(pdf: Path) -> int:
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf)).pages)
    except Exception:
        pass
    try:
        out = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, timeout=60).stdout.decode("utf-8", "replace")
        m = re.search(r"Pages:\s+(\d+)", out)
        return int(m.group(1)) if m else 0
    except Exception:
        return 0


def _extract_page(pdf: Path, index1: int, dest: Path) -> bool:
    try:
        from pypdf import PdfReader, PdfWriter
        r = PdfReader(str(pdf))
        if index1 > len(r.pages):
            return False
        w = PdfWriter()
        w.add_page(r.pages[index1 - 1])
        with open(dest, "wb") as fh:
            w.write(fh)
        return dest.exists()
    except Exception:
        pass
    exe = _which("pdfseparate")
    if exe:
        try:
            subprocess.run([exe, "-f", str(index1), "-l", str(index1), str(pdf), str(dest)], timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return dest.exists()
        except Exception:
            return False
    return False


def _pdf_page_size_pt(pdf: Path) -> tuple[float, float]:
    try:
        from pypdf import PdfReader
        box = PdfReader(str(pdf)).pages[0].mediabox
        return float(box.width), float(box.height)
    except Exception:
        return 0.0, 0.0


def _pdf_to_svg(pdf: Path) -> bytes:
    try:                       # PyMuPDF: xuất SVG thuần Python, không cần công cụ ngoài
        import fitz
        doc = fitz.open(str(pdf))
        svg = doc[0].get_svg_image(text_as_path=True)
        doc.close()
        if svg and len(svg) > 200:
            return svg.encode("utf-8")
    except Exception:
        pass
    exe = _which("dvisvgm")
    if exe:
        dest = pdf.with_suffix(".svg")
        try:
            subprocess.run([exe, "--pdf", "--no-fonts", "--exact-bbox", "-o", str(dest), str(pdf)],
                           timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if dest.exists() and dest.stat().st_size > 200:
                return dest.read_bytes()
        except Exception:
            pass
    exe = _which("pdftocairo")
    if exe:
        dest = pdf.with_suffix(".svg")
        subprocess.run([exe, "-svg", str(pdf), str(dest)], timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if dest.exists():
            return dest.read_bytes()
    ink = _which("inkscape")
    if ink:
        dest = pdf.with_suffix(".svg")
        subprocess.run([ink, str(pdf), "--export-type=svg", f"--export-filename={dest}"], timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if dest.exists():
            return dest.read_bytes()
    return b""


def _pdf_to_emf(pdf: Path) -> bytes:
    ink = _which("inkscape")
    if ink:
        dest = pdf.with_suffix(".emf")
        try:
            subprocess.run([ink, str(pdf), "--export-type=emf", f"--export-filename={dest}"], timeout=180,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if dest.exists() and dest.stat().st_size > 200:
                return dest.read_bytes()
        except Exception:
            pass
    soffice = _which("soffice", "libreoffice")
    if soffice:
        # SVG → EMF cho nét sạch hơn; nếu không có SVG thì chuyển thẳng từ PDF
        outdir = pdf.parent / "emf_out"
        outdir.mkdir(exist_ok=True)
        src = pdf
        svg = _pdf_to_svg(pdf)
        if svg:
            src = pdf.with_suffix(".svg")
        try:
            subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "emf",
                            str(src), "--outdir", str(outdir)], timeout=240,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            cand = outdir / (src.stem + ".emf")
            if cand.exists() and cand.stat().st_size > 200:
                return cand.read_bytes()
        except Exception:
            pass
    return b""


def _pdf_to_png(pdf: Path, dpi: int) -> bytes:
    """PDF 1 trang → PNG. Thứ tự ưu tiên theo mức độ "chắc chắn có trên máy người dùng"."""
    # (1) pypdfium2 — cài bằng pip, không cần công cụ ngoài, chạy được trên mọi Windows
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(str(pdf))
        bmp = doc[0].render(scale=dpi / 72.0, draw_annots=False)
        dest = pdf.with_suffix(".png")
        bmp.to_pil().save(dest)
        doc.close()
        if dest.exists() and dest.stat().st_size > 100:
            return dest.read_bytes()
    except Exception:
        pass
    # (2) PyMuPDF nếu người dùng đã có
    try:
        import fitz
        doc = fitz.open(str(pdf))
        pix = doc[0].get_pixmap(dpi=dpi, alpha=True)
        dest = pdf.with_suffix(".png")
        pix.save(str(dest))
        doc.close()
        if dest.exists():
            return dest.read_bytes()
    except Exception:
        pass
    # (3) Poppler
    for exe_names, extra in ((("pdftocairo",), ["-transp"]), (("pdftoppm",), [])):
        exe = _which(*exe_names)
        if not exe:
            continue
        stem = pdf.with_suffix("")
        try:
            subprocess.run([exe, "-png", "-r", str(dpi), "-singlefile", *extra, str(pdf), str(stem)],
                           timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            cand = Path(str(stem) + ".png")
            if cand.exists():
                return cand.read_bytes()
        except Exception:
            pass
    # (4) ImageMagick
    exe = _which("magick", "convert")
    if exe:
        dest = pdf.with_suffix(".png")
        try:
            subprocess.run([exe, "-density", str(dpi), str(pdf), "-background", "none", str(dest)],
                           timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if dest.exists():
                return dest.read_bytes()
        except Exception:
            pass
    return b""
