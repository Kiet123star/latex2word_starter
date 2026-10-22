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
_MISSING_FILE_RE = re.compile(r"(?:File|file)\s+[`'\"]([A-Za-z0-9@_\-.]+?)(?:\.sty)?['\"`]\s+not found")
_MISSING_PKG_RE = re.compile(r"LaTeX Error: File [`'\"]([^'\"`]+)['\"`] not found")


def _log(msg: str) -> None:
    try:
        print(f"[tex_truth] {msg}", file=sys.stderr, flush=True)
    except Exception:
        pass


# ───────────────────────── dò tìm công cụ ngoài ─────────────────────────
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


def describe_toolchain() -> dict[str, Any]:
    """Báo cáo công cụ sẵn có — dùng để hiện trạng thái trên giao diện."""
    engine = tex_engine_available()
    soffice = _which("soffice", "libreoffice")
    inkscape = _which("inkscape")
    pdftocairo = _which("pdftocairo")
    pdftoppm = _which("pdftoppm")
    magick = _which("magick", "convert")
    dvisvgm = _which("dvisvgm")
    if pdftocairo or dvisvgm:
        best = "svg"          # vector, chữ đã thành đường path → không phụ thuộc font máy
    elif inkscape or soffice:
        best = "emf"
    elif pdftoppm or magick:
        best = "png"
    else:
        best = "none"
    return {
        "tex_engine": engine,
        "soffice": soffice,
        "inkscape": inkscape,
        "dvisvgm": dvisvgm,
        "pdftocairo": pdftocairo,
        "pdftoppm": pdftoppm,
        "magick": magick,
        "best_vector_format": best,
        "ready": bool(engine) and best != "none",
    }


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


@dataclass
class _Snippet:
    key: str
    body: str
    inline: bool
    label: str = ""


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
        """Biên dịch tất cả khối đã đăng ký. Không bao giờ ném lỗi ra ngoài."""
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
        if not self.engine:
            for sn in pending:
                out[sn.key] = RenderedAsset(key=sn.key, inline=sn.inline, ok=False,
                                            error="Không tìm thấy trình biên dịch TeX (pdflatex/xelatex).")
                self.stats["failed"] += 1
            return out
        # bước 1 — gộp một lần biên dịch cho tất cả
        batch = self._compile_batch(pending)
        if batch is not None:
            pdf, boxes = batch
            got = self._slice_pdf(pdf, pending, boxes)
            out.update(got)
            missing = [sn for sn in pending if not got.get(sn.key, RenderedAsset(key=sn.key)).ok]
        else:
            missing = list(pending)
        # bước 2 — khối nào hỏng thì biên dịch riêng (chậm hơn nhưng cứu được phần còn lại)
        for sn in missing:
            single = self._compile_batch([sn], tag=f"solo_{sn.key}")
            asset = RenderedAsset(key=sn.key, inline=sn.inline, ok=False, error="Biên dịch TeX thất bại.")
            if single is not None:
                pdf, boxes = single
                got = self._slice_pdf(pdf, [sn], boxes)
                asset = got.get(sn.key, asset)
            out[sn.key] = asset
            if not asset.ok:
                self.stats["failed"] += 1
        for key, asset in out.items():
            if asset.ok:
                self._cache_store(asset)
        return out

    # ---------- biên dịch ----------
    def _compile_batch(self, snippets: list[_Snippet], tag: str = "batch") -> tuple[Path, dict[int, tuple[float, float, float]]] | None:
        d = self.workdir / tag
        d.mkdir(parents=True, exist_ok=True)
        tex_path = d / "l2w_truth.tex"
        tex_path.write_text(self._build_source(snippets), encoding="utf-8")
        shims: set[str] = set()
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
                self.stats["compiled"] += len(snippets)
                return pdf, boxes
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

    def _run_engine(self, tex_path: Path, cwd: Path) -> str:
        self.stats["tex_runs"] += 1
        env = dict(os.environ)
        env.setdefault("TEXINPUTS", f"{cwd}{os.pathsep}{os.pathsep}")
        env["max_print_line"] = "1000"
        cmd = [self.engine or "pdflatex", "-interaction=nonstopmode", "-file-line-error", tex_path.name]
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

    def _build_source(self, snippets: list[_Snippet]) -> str:
        parts = [
            "\\documentclass[12pt]{article}\n",
            "\\usepackage[active,tightpage]{preview}\n",
            "\\setlength\\PreviewBorder{1pt}\n",
            "\\IfFileExists{varwidth.sty}{\\usepackage{varwidth}}{}\n",
            "\\makeatletter\\@ifundefined{varwidth}{\\newenvironment{l2wbox}[1]"
            "{\\begin{minipage}{#1}}{\\end{minipage}}}"
            "{\\newenvironment{l2wbox}[1]{\\begin{varwidth}{#1}}{\\end{varwidth}}}\\makeatother\n",
            self.preamble,
            "\n\\pagestyle{empty}\n",
            "\\begin{document}\n",
        ]
        for sn in snippets:
            parts.append("\\begin{preview}%\n")
            if sn.inline:
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


def extract_tex_errors(log: str, limit: int = 8) -> list[str]:
    seen: list[str] = []
    for m in _TEX_ERROR_RE.finditer(log or ""):
        line = m.group(0).strip()
        if line and line not in seen:
            seen.append(line)
        if len(seen) >= limit:
            break
    return seen


def _parse_preview_boxes(log: str) -> dict[int, tuple[float, float, float]]:
    out: dict[int, tuple[float, float, float]] = {}
    for m in _PREVIEW_LOG_RE.finditer(log):
        out[int(m.group(1))] = (float(m.group(2)), float(m.group(3)), float(m.group(4)))
    return out


def _first_missing_package(log: str) -> str | None:
    m = _MISSING_PKG_RE.search(log) or _MISSING_FILE_RE.search(log)
    if not m:
        return None
    name = m.group(1).strip()
    if name.endswith(".sty"):
        name = name[:-4]
    if not re.fullmatch(r"[A-Za-z0-9@_\-]+", name):
        return None
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
    exe = _which("pdftocairo")
    if exe:
        stem = pdf.with_suffix("")
        subprocess.run([exe, "-png", "-r", str(dpi), "-singlefile", "-transp", str(pdf), str(stem)],
                       timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        cand = Path(str(stem) + ".png")
        if cand.exists():
            return cand.read_bytes()
    exe = _which("pdftoppm")
    if exe:
        stem = pdf.with_suffix("")
        subprocess.run([exe, "-png", "-r", str(dpi), "-singlefile", str(pdf), str(stem)],
                       timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        cand = Path(str(stem) + ".png")
        if cand.exists():
            return cand.read_bytes()
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(str(pdf))
        bmp = doc[0].render(scale=dpi / 72.0)
        dest = pdf.with_suffix(".png")
        bmp.to_pil().save(dest)
        return dest.read_bytes()
    except Exception:
        pass
    exe = _which("magick", "convert")
    if exe:
        dest = pdf.with_suffix(".png")
        subprocess.run([exe, "-density", str(dpi), str(pdf), "-background", "none", "-flatten", str(dest)],
                       timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if dest.exists():
            return dest.read_bytes()
    return b""
