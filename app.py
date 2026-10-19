#!/usr/bin/env python3
"""Giao diện Web (Streamlit) & Bộ xử lý Hybrid LaTeX-to-Word (Tích hợp Cloud TikZ + Direct DOCX Image Injection).

Đặt file `app.py` này cùng thư mục với `latex2word.py` (trong `latex2word_starter/`).
Chạy ứng dụng bằng lệnh:
    streamlit run app.py
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import streamlit as st

# Bắt buộc dùng UTF-8 cho mọi tiến trình con trên Windows (tránh lỗi cp1252 trên Python 3.14)
os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import numpy as np
    from PIL import Image, ImageChops, ImageDraw, ImageEnhance

    HAS_VISION = True
except ImportError:
    HAS_VISION = False

try:
    import pypdfium2 as pdfium

    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False


# ==============================================================================
# 1. TỰ ĐỘNG DÒ TÌM CÔNG CỤ TRÊN WINDOWS (AUTO-DISCOVER PATH)
# ==============================================================================
def auto_discover_windows_tools() -> None:
    """Tự động tìm Pandoc, MiKTeX, TeX Live, Poppler, LibreOffice trên Windows nếu chưa có trong PATH."""
    if os.name != "nt":
        return

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    user_profile = os.environ.get("USERPROFILE", "")

    candidate_dirs = [
        Path(local_app_data) / "Pandoc",
        Path(program_files) / "Pandoc",
        Path(local_app_data) / r"Programs\MiKTeX\miktex\bin\x64",
        Path(program_files) / r"MiKTeX\miktex\bin\x64",
        Path(program_files) / r"LibreOffice\program",
        Path(program_files_x86) / r"LibreOffice\program",
        Path(user_profile) / r"scoop\shims",
        Path(r"C:\ProgramData\chocolatey\bin"),
    ]

    for root_dir in [Path("C:/texlive"), Path(program_files) / "poppler", Path("C:/poppler")]:
        if root_dir.exists():
            candidate_dirs.extend(root_dir.glob("**/bin"))
            candidate_dirs.extend(root_dir.glob("**/bin/windows"))

    current_path = os.environ.get("PATH", "")
    added_paths = [str(d) for d in candidate_dirs if d.exists() and str(d) not in current_path]
    if added_paths:
        os.environ["PATH"] = current_path + os.pathsep + os.pathsep.join(added_paths)


auto_discover_windows_tools()

# ==============================================================================
# CẤU HÌNH & NAMESPACE XML
# ==============================================================================
BASE_DIR = Path(__file__).resolve().parent
CONVERTER_SCRIPT = BASE_DIR / "latex2word.py"
EXAMPLE_TEX = BASE_DIR / "example.tex"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

DEFAULT_PREAMBLE = r"""\documentclass[12pt]{article}
\usepackage{amsmath,amssymb,amsfonts}
\usepackage{tikz,pgfplots}
\usepackage{booktabs,array,multirow,longtable}
\usepackage{graphicx,xcolor}
\pgfplotsset{compat=1.18}
\usetikzlibrary{calc,angles,quotes,intersections,arrows,arrows.meta,patterns,positioning}
"""

DEFAULT_LATEX_TEMPLATE = r"""\documentclass[12pt]{article}
\usepackage{amsmath,amssymb}
\usepackage{tikz}
\usepackage{pgfplots}
\usepackage{booktabs}
\pgfplotsset{compat=1.18}

\title{Mẫu Kiểm Thử LaTeX sang Word (.docx)}
\author{Hệ Thống Chuyển Đổi Hybrid}
\date{\today}

\begin{document}
\maketitle

\section{Công thức Toán học (OMML)}
Phương trình bậc hai và tích phân xác định:
\[
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}, \quad
I = \int_{0}^{\infty} e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}
\]

\section{Đồ thị Hàm số (PGFPlots)}
\begin{center}
\begin{tikzpicture}
\begin{axis}[
    title={Đồ thị hàm số $f(x) = x^2 - 2x - 3$},
    xlabel={$x$},
    ylabel={$f(x)$},
    grid=major,
    width=8cm,
    height=6cm,
    domain=-2.5:4.5,
    samples=100
]
\addplot[blue, thick] {x^2 - 2*x - 3};
\end{axis}
\end{tikzpicture}
\end{center}

\section{Bảng Dữ liệu}
\begin{table}[h!]
\centering
\begin{tabular}{lcc}
\toprule
Thành phần & Công nghệ xử lý & Đầu ra Word \\
\midrule
Văn bản & Pandoc & Text có thể chỉnh sửa \\
Toán học & LaTeX $\to$ OMML & Equation gốc của Word \\
Đồ họa TikZ & Hybrid Local / Cloud Compiler & Ảnh độ phân giải cao \\
\bottomrule
\end{tabular}
\caption{Bảng tổng hợp các thành phần chuyển đổi.}
\end{table}

\end{document}
"""


@dataclass
class ValidationLayerResult:
    """Kết quả kiểm tra từng lớp trong 5 lớp xác thực (5 Validation Layers)."""

    layer_name: str
    status: str  # "PASS", "WARN", "FAIL", "SKIP"
    summary: str
    details: dict[str, Any] = field(default_factory=dict)


def run_cmd_utf8(cmd: list[str], cwd: Path, timeout: int = 120) -> subprocess.CompletedProcess[str]:
    """Chạy tiến trình con với bảng mã UTF-8 an toàn tuyệt đối trên Windows."""
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run(
        cmd,
        cwd=cwd,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )


# ==============================================================================
# 2. BỘ BIÊN DỊCH TIKZ ĐA TẦNG (LOCAL XELATEX + CLOUD QUICKLATEX / TEXLIVE.NET)
# ==============================================================================
def convert_pdf_to_png(pdf_path: Path, out_png_path: Path, dpi: int = 400) -> bool:
    """Chuyển 1 trang PDF sang PNG bằng pypdfium2 hoặc pdftocairo (Poppler)."""
    if HAS_PDFIUM:
        try:
            pdf = pdfium.PdfDocument(str(pdf_path))
            page = pdf[0]
            bitmap = page.render(scale=dpi / 72.0)
            pil_image = bitmap.to_pil()
            pil_image.save(str(out_png_path))
            page.close()
            pdf.close()
            if out_png_path.exists():
                return True
        except Exception:
            pass

    pdftocairo_bin = shutil.which("pdftocairo")
    if pdftocairo_bin:
        prefix = out_png_path.with_suffix("")
        proc = run_cmd_utf8(
            [pdftocairo_bin, "-png", "-singlefile", "-r", str(dpi), str(pdf_path), str(prefix)],
            cwd=pdf_path.parent,
        )
        if proc.returncode == 0 and out_png_path.exists():
            return True

    return False


def compile_tikz_via_quicklatex(
    tikz_code: str, extra_preamble: str, out_png_path: Path
) -> tuple[bool, str]:
    """Biên dịch khối TikZ/PGFPlots trực tiếp ra ảnh PNG qua dịch vụ QuickLaTeX API.

    Ưu điểm: Trả thẳng về file ảnh `.png` sắc nét mà KHÔNG cần máy cài MiKTeX hay Poppler!
    """
    preamble_parts = [
        r"\usepackage{amsmath,amssymb,amsfonts}",
        r"\usepackage{tikz}",
        r"\usepackage{xcolor}",
        r"\usetikzlibrary{calc,angles,quotes,intersections,arrows,arrows.meta,patterns,positioning,decorations.markings,shapes.geometric}",
    ]
    if "\\begin{axis}" in tikz_code or "\\addplot" in tikz_code:
        preamble_parts.append(r"\usepackage{pgfplots}\pgfplotsset{compat=1.18}")
    if "\\tkzTab" in tikz_code:
        preamble_parts.append(r"\usepackage{tkz-tab}")
    if "\\tkzDef" in tikz_code or "\\tkzDraw" in tikz_code:
        preamble_parts.append(r"\usepackage{tkz-euclide}")
    if any(ord(ch) > 127 for ch in tikz_code):
        preamble_parts.insert(0, r"\usepackage[utf8]{inputenc}\usepackage[T5]{fontenc}")
    if extra_preamble.strip():
        preamble_parts.append(extra_preamble.strip())

    full_preamble = "\n".join(preamble_parts)
    payload = {
        "formula": tikz_code.strip(),
        "fsize": "18px",
        "fcolor": "000000",
        "mode": "0",
        "out": "1",
        "remhost": "quicklatex.com",
        "preamble": full_preamble,
    }
    encoded_body = urllib.parse.urlencode(payload, quote_via=urllib.parse.quote).encode("utf-8")
    req = urllib.request.Request(
        "https://quicklatex.com/latex3.f",
        data=encoded_body,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=25) as resp:
            text_resp = resp.read().decode("utf-8", errors="replace").strip()
        lines = text_resp.splitlines()
        if len(lines) >= 2 and lines[0].strip() == "0":
            img_url = lines[1].split()[0]
            if img_url and not img_url.endswith("error.png"):
                img_req = urllib.request.Request(img_url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(img_req, timeout=20) as img_resp:
                    png_bytes = img_resp.read()
                if png_bytes.startswith(b"\x89PNG"):
                    out_png_path.write_bytes(png_bytes)
                    return True, "QuickLaTeX Cloud PNG"
        err_detail = "\n".join(lines[2:]) if len(lines) > 2 else text_resp
        return False, f"QuickLaTeX: {err_detail[:300]}"
    except Exception as exc:
        return False, f"QuickLaTeX kết nối thất bại: {exc}"


def compile_tikz_via_texlivenet(
    tikz_code: str, extra_preamble: str, workdir: Path, out_png_path: Path, dpi: int = 400
) -> tuple[bool, str]:
    """Biên dịch khối TikZ/PGFPlots/tkz-tab qua máy chủ TeXLive.net CGI (hỗ trợ đầy đủ gói tiếng Việt & Toán)."""
    if not (HAS_PDFIUM or shutil.which("pdftocairo")):
        return False, "Cần cài pypdfium2 (`pip install pypdfium2`) để chuyển PDF từ TeXLive.net sang PNG."

    standalone_doc = (
        "\\documentclass[tikz,border=4pt]{standalone}\n"
        "\\usepackage[utf8]{inputenc}\n"
        "\\usepackage[T5]{fontenc}\n"
        "\\usepackage{amsmath,amssymb,amsfonts}\n"
        "\\usepackage{pgfplots,xcolor}\n"
        "\\pgfplotsset{compat=1.18}\n"
        "\\usetikzlibrary{calc,angles,quotes,intersections,arrows,arrows.meta,patterns,positioning,decorations.markings,shapes.geometric,3d}\n"
        + ("\\usepackage{tkz-tab}\n" if "\\tkzTab" in tikz_code else "")
        + ("\\usepackage{tkz-euclide}\n" if ("\\tkzDef" in tikz_code or "\\tkzDraw" in tikz_code) else "")
        + extra_preamble
        + "\n\\begin{document}\n"
        + tikz_code
        + "\n\\end{document}\n"
    )

    boundary = f"----WebKitFormBoundary{uuid.uuid4().hex}"
    fields = [
        ("filename[]", "document.tex"),
        ("filecontents[]", standalone_doc),
        ("engine", "pdflatex"),
        ("return", "pdf"),
    ]
    body_chunks: list[bytes] = []
    for name, val in fields:
        body_chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        body_chunks.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"))
        body_chunks.append(f"{val}\r\n".encode("utf-8"))
    body_chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    body_bytes = b"".join(body_chunks)

    req = urllib.request.Request(
        "https://texlive.net/cgi-bin/latexcgi",
        data=body_bytes,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            ctype = resp.headers.get("Content-Type", "")
            resp_bytes = resp.read()
        if "pdf" in ctype.lower() or resp_bytes.startswith(b"%PDF"):
            tmp_pdf = workdir / f"{out_png_path.stem}.pdf"
            tmp_pdf.write_bytes(resp_bytes)
            if convert_pdf_to_png(tmp_pdf, out_png_path, dpi=dpi):
                return True, "TeXLive.net Cloud PDF->PNG"
        return False, "TeXLive.net không trả về PDF hợp lệ."
    except Exception as exc:
        return False, f"TeXLive.net lỗi kết nối: {exc}"


def render_single_tikz_figure(
    idx: int,
    tikz_code: str,
    custom_preamble_str: str,
    tikz_build_dir: Path,
    workdir: Path,
    dpi: int,
    latex_engine: str | None,
) -> tuple[int, Path | None, str, str | None]:
    """Biên dịch 1 hình TikZ theo thứ tự: Local XeLaTeX/pdfLaTeX -> TeXLive.net -> QuickLaTeX."""
    out_png_name = f"rendered_tikz_{idx:04d}.png"
    out_png_path = workdir / out_png_name

    # 1. Thử biên dịch bằng Local XeLaTeX / pdfLaTeX nếu máy đã cài
    if latex_engine:
        standalone_tex = (
            "\\documentclass[tikz,border=4pt]{standalone}\n"
            "\\usepackage{amsmath,amssymb,amsfonts}\n"
            "\\usepackage{pgfplots,xcolor}\n"
            "\\pgfplotsset{compat=1.18}\n"
            "\\usetikzlibrary{calc,angles,quotes,intersections,arrows,arrows.meta,patterns,positioning,decorations.markings,shapes.geometric}\n"
            + ("\\usepackage{tkz-tab}\n" if "\\tkzTab" in tikz_code else "")
            + ("\\usepackage{tkz-euclide}\n" if ("\\tkzDef" in tikz_code or "\\tkzDraw" in tikz_code) else "")
            + custom_preamble_str
            + "\n\\begin{document}\n"
            + tikz_code
            + "\n\\end{document}\n"
        )
        sub_tex_path = tikz_build_dir / f"tikz_fig_{idx:04d}.tex"
        sub_pdf_path = tikz_build_dir / f"tikz_fig_{idx:04d}.pdf"
        sub_tex_path.write_text(standalone_tex, encoding="utf-8")

        cmd = [
            latex_engine,
            "-interaction=nonstopmode",
            "-halt-on-error",
            f"-output-directory={tikz_build_dir}",
            str(sub_tex_path),
        ]
        proc = run_cmd_utf8(cmd, cwd=workdir, timeout=60)
        if proc.returncode == 0 and sub_pdf_path.exists():
            if convert_pdf_to_png(sub_pdf_path, out_png_path, dpi=dpi):
                return idx, out_png_path, f"Hình #{idx}: Biên dịch bằng Local `{Path(latex_engine).name}` ({dpi} DPI).", None

    # 2. Máy chưa cài LaTeX hoặc thiếu gói local -> Tự động gọi Cloud Compiler (QuickLaTeX & TeXLive.net)
    if HAS_PDFIUM or shutil.which("pdftocairo"):
        ok_tl, msg_tl = compile_tikz_via_texlivenet(tikz_code, custom_preamble_str, tikz_build_dir, out_png_path, dpi=dpi)
        if ok_tl and out_png_path.exists():
            return idx, out_png_path, f"Hình #{idx}: Biên dịch thành công qua `{msg_tl}` ({dpi} DPI).", None

    ok_ql, msg_ql = compile_tikz_via_quicklatex(tikz_code, custom_preamble_str, out_png_path)
    if ok_ql and out_png_path.exists():
        return idx, out_png_path, f"Hình #{idx}: Biên dịch thành công qua `{msg_ql}`.", None

    # Nếu chưa cài pypdfium2 nên bước 2a bị bỏ qua, thử lại TeXLive.net nếu có thể
    return idx, None, "", f"Hình #{idx} không thể biên dịch ({msg_ql})"


# ==============================================================================
# 3. TIỀN XỬ LÝ & ĐẶT MARKER HÌNH VẼ TRƯỚC KHI CHẠY PANDOC / LATEX2WORD.PY
# ==============================================================================
def normalize_and_wrap_latex(raw_tex: str) -> tuple[str, list[str]]:
    """Tự động chuẩn hóa các macro đề toán Việt Nam (\\immini, \\heva, \\hoac) và bọc khung tài liệu."""
    notes: list[str] = []
    tex = raw_tex.strip()

    has_tikz_env = "\\begin{tikzpicture}" in tex
    looks_like_raw_tikz = bool(
        re.search(r"\\(?:draw|fill|coordinate|node|path|clip|shade|filldraw)\b", tex)
    )
    if not has_tikz_env and looks_like_raw_tikz and "\\begin{document}" not in tex:
        tex = "\\begin{center}\n\\begin{tikzpicture}\n" + tex + "\n\\end{tikzpicture}\n\\end{center}"
        notes.append("Phát hiện các lệnh vẽ TikZ rời rạc: Đã tự động bọc trong môi trường `\\begin{tikzpicture}`.")

    if "\\begin{document}" not in tex:
        if "\\documentclass" in tex:
            tex = tex + "\n\\begin{document}\n\\end{document}\n"
        else:
            tex = DEFAULT_PREAMBLE + "\n\\begin{document}\n" + tex + "\n\\end{document}\n"
            notes.append("Đã tự động bổ sung khung `\\documentclass{article}` cho đoạn mã LaTeX.")

    # Hỗ trợ macro \immini{văn bản}{hình vẽ} rất phổ biến trong đề thi Toán THPT Việt Nam (gói ex_test)
    if "\\immini" in tex and "\\newcommand{\\immini}" not in tex and "\\providecommand{\\immini}" not in tex:
        tex = tex.replace(
            "\\begin{document}",
            "\\providecommand{\\immini}[2]{#1\\par\\noindent #2\\par}\n\\begin{document}",
            1,
        )
        notes.append("Phát hiện lệnh `\\immini{...}{...}`: Đã tự động mở rộng để giữ nguyên cả đề bài và hình vẽ.")

    return tex, notes


def extract_and_render_all_tikz(
    tex_content: str, workdir: Path, dpi: int = 400
) -> tuple[str, dict[str, Path], list[str], list[str]]:
    """Biên dịch song song tất cả các khối TikZ/PGFPlots và đặt Marker `TIKZFIGMARKERxxxx` vào văn bản.

    Sau khi Pandoc / `latex2word.py` tạo xong file `.docx`, hàm `inject_tikz_images_into_docx`
    sẽ tìm từng Marker này và nhúng thẳng ảnh `.png` vào đúng vị trí trong Word!
    """
    logs: list[str] = []
    warnings: list[str] = []
    marker_to_png: dict[str, Path] = {}

    doc_split = tex_content.split("\\begin{document}", 1)
    if len(doc_split) < 2:
        return tex_content, marker_to_png, logs, warnings

    raw_preamble, body = doc_split[0], doc_split[1]
    custom_preamble_lines = []
    for line in raw_preamble.splitlines():
        stripped = line.strip()
        if any(
            stripped.startswith(cmd)
            for cmd in ("\\newcommand", "\\renewcommand", "\\def", "\\definecolor", "\\tikzset", "\\tikzstyle")
        ):
            custom_preamble_lines.append(line)
    custom_preamble_str = "\n".join(custom_preamble_lines)

    # Bắt cả trường hợp $\begin{tikzpicture}...\end{tikzpicture}$ (bọc trong dấu $)
    tikz_pattern = re.compile(
        r"\$?\s*\\begin\{tikzpicture\}(?:\[[^\]]*\])?.*?\\end\{tikzpicture\}\s*\$?",
        re.S,
    )
    matches = list(tikz_pattern.finditer(body))
    if not matches:
        return tex_content, marker_to_png, logs, warnings

    latex_engine = shutil.which("xelatex") or shutil.which("pdflatex")
    if not latex_engine:
        logs.append(
            "Máy chưa cài `xelatex`/`pdflatex` cục bộ — Tự động kích hoạt **Cloud TikZ Compiler (QuickLaTeX / TeXLive.net)** để vẽ toàn bộ hình TikZ!"
        )

    tikz_build_dir = workdir / "_smart_tikz_build"
    tikz_build_dir.mkdir(exist_ok=True)

    # Trích xuất mã thuần của từng khối tikzpicture
    inner_tikz_pattern = re.compile(r"\\begin\{tikzpicture\}(?:\[[^\]]*\])?.*?\\end\{tikzpicture\}", re.S)
    tasks: list[tuple[int, str]] = []
    for idx, match in enumerate(matches, start=1):
        full_match_str = match.group(0)
        inner_m = inner_tikz_pattern.search(full_match_str)
        clean_tikz = inner_m.group(0) if inner_m else full_match_str
        tasks.append((idx, clean_tikz))

    rendered_map: dict[int, Path] = {}
    # Chạy song song tối đa 4 luồng để tăng tốc độ biên dịch 14+ hình vẽ trong đề thi
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(
                render_single_tikz_figure,
                idx,
                code,
                custom_preamble_str,
                tikz_build_dir,
                workdir,
                dpi,
                latex_engine,
            ): idx
            for idx, code in tasks
        }
        for fut in as_completed(futures):
            idx, png_path, log_msg, warn_msg = fut.result()
            if png_path and png_path.exists():
                rendered_map[idx] = png_path
            if log_msg:
                logs.append(log_msg)
            if warn_msg:
                warnings.append(warn_msg)

    logs.sort()
    warnings.sort()

    # Thay thế từng khối tikzpicture bằng Marker thuần chữ `TIKZFIGMARKER0001`
    new_body = body
    for idx, match in reversed(list(enumerate(matches, start=1))):
        marker = f"TIKZFIGMARKER{idx:04d}"
        if idx in rendered_map:
            marker_to_png[marker] = rendered_map[idx]
            replacement = f"\n\n\\noindent {marker}\n\n"
        else:
            replacement = f"\n\n\\noindent [Hinh ve TikZ so {idx} can kiem tra]\n\n"
        start, end = match.span()
        new_body = new_body[:start] + replacement + new_body[end:]

    updated_tex = raw_preamble + "\\begin{document}" + new_body
    return updated_tex, marker_to_png, logs, warnings


# ==============================================================================
# 4. NHÚNG TRỰC TIẾP ẢNH TIKZ VÀO FILE WORD (.DOCX) BẰNG PYTHON-DOCX
# ==============================================================================
def inject_tikz_images_into_docx(docx_path: Path, marker_to_png: dict[str, Path]) -> int:
    """Tìm tất cả các Marker `TIKZFIGMARKERxxxx` trong đoạn văn và bảng Word để chèn ảnh PNG trực tiếp."""
    if not HAS_DOCX or not docx_path.exists() or not marker_to_png:
        return 0

    doc = docx.Document(str(docx_path))
    inserted_count = 0

    def process_paragraph(para: Any) -> int:
        count = 0
        p_text = para.text
        for marker, png_path in marker_to_png.items():
            if marker in p_text and png_path.exists():
                # Xóa chuỗi marker bên trong các thẻ <w:t> mà KHÔNG chạm vào công thức <m:oMath>
                for wt in para._p.findall(".//w:t", NS):
                    if wt.text and marker in wt.text:
                        wt.text = wt.text.replace(marker, "")

                # Tính chiều rộng hiển thị hợp lý trong Word dựa trên kích thước ảnh gốc
                width_inches = 3.2
                if HAS_VISION:
                    try:
                        with Image.open(png_path) as im:
                            w_px, h_px = im.size
                            # Quy đổi kích thước cân đối cho hình vẽ đề toán (từ 2.2 đến 4.5 inches)
                            est_in = (w_px / 220.0)
                            width_inches = max(2.2, min(4.5, est_in))
                    except Exception:
                        pass

                if not para.text.strip():
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(str(png_path), width=Inches(width_inches))
                count += 1
        return count

    # 1. Quét toàn bộ các đoạn văn (Paragraphs) ngoài bảng
    for p in doc.paragraphs:
        inserted_count += process_paragraph(p)

    # 2. Quét toàn bộ các ô bên trong Bảng (Tables) — phòng trường hợp hình nằm trong bảng chia cột
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    inserted_count += process_paragraph(p)

    doc.save(str(docx_path))
    return inserted_count


# ==============================================================================
# 5. ENGINE DỰ PHÒNG THUẦN PYTHON & HẬU XỬ LÝ BẢNG OOXML
# ==============================================================================
def pure_python_latex_to_docx(tex_content: str, workdir: Path, output_docx: Path) -> bool:
    """Bộ chuyển đổi dự phòng dùng trực tiếp `python-docx` khi Pandoc hoặc `latex2word.py` gặp sự cố."""
    if not HAS_DOCX:
        return False

    doc = docx.Document()
    body = tex_content.split("\\begin{document}", 1)[-1].split("\\end{document}", 1)[0]

    title_match = re.search(r"\\title\{([^{}]+)\}", tex_content)
    if title_match:
        doc.add_heading(title_match.group(1).strip(), level=0)

    for block in re.split(r"\n\s*\n", body):
        cleaned = block.strip()
        if not cleaned or cleaned == "\\maketitle":
            continue
        doc.add_paragraph(cleaned)

    doc.save(str(output_docx))
    return output_docx.exists()


def postprocess_docx_tables(docx_path: Path) -> dict[str, int]:
    """Chuẩn hóa đường viền, căn chỉnh ô gộp và cấu trúc bảng Word bằng python-docx."""
    stats = {"tables_processed": 0, "cells_formatted": 0, "merged_cells_detected": 0}
    if not HAS_DOCX or not docx_path.exists():
        return stats

    doc = docx.Document(str(docx_path))
    for table in doc.tables:
        stats["tables_processed"] += 1
        tbl_pr = table._tbl.tblPr
        if tbl_pr is not None:
            borders = tbl_pr.find(qn("w:tblBorders"))
            if borders is None:
                borders = OxmlElement("w:tblBorders")
                for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
                    border = OxmlElement(f"w:{border_name}")
                    border.set(qn("w:val"), "single")
                    border.set(qn("w:sz"), "4")
                    border.set(qn("w:space"), "0")
                    border.set(qn("w:color"), "BFBFBF")
                    borders.append(border)
                tbl_pr.append(borders)

        for row in table.rows:
            for cell in row.cells:
                stats["cells_formatted"] += 1
                tc_pr = cell._tc.get_or_add_tcPr()
                if tc_pr.find(qn("w:gridSpan")) is not None or tc_pr.find(qn("w:vMerge")) is not None:
                    stats["merged_cells_detected"] += 1

    doc.save(str(docx_path))
    return stats


# ==============================================================================
# 6. NĂM TẦNG KIỂM THỬ TỰ ĐỘNG (5-LAYER AUTOMATIC ERROR DETECTION)
# ==============================================================================
def strip_latex_commands(tex: str) -> str:
    """Lược bỏ các lệnh điều khiển LaTeX để lấy văn bản thuần phục vụ đối chiếu Layer 1."""
    text = re.sub(r"%.*$", "", tex, flags=re.MULTILINE)
    text = re.sub(r"\\begin\{(?:tikzpicture|axis)\}.*?\\end\{(?:tikzpicture|axis)\}", " ", text, flags=re.S)
    text = re.sub(r"\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?", " ", text)
    text = re.sub(r"[{}$&#^_~]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def run_five_layer_validation(
    tex_content: str,
    docx_path: Path,
    starter_report: dict[str, Any],
    visual_diff_info: dict[str, Any] | None = None,
) -> list[ValidationLayerResult]:
    """Thực thi kiểm tra độc lập trên 5 tầng theo đúng kiến trúc mục tiêu."""
    results: list[ValidationLayerResult] = []
    if not docx_path.exists():
        return [
            ValidationLayerResult(
                layer_name="Tất cả các tầng",
                status="FAIL",
                summary="Không tìm thấy file .docx đầu ra để kiểm thử.",
            )
        ]

    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
        media_files = [f for f in zf.namelist() if f.startswith("word/media/")]

    root = ET.fromstring(xml_bytes)

    # --- LAYER 1: TEXT VERIFICATION ---
    docx_texts = [node.text for node in root.findall(".//w:t", NS) if node.text]
    docx_plain = " ".join(docx_texts).strip()
    source_plain = strip_latex_commands(tex_content)
    docx_word_count = len(docx_plain.split())
    src_word_count = len(source_plain.split())
    retention_ratio = 100.0 if src_word_count == 0 else round(min(docx_word_count / src_word_count, 1.0) * 100, 1)

    l1_status = "PASS" if (src_word_count == 0 or retention_ratio >= 50) else "WARN"
    results.append(
        ValidationLayerResult(
            layer_name="Layer 1 — Text Verification (Văn bản & Ký tự)",
            status=l1_status,
            summary=f"Trích xuất {docx_word_count} từ trong Word (tỷ lệ bảo toàn ước tính ~{retention_ratio}%).",
            details={
                "source_estimated_words": src_word_count,
                "docx_words": docx_word_count,
                "paragraphs_count": len(root.findall(".//w:p", NS)),
            },
        )
    )

    # --- LAYER 2: MATHEMATICAL VERIFICATION (OMML) ---
    omath_nodes = root.findall(".//m:oMath", NS)
    fractions = root.findall(".//m:f", NS)
    radicals = root.findall(".//m:rad", NS)
    superscripts = root.findall(".//m:sSup", NS)
    subscripts = root.findall(".//m:sSub", NS)
    matrices = root.findall(".//m:m", NS)
    nary_ops = root.findall(".//m:nary", NS)

    src_math_blocks = len(
        re.findall(r"(\\\[.*?\\\]|\$\$.*?\$\$|\\begin\{(equation\*?|align\*?|gather\*?)\})", tex_content, re.S)
    )
    l2_status = "PASS" if (src_math_blocks == 0 or len(omath_nodes) > 0) else "WARN"
    results.append(
        ValidationLayerResult(
            layer_name="Layer 2 — Mathematical Verification (Cấu trúc OMML)",
            status=l2_status,
            summary=f"Phát hiện {len(omath_nodes)} khối công thức OMML (Phân số: {len(fractions)}, Căn: {len(radicals)}, Tích phân/Tổng: {len(nary_ops)}, Ma trận: {len(matrices)}).",
            details={
                "omml_equations": len(omath_nodes),
                "fractions_m_f": len(fractions),
                "radicals_m_rad": len(radicals),
                "superscripts_m_sSup": len(superscripts),
                "subscripts_m_sSub": len(subscripts),
                "nary_integrals_sums": len(nary_ops),
                "matrices_m_m": len(matrices),
            },
        )
    )

    # --- LAYER 3: ILLUSTRATION VERIFICATION ---
    tikz_count = len(re.findall(r"\\begin\{tikzpicture\}", tex_content))
    standalone_axis = len(re.findall(r"\\begin\{axis\}", tex_content))
    includegraphics_count = len(re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}", tex_content))
    expected_min_graphics = tikz_count + includegraphics_count

    l3_status = "PASS" if len(media_files) >= expected_min_graphics else "WARN"
    results.append(
        ValidationLayerResult(
            layer_name="Layer 3 — Illustration Verification (Hình vẽ TikZ/PGFPlots)",
            status=l3_status,
            summary=f"Nhúng thành công {len(media_files)}/{expected_min_graphics} hình ảnh vào Word (Nguồn có {tikz_count} TikZ, {standalone_axis} trục PGFPlots, {includegraphics_count} ảnh tĩnh).",
            details={
                "embedded_media_files": media_files,
                "source_tikzpicture_count": tikz_count,
                "source_axis_count": standalone_axis,
                "source_includegraphics_count": includegraphics_count,
            },
        )
    )

    # --- LAYER 4: TABLE VERIFICATION ---
    src_tables = len(re.findall(r"\\begin\{(?:tabular\*?|tabularx|longtable)\}", tex_content))
    docx_tables = root.findall(".//w:tbl", NS)
    table_structures = []
    for idx, tbl in enumerate(docx_tables, start=1):
        rows = tbl.findall("./w:tr", NS)
        col_counts = [len(r.findall("./w:tc", NS)) for r in rows]
        merged_spans = len(tbl.findall(".//w:gridSpan", NS)) + len(tbl.findall(".//w:vMerge", NS))
        table_structures.append(
            {
                "table_index": idx,
                "rows": len(rows),
                "max_cols": max(col_counts) if col_counts else 0,
                "merged_tags": merged_spans,
            }
        )

    l4_status = "PASS" if len(docx_tables) >= src_tables else "WARN"
    results.append(
        ValidationLayerResult(
            layer_name="Layer 4 — Table Verification (Cấu trúc Bảng & Ô gộp)",
            status=l4_status,
            summary=f"Đã chuyển đổi {len(docx_tables)}/{src_tables} bảng sang cấu trúc bảng Word.",
            details={
                "source_tables": src_tables,
                "docx_tables": len(docx_tables),
                "table_structures": table_structures,
                "starter_report_excerpt": starter_report,
            },
        )
    )

    # --- LAYER 5: VISUAL VERIFICATION (PDF DIFF) ---
    if visual_diff_info and visual_diff_info.get("enabled"):
        similarity = visual_diff_info.get("similarity_percent", 0.0)
        l5_status = "PASS" if similarity >= 85.0 else "WARN"
        results.append(
            ValidationLayerResult(
                layer_name="Layer 5 — Visual Verification (Đối chiếu Thị giác PDF)",
                status=l5_status,
                summary=f"Độ tương đồng bố cục trang tài liệu đạt {similarity:.1f}% giữa PDF LaTeX gốc và PDF xuất từ Word.",
                details=visual_diff_info,
            )
        )
    else:
        results.append(
            ValidationLayerResult(
                layer_name="Layer 5 — Visual Verification (Đối chiếu Thị giác PDF)",
                status="SKIP",
                summary="Bỏ qua đối chiếu PDF (Chưa bật --qa-pdf hoặc máy chưa cài đặt LibreOffice).",
                details=visual_diff_info or {},
            )
        )

    return results


# ==============================================================================
# GIAO DIỆN WEB CHÍNH (STREAMLIT APPLICATION)
# ==============================================================================
def main() -> None:
    st.set_page_config(
        page_title="LaTeX to Word Hybrid Converter",
        page_icon="📄",
        layout="wide",
    )

    st.title("📄 Hệ Thống Chuyển Đổi LaTeX sang Word (.docx) Độ Chính Xác Cao")
    st.caption(
        "Tích hợp **Cloud TikZ Compiler (QuickLaTeX & TeXLive.net)** + **Direct DOCX Image Injector** — "
        "Tự động vẽ và nhúng toàn bộ hình TikZ/PGFPlots ngay cả khi máy chưa cài MiKTeX / XeLaTeX!"
    )

    has_soffice = bool(shutil.which("soffice") or shutil.which("libreoffice"))
    has_local_latex = bool(shutil.which("xelatex") or shutil.which("pdflatex"))

    with st.sidebar:
        st.header("⚙️ Cấu hình Chuyển đổi")
        conversion_mode = st.radio(
            "Chế độ ưu tiên (Mục 7):",
            options=[
                "Mode 1: Maximum Visual Accuracy (Giữ nguyên giao diện)",
                "Mode 2: Maximum Editability (Tối ưu chỉnh sửa Word)",
            ],
            index=1,
        )

        dpi = st.select_slider(
            "Độ phân giải đồ họa TikZ/PGFPlots (DPI):",
            options=[200, 300, 400, 600],
            value=400,
        )
        strict_mode = st.checkbox(
            "Bật kiểm tra nghiêm ngặt (--strict)",
            value=False,
        )
        enable_qa_pdf = st.checkbox(
            "Bật đối chiếu thị giác PDF (--qa-pdf / Layer 5)",
            value=False,
        )

        st.divider()
        st.subheader("🔍 Kiểm tra Môi trường Hệ thống")
        tools_status = {
            "latex2word.py": CONVERTER_SCRIPT.exists(),
            "pandoc": shutil.which("pandoc") is not None,
            "TikZ Engine (Local hoặc Cloud API)": True,
            "Local xelatex / pdflatex (Tùy chọn)": has_local_latex,
            "python-docx (Nhúng ảnh & Bảng OOXML)": HAS_DOCX,
            "Pillow + NumPy (Xử lý ảnh)": HAS_VISION,
        }
        for tool_name, ok in tools_status.items():
            icon = "✅" if ok else "ℹ️"
            st.write(f"{icon} **{tool_name}**")

    tab_paste, tab_upload = st.tabs(
        ["✍️ Cách 1: Dán trực tiếp mã LaTeX / TikZ", "📂 Cách 2: Tải lên file .tex hoặc gói .zip"]
    )

    initial_code = DEFAULT_LATEX_TEMPLATE
    if EXAMPLE_TEX.exists():
        try:
            initial_code = EXAMPLE_TEX.read_text(encoding="utf-8")
        except Exception:
            pass

    with tab_paste:
        tex_input_code = st.text_area(
            "Dán toàn bộ tài liệu LaTeX hoặc đề thi có chứa `\\begin{tikzpicture} ... \\end{tikzpicture}`:",
            value=initial_code,
            height=360,
        )
        convert_from_paste = st.button("🚀 Chuyển đổi mã LaTeX sang Word", type="primary", use_container_width=True)

    with tab_upload:
        uploaded_file = st.file_uploader(
            "Chọn file `.tex` đơn lẻ hoặc file nén `.zip` (nếu có kèm thư mục ảnh và các file \\input):",
            type=["tex", "zip"],
        )
        main_tex_name = st.text_input(
            "Tên file .tex chính (chỉ cần nhập nếu tải lên gói .zip có nhiều file .tex):",
            value="example.tex",
        )
        convert_from_upload = st.button(
            "🚀 Chuyển đổi File tải lên sang Word",
            type="primary",
            use_container_width=True,
            disabled=(uploaded_file is None),
        )

    if convert_from_paste or convert_from_upload:
        with tempfile.TemporaryDirectory(prefix="latex2word_web_") as tmp_str:
            workdir = Path(tmp_str)
            target_tex = workdir / "document.tex"

            if convert_from_paste:
                target_tex.write_text(tex_input_code, encoding="utf-8")
            else:
                assert uploaded_file is not None
                if uploaded_file.name.endswith(".zip"):
                    zip_path = workdir / uploaded_file.name
                    zip_path.write_bytes(uploaded_file.getvalue())
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        zf.extractall(workdir)
                    tex_candidates = list(workdir.rglob("*.tex"))
                    if not tex_candidates:
                        st.error("Không tìm thấy file `.tex` nào bên trong gói `.zip`.")
                        return
                    matched = [p for p in tex_candidates if p.name == main_tex_name]
                    target_tex = matched[0] if matched else tex_candidates[0]
                else:
                    target_tex = workdir / uploaded_file.name
                    target_tex.write_bytes(uploaded_file.getvalue())

            raw_tex_source = target_tex.read_text(encoding="utf-8", errors="ignore")
            output_docx = workdir / f"{target_tex.stem}.docx"
            report_json_path = output_docx.with_suffix(".report.json")
            qa_pdf_path = output_docx.with_suffix(".pdf")

            with st.status("Đang thực thi Pipeline chuyển đổi LaTeX → Word (kèm biên dịch hình TikZ)...", expanded=True) as status_box:
                # BƯỚC 1: Chuẩn hóa macro và biên dịch toàn bộ khối TikZ (Local hoặc Cloud)
                normalized_tex, wrap_notes = normalize_and_wrap_latex(raw_tex_source)
                for note in wrap_notes:
                    st.write(f"ℹ️ {note}")

                st.write("🎨 Đang biên dịch các hình vẽ TikZ / PGFPlots sang ảnh độ phân giải cao...")
                marked_tex, marker_to_png, tikz_logs, tikz_warnings = extract_and_render_all_tikz(
                    normalized_tex, target_tex.parent, dpi=dpi
                )
                target_tex.write_text(marked_tex, encoding="utf-8")
                for tlog in tikz_logs:
                    st.write(f"✅ {tlog}")

                # BƯỚC 2: Chuyển đổi văn bản, toán OMML và bảng qua `latex2word.py` hoặc `pandoc`
                cmd = [
                    sys.executable,
                    str(CONVERTER_SCRIPT),
                    str(target_tex),
                    "-o",
                    str(output_docx),
                    "--dpi",
                    str(dpi),
                ]
                if strict_mode:
                    cmd.append("--strict")
                if enable_qa_pdf and has_soffice:
                    cmd.append("--qa-pdf")

                proc = None
                if CONVERTER_SCRIPT.exists():
                    st.write("1️⃣ Đang chuyển đổi văn bản, công thức OMML và bảng qua `latex2word.py`...")
                    proc = run_cmd_utf8(cmd, cwd=target_tex.parent)

                if not output_docx.exists() and shutil.which("pandoc"):
                    st.write("🔄 Đang chạy Pandoc trực tiếp...")
                    fallback_cmd = [
                        shutil.which("pandoc") or "pandoc",
                        str(target_tex),
                        "--from=latex+raw_tex",
                        "--to=docx",
                        "--standalone",
                        f"--resource-path={target_tex.parent}",
                        "--output",
                        str(output_docx),
                    ]
                    proc = run_cmd_utf8(fallback_cmd, cwd=target_tex.parent)

                if not output_docx.exists() and HAS_DOCX:
                    pure_python_latex_to_docx(marked_tex, target_tex.parent, output_docx)

                # BƯỚC 3: Nhúng trực tiếp toàn bộ ảnh TikZ đã biên dịch vào đúng vị trí Marker trong file .docx
                inserted_imgs = 0
                if output_docx.exists() and marker_to_png:
                    st.write(f"🖼️ Đang nhúng trực tiếp {len(marker_to_png)} hình vẽ TikZ vào đúng vị trí câu hỏi trong Word...")
                    inserted_imgs = inject_tikz_images_into_docx(output_docx, marker_to_png)

                if not output_docx.exists():
                    status_box.update(label="Chuyển đổi thất bại!", state="error", expanded=True)
                else:
                    ooxml_stats = {}
                    if "Mode 2" in conversion_mode:
                        st.write("2️⃣ Đang chuẩn hóa khung viền bảng OOXML...")
                        ooxml_stats = postprocess_docx_tables(output_docx)

                    starter_report: dict[str, Any] = {}
                    if report_json_path.exists():
                        try:
                            starter_report = json.loads(report_json_path.read_text(encoding="utf-8"))
                        except Exception:
                            pass

                    st.write("3️⃣ Đang tổng hợp báo cáo 5 tầng xác thực (5-Layer Automatic Validation)...")
                    layer_results = run_five_layer_validation(
                        raw_tex_source, output_docx, starter_report, None
                    )
                    status_box.update(
                        label=f"Hoàn tất! Đã tạo Word và nhúng thành công {inserted_imgs} hình vẽ TikZ!",
                        state="complete",
                        expanded=False,
                    )

            if not output_docx.exists():
                full_log = (((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()) if proc else ""
                st.error("❌ **Quá trình chuyển đổi không tạo ra được file `.docx`.**")
                st.code(full_log or "Hãy kiểm tra cài đặt Pandoc và python-docx.", language="text")
                return

            if tikz_warnings:
                with st.expander(f"⚠️ Có {len(tikz_warnings)} hình vẽ TikZ cần lưu ý (Bấm để xem chi tiết)"):
                    for w in tikz_warnings:
                        st.warning(w)

            st.success(
                f"✅ Đã tạo thành công file `{output_docx.name}` (Đã nhúng **{inserted_imgs}** hình vẽ TikZ/PGFPlots vào Word)!"
            )
            if ooxml_stats:
                st.info(
                    f"✨ **Mode 2 OOXML Post-Processor**: Đã chuẩn hóa khung viền cho "
                    f"**{ooxml_stats.get('tables_processed', 0)}** bảng "
                    f"({ooxml_stats.get('cells_formatted', 0)} ô dữ liệu, "
                    f"{ooxml_stats.get('merged_cells_detected', 0)} ô gộp)."
                )

            # Xem trước thư viện hình TikZ vừa biên dịch ngay trên Web
            if marker_to_png:
                with st.expander(f"🖼️ Xem trước {len(marker_to_png)} hình vẽ TikZ/PGFPlots đã được biên dịch & nhúng vào Word", expanded=True):
                    cols = st.columns(min(3, len(marker_to_png)))
                    for i, (marker, p_img) in enumerate(marker_to_png.items()):
                        with cols[i % len(cols)]:
                            st.image(str(p_img), caption=f"Hình #{i+1} ({marker})", use_container_width=True)

            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label=f"📥 Tải xuống Word ({output_docx.name})",
                    data=output_docx.read_bytes(),
                    file_name=output_docx.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )
            with dl_col2:
                combined_report = {
                    "mode": conversion_mode,
                    "embedded_tikz_images": inserted_imgs,
                    "ooxml_postprocess": ooxml_stats,
                    "five_layer_validation": [
                        {
                            "layer": r.layer_name,
                            "status": r.status,
                            "summary": r.summary,
                            "details": r.details,
                        }
                        for r in layer_results
                    ],
                    "starter_engine_report": starter_report,
                }
                st.download_button(
                    label="📊 Tải báo cáo kiểm thử (.report.json)",
                    data=json.dumps(combined_report, ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name=f"{output_docx.stem}.report.json",
                    mime="application/json",
                    use_container_width=True,
                )

            st.subheader("🛡️ Báo cáo Kiểm thử Tự động 5 Tầng (5 Validation Layers)")
            for res in layer_results:
                badge = {"PASS": "✅ PASS", "WARN": "⚠️ WARN", "FAIL": "❌ FAIL", "SKIP": "⏭️ SKIP"}.get(
                    res.status, res.status
                )
                with st.expander(f"{badge} | **{res.layer_name}** — {res.summary}"):
                    st.json(res.details)


if __name__ == "__main__":
    main()