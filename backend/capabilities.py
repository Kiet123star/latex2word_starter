"""`backend/capabilities.py` — dò năng lực THẬT của máy đang chạy.

Thay cho các hằng `True` cứng trong `_check_tools()` / giao diện cũ.  Mỗi mục trả về:

    {"available": bool, "detail": str, "version": str | None}

* thư viện Python: thử `import` thật (không đoán theo tên gói);
* module của dự án: thử import thật, nếu lỗi ghi rõ thiếu tệp nào / lỗi gì;
* công cụ hệ thống: `shutil.which` + gọi `--version` có timeout ngắn (không cài gì cả).

Kết quả được nhớ đệm trong tiến trình; gọi `probe_capabilities(refresh=True)` để dò lại.
"""
from __future__ import annotations

import importlib
import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any

__all__ = ["probe_capabilities", "legacy_tools_view", "capability_summary", "tesseract_languages", "find_tool"]

#: Vị trí cài mặc định (Windows/macOS thường KHÔNG thêm LibreOffice/Tesseract vào PATH)
_STD_LOCATIONS: dict[str, list[str]] = {
    "soffice": [r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice"],
    "tesseract": [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                  "/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract"],
}


def find_tool(name: str, *alternatives: str) -> str | None:
    """Đường dẫn chương trình: biến môi trường L2W_<TÊN> (vd. L2W_SOFFICE) → PATH → vị trí cài mặc định.
    L2W_PATH_ONLY=1: chỉ tìm trong PATH (dùng khi kiểm thử "thiếu công cụ")."""
    env = os.environ.get("L2W_" + name.upper())
    if env and Path(env).is_file():
        return env
    for n in (name, *alternatives):
        p = shutil.which(n)
        if p:
            return p
    if os.environ.get("L2W_PATH_ONLY") == "1":
        return None
    for cand in _STD_LOCATIONS.get(name, []):
        if Path(cand).is_file():
            return cand
    return None

_LOCK = threading.Lock()
_CACHE: dict[str, Any] | None = None


def _try_import(name: str, attr: str | None = None) -> dict[str, Any]:
    try:
        mod = importlib.import_module(name)
    except Exception as exc:  # ImportError, SyntaxError, lỗi khởi tạo module…
        return {"available": False, "detail": f"{type(exc).__name__}: {exc}", "version": None}
    if attr and not hasattr(mod, attr):
        return {"available": False, "detail": f"module '{name}' không có '{attr}'", "version": None}
    ver = getattr(mod, "__version__", None) or getattr(mod, "VERSION", None)
    return {"available": True, "detail": "ok", "version": str(ver) if ver is not None else None}


def _tool(cmd: str, *version_args: str, alternatives: tuple[str, ...] = ()) -> dict[str, Any]:
    path = find_tool(cmd, *alternatives)
    if not path:
        return {"available": False, "detail": f"không tìm thấy '{cmd}' (PATH / vị trí cài mặc định / L2W_{cmd.upper()})",
                "version": None}
    ver = None
    if version_args:
        try:
            cp = subprocess.run([path, *version_args], capture_output=True, text=True, timeout=20,
                                encoding="utf-8", errors="replace")
            first = (cp.stdout or cp.stderr or "").strip().splitlines()
            ver = first[0][:160] if first else None
        except Exception as exc:  # công cụ có nhưng không chạy được
            return {"available": False, "detail": f"{path}: không chạy được ({type(exc).__name__})", "version": None}
    return {"available": True, "detail": path, "version": ver}


def tesseract_languages() -> list[str]:
    path = find_tool("tesseract")
    if not path:
        return []
    try:
        cp = subprocess.run([path, "--list-langs"], capture_output=True, text=True, timeout=20,
                            encoding="utf-8", errors="replace")
        lines = [l.strip() for l in (cp.stdout + "\n" + cp.stderr).splitlines()]
        return sorted({l for l in lines if l and " " not in l and not l.lower().startswith("list")})
    except Exception:
        return []


def _probe() -> dict[str, Any]:
    caps: dict[str, Any] = {"python": sys.version.split()[0], "platform": sys.platform}
    libs = {
        "python_docx": _try_import("docx"),
        "lxml": _try_import("lxml.etree"),
        "pillow": _try_import("PIL.Image"),
        "numpy": _try_import("numpy"),
        "pdfplumber": _try_import("pdfplumber"),
        "pypdfium2": _try_import("pypdfium2"),
        "pikepdf": _try_import("pikepdf"),
        "pytesseract": _try_import("pytesseract"),
        "fastapi": _try_import("fastapi"),
        "streamlit": _try_import("streamlit"),
    }
    caps["libraries"] = libs
    modules = {
        "core_converter": _try_import("backend.core_converter", "convert_latex_file_unified"),
        "drawingml_engine": _try_import("backend.drawingml_engine", "convert_tikz_or_axis_to_drawingml"),
        "math_tables": _try_import("backend.math_tables", "build_multiple_choice_table"),
        "omml_math": _try_import("backend.omml_math", "latex_to_omml_element"),
        "latex_preprocessor": _try_import("backend.latex_preprocessor", "preprocess_latex"),
        "tex_truth": _try_import("backend.tex_truth", "TexTruth"),
        "smart_paste": _try_import("backend.smart_paste", "convert_pasted_latex_unified"),
        "universal_fallback": _try_import("backend.universal_fallback", "protect"),
        "macro_fidelity": _try_import("backend.macro_fidelity", "rewrite_user_macros"),
        "pdf_reader": _try_import("backend.pdf_reader", "analyze_pdf"),
    }
    legacy = {"available": False, "detail": "core_converter không nạp được", "version": None}
    if modules["core_converter"]["available"]:
        cc = importlib.import_module("backend.core_converter")
        legacy = ({"available": True, "detail": "core_converter.decompile_pdf_to_latex có mặt (chỉ dùng khi bật rõ)",
                   "version": None} if hasattr(cc, "decompile_pdf_to_latex")
                  else {"available": False, "detail": "core_converter không có decompile_pdf_to_latex",
                        "version": None})
    modules["legacy_pdf_decompiler"] = legacy
    caps["modules"] = modules
    tools = {
        "xelatex": _tool("xelatex", "--version"),
        "pdflatex": _tool("pdflatex", "--version"),
        "lualatex": _tool("lualatex", "--version"),
        "pdftocairo": _tool("pdftocairo", "-v"),
        "soffice": _tool("soffice", "--version", alternatives=("libreoffice",)),
        "tesseract": _tool("tesseract", "--version"),
        "ghostscript": _tool("gs", "--version", alternatives=("gswin64c", "gswin32c")),
        "imagemagick": _tool("magick", "-version") if shutil.which("magick") else _tool("convert", "-version"),
        "pandoc": _tool("pandoc", "--version"),
    }
    if sys.platform == "win32":
        tools["docx2pdf"] = _try_import("docx2pdf")
    caps["tools"] = tools
    langs = tesseract_languages() if tools["tesseract"]["available"] else []
    caps["tesseract_languages"] = langs

    def ok(*flags: bool) -> bool:
        return all(flags)

    L, M, T = libs, modules, tools
    word_render = T["soffice"]["available"] or bool(tools.get("docx2pdf", {}).get("available"))
    derived = {
        "tex_upload": {
            "available": ok(L["python_docx"]["available"], M["core_converter"]["available"]),
            "detail": "engine dựng Word từ .tex/.zip" if M["core_converter"]["available"]
            else "thiếu backend/core_converter.py: " + M["core_converter"]["detail"],
        },
        "paste_smart": {
            "available": M["smart_paste"]["available"],
            "detail": "tiền xử lý thông minh (dán mã)" if M["smart_paste"]["available"]
            else "không nạp được smart_paste: " + M["smart_paste"]["detail"],
        },
        "pdf_reader": {
            "available": ok(L["pdfplumber"]["available"], L["pypdfium2"]["available"], L["pikepdf"]["available"],
                            M["pdf_reader"]["available"]),
            "detail": "pdfplumber (chữ/đường nét) + pypdfium2 (render/crop) + pikepdf (cấu trúc)",
        },
        "pdf_vector_crop": {
            "available": ok(T["pdftocairo"]["available"], L["pikepdf"]["available"]),
            "detail": "cắt vùng PDF thành SVG bằng pdftocairo" if T["pdftocairo"]["available"]
            else "thiếu pdftocairo → hình cắt từ PDF chỉ có PNG",
        },
        "ocr": {
            "available": ok(T["tesseract"]["available"], L["pytesseract"]["available"]),
            "detail": ("tesseract, ngôn ngữ: " + ", ".join(langs)) if T["tesseract"]["available"]
            else "thiếu tesseract → trang scan giữ ảnh",
            "vietnamese": "vie" in langs,
        },
        "math_ocr": {"available": False,
                     "detail": "chưa cấu hình bộ nhận dạng công thức từ ảnh; công thức không chắc → ảnh vùng"},
        "visual_qa": {
            "available": ok(word_render, L["pypdfium2"]["available"]),
            "detail": "render DOCX→PDF bằng LibreOffice/Word rồi so với nguồn" if word_render
            else "thiếu LibreOffice/Word → không đo được hiển thị (NOT_MEASURED)",
        },
        "tex_reference_compile": {
            "available": T["xelatex"]["available"] or T["pdflatex"]["available"] or T["lualatex"]["available"],
            "detail": "biên dịch PDF tham chiếu từ nguồn TeX (shell-escape tắt)",
        },
    }
    caps["derived"] = derived
    return caps


def probe_capabilities(refresh: bool = False) -> dict[str, Any]:
    global _CACHE
    with _LOCK:
        if _CACHE is None or refresh:
            _CACHE = _probe()
        return _CACHE


def legacy_tools_view(caps: dict[str, Any] | None = None) -> dict[str, bool]:
    """Các khóa cũ của `/api/status` (frontend cũ đọc) — nay lấy từ dò thật."""
    caps = caps or probe_capabilities()
    m, t, d = caps["modules"], caps["tools"], caps["derived"]
    return {
        "latex2word_py": d["tex_upload"]["available"],
        "drawingml_engine": m["drawingml_engine"]["available"],
        "math_tables_engine": m["math_tables"]["available"],
        "omml_math_engine": m["omml_math"]["available"],
        "pdf_decompiler": d["pdf_reader"]["available"],
        "pdf_reader": d["pdf_reader"]["available"],
        "paste_smart": d["paste_smart"]["available"],
        "ocr": d["ocr"]["available"],
        "pandoc": t["pandoc"]["available"],
        "xelatex": t["xelatex"]["available"],
        "pdflatex": t["pdflatex"]["available"],
        "pdftocairo": t["pdftocairo"]["available"],
        "libreoffice": d["visual_qa"]["available"],
        "word_installed": bool(t.get("docx2pdf", {}).get("available")),
    }


def capability_summary(caps: dict[str, Any] | None = None) -> dict[str, Any]:
    """Bản rút gọn để ghi vào báo cáo job."""
    caps = caps or probe_capabilities()
    return {
        "derived": {k: {"available": v["available"], "detail": v["detail"]} for k, v in caps["derived"].items()},
        "tools": {k: {"available": v["available"], "version": v.get("version")} for k, v in caps["tools"].items()},
        "tesseract_languages": caps.get("tesseract_languages", []),
        "missing_modules": {k: v["detail"] for k, v in caps["modules"].items() if not v["available"]},
    }
