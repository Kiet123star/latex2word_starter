"""`backend/tex_compile.py` — chạy TeX AN TOÀN cho bước chụp hình (snapshot) và PDF tham chiếu QA.

* luôn `-no-shell-escape` (và biến môi trường `shell_escape=f`), `-interaction=nonstopmode -halt-on-error`;
* kpathsea chế độ `paranoid` (`openin_any=p`, `openout_any=p`): không đọc/ghi đường dẫn tuyệt đối, tệp ẩn,
  hay `..` ra ngoài thư mục làm việc của job;
* thư mục làm việc = thư mục job; đầu ra trong thư mục build riêng; timeout cho mỗi lần chạy;
* không tải hay cài gói nào (MiKTeX "install on the fly" bị tắt bằng biến môi trường tương ứng).

Giới hạn thực tế: đây KHÔNG phải sandbox hệ điều hành; muốn cô lập hoàn toàn cần container/tài khoản riêng.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any

__all__ = ["choose_engine", "run_tex", "compile_reference_pdf", "tex_available"]

_FONTSPEC_RE = re.compile(r"\\usepackage(?:\[[^\]]*\])?\{[^}]*(?:fontspec|unicode-math|xeCJK|polyglossia)[^}]*\}")
_PDFLATEX_ONLY_RE = re.compile(r"\\usepackage(?:\[[^\]]*\])?\{[^}]*(?:inputenc|vietnam|vntex|fontenc)[^}]*\}")


def tex_available() -> dict[str, str | None]:
    return {e: shutil.which(e) for e in ("xelatex", "pdflatex", "lualatex")}


def choose_engine(preamble: str, preferred: str = "xelatex") -> str | None:
    """fontspec/unicode-math → XeLaTeX/LuaLaTeX; inputenc/vietnam/vntex (pdfTeX) → pdfLaTeX; còn lại theo lựa chọn."""
    av = tex_available()
    if _FONTSPEC_RE.search(preamble or ""):
        order = ["xelatex", "lualatex"]
    elif _PDFLATEX_ONLY_RE.search(preamble or ""):
        order = ["pdflatex", "xelatex", "lualatex"]
    else:
        order = [preferred, "xelatex", "pdflatex", "lualatex"]
    for e in order:
        if av.get(e):
            return e
    return None


def _safe_env() -> dict[str, str]:
    env = dict(os.environ)
    env.update({
        "openin_any": "p", "openout_any": "p", "shell_escape": "f", "shell_escape_commands": "",
        "max_print_line": "10000", "MIKTEX_ENABLEINSTALLER": "f", "MIKTEX_AUTOINSTALL": "0",
        "TEXMF_OUTPUT_DIRECTORY": "",
    })
    return env


_ERR_RE = re.compile(r"^(?:! |.*?:\d+: )(.*)$", re.M)


def run_tex(tex_path: Path, out_dir: Path, *, engine: str, timeout: int = 120, runs: int = 1) -> dict[str, Any]:
    """Biên dịch `tex_path` (cwd = thư mục chứa nó) ra `out_dir`. Trả về {ok, pdf, errors, seconds, engine}."""
    exe = shutil.which(engine)
    if not exe:
        return {"ok": False, "error": f"không có {engine}", "engine": engine, "errors": []}
    out_dir.mkdir(parents=True, exist_ok=True)
    # \include cần thư mục con tương ứng trong out_dir để ghi .aux
    for d in tex_path.parent.rglob("*"):
        if d.is_dir() and out_dir not in (d, *d.parents) and ".l2w" not in d.name:
            try:
                (out_dir / d.relative_to(tex_path.parent)).mkdir(parents=True, exist_ok=True)
            except (ValueError, OSError):
                pass
    cmd = [exe, "-no-shell-escape", "-interaction=nonstopmode", "-halt-on-error", "-file-line-error",
           f"-output-directory={out_dir}", tex_path.name]
    t0 = time.time()
    log = ""
    rc = None
    for _ in range(max(1, runs)):
        try:
            cp = subprocess.run(cmd, cwd=str(tex_path.parent), env=_safe_env(), capture_output=True,
                                timeout=timeout, stdin=subprocess.DEVNULL)
            rc = cp.returncode
            log = (cp.stdout or b"").decode("utf-8", "replace")[-60000:]
        except subprocess.TimeoutExpired:
            return {"ok": False, "error": f"quá thời gian {timeout}s", "engine": engine, "errors": [],
                    "seconds": round(time.time() - t0, 1)}
        if rc != 0:
            break
        if "Rerun to get" not in log and "Label(s) may have changed" not in log:
            break
    pdf = out_dir / (tex_path.stem + ".pdf")
    errors = [m.group(1).strip()[:300] for m in _ERR_RE.finditer(log)][:20]
    ok = rc == 0 and pdf.exists()
    return {"ok": ok, "pdf": str(pdf) if pdf.exists() else None, "returncode": rc, "errors": errors,
            "engine": engine, "seconds": round(time.time() - t0, 1), "log": log,
            "error": None if ok else (errors[0] if errors else f"mã thoát {rc}")}


def compile_reference_pdf(tex_file: Path, build_dir: Path, *, engine: str | None = None,
                          timeout: int = 180) -> dict[str, Any]:
    """PDF tham chiếu cho QA = biên dịch CHÍNH nguồn gốc (không phải LaTeX suy ra từ PDF)."""
    try:
        text = tex_file.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = tex_file.read_bytes().decode("latin-1")
    if not re.search(r"\\documentclass", text):
        return {"ok": False, "error": "nguồn không có \\documentclass — không biên dịch độc lập được"}
    pre = text.split("\\begin{document}", 1)[0]
    eng = engine or choose_engine(pre)
    if not eng:
        return {"ok": False, "error": "máy không có TeX (xelatex/pdflatex/lualatex)"}
    res = run_tex(tex_file, build_dir, engine=eng, timeout=timeout, runs=2)
    res["source"] = str(tex_file)
    return res
