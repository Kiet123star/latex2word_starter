"""Kiểm tra nhanh lớp tiền xử lý thông minh (chạy: python kiem_tra_smart_paste.py).

In ra: phiên bản Python/python-docx, file backend nào đang được nạp, app.py/server.py đã gắn hook chưa,
rồi chuyển đổi thử đúng đoạn \\cau / \\bonpa bằng ENGINE THẬT của bạn → tạo file kiem_tra_ket_qua.docx.
Nếu còn lỗi, chụp toàn bộ chữ in ra màn hình gửi lại.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import traceback
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

SAMPLE = r"""\documentclass[12pt,a4paper]{article}
\usepackage{amsmath,amssymb,tikz,tkz-tab}
\begin{document}
\section*{DẠNG 1: CỰC TRỊ CỦA HÀM SỐ $y=|f(x)|$}
\cau1 Cho hàm số $f(x)$ có bảng biến thiên như sau:
\begin{center}
\begin{tikzpicture}
\tkzTabInit[nocadre=false,lgt=1,espcl=2]{$x$ /1,$y'$ /1,$y$ /2}{$-\infty$,$-1$,$1$,$+\infty$}
\tkzTabLine{,+,$0$,-,$0$,+,}
\tkzTabVar{-/$-\infty$,+/$1$,-/$-3$,+/$+\infty$}
\end{tikzpicture}
\end{center}
Hàm số $y=|f(x)|$ có bao nhiêu điểm cực trị?
\bonpa{2.}{3.}{5.}{4.}

\cau2 Số điểm cực trị của hàm số $y=|f(|x|)|$ là
\bonpa{9.}{5.}{7.}{11.}
\end{document}
"""


def main() -> int:
    print("Python      :", sys.version.split()[0], "|", sys.executable)
    try:
        import docx
        print("python-docx :", getattr(docx, "__version__", "?"))
    except Exception as exc:
        print("python-docx : LỖI", exc)
        return 1
    for name in ("app.py", "backend/server.py"):
        f = ROOT / name
        ok = f.exists() and "convert_pasted_latex_unified" in f.read_text(encoding="utf-8", errors="ignore")
        print(f"Hook trong {name:18s}:", "CÓ" if ok else "CHƯA CÓ  ← hãy chép đè file này từ gói zip")
    for name in ("latex_preprocessor.py", "latex_preprocessor_parts.py", "smart_paste.py", "core_converter.py"):
        f = ROOT / "backend" / name
        print(f"backend/{name:28s}:", "có" if f.exists() else "THIẾU", f"({f.stat().st_size // 1024} KB)" if f.exists() else "")
    try:
        import backend.core_converter as cc
        from backend.smart_paste import _docx_plain_text, convert_pasted_latex_unified
        print("Đang nạp backend từ:", Path(cc.__file__).parent)
    except Exception:
        print("KHÔNG import được backend:")
        traceback.print_exc()
        return 1
    tmp = Path(tempfile.mkdtemp(prefix="l2w_check_"))
    try:
        tex = tmp / "document.tex"
        tex.write_text(SAMPLE, encoding="utf-8")
        out = tmp / "document.docx"
        rep = convert_pasted_latex_unified(tex, out, qa_pdf=False)
        sp = rep.get("smart_paste", {})
        text = _docx_plain_text(out)
        raw = [c for c in ("\\cau", "\\bonpa") if c in text]
        print("\nTrạng thái tiền xử lý :", sp.get("status"), "|", rep.get("smart_paste_status"))
        print("Bảng trắc nghiệm      :", rep.get("choice_tables"))
        print("Còn lệnh thô trong Word:", raw or "không")
        if sp.get("status") == "fallback_original_engine":
            print("\nLỖI chi tiết:\n", sp.get("trace"))
        dst = ROOT / "kiem_tra_ket_qua.docx"
        shutil.copy(out, dst)
        print("\nĐã tạo:", dst)
        print("KẾT LUẬN:", "ĐẠT ✓" if (sp.get("status") == "ok" and not raw) else "CHƯA ĐẠT ✗ — gửi toàn bộ nội dung trên cho Claude")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
