#!/usr/bin/env python3
"""Giao diện Web (Streamlit) đồng bộ với Nhân Xử Lý Duy Nhất (`backend/core_converter.py`).

Chạy ứng dụng bằng lệnh:
    streamlit run app.py
"""
from __future__ import annotations

import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from backend.core_converter import convert_latex_file_unified
from backend.validators import generate_visual_comparison, run_all_layers

os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

BASE_DIR = Path(__file__).resolve().parent
EXAMPLE_TEX = BASE_DIR / "example.tex"

DEFAULT_LATEX_TEMPLATE = r"""\documentclass[12pt,a4paper]{article}
\usepackage{amsmath,amssymb}
\usepackage{tikz,pgfplots,tkz-tab}
\usepackage{booktabs,multirow}
\pgfplotsset{compat=1.18}

\title{Mẫu Kiểm Thử LaTeX sang Word (Native DrawingML Shapes \& Bảng Toán)}
\author{Nhân Xử Lý Hợp Nhất (Unified Core Engine)}
\date{\today}

\begin{document}
\maketitle

\textbf{Câu 1.} Cho hàm số bậc ba $y = f(x)$ có đồ thị như hình vẽ bên. Hỏi phương trình $f(f(x)) = 0$ có bao nhiêu nghiệm thực phân biệt?
\immini{
Khảo sát hàm số hợp $g(x) = f(f(x))$ dựa trên đồ thị hàm số $y = f(x)$:
\[
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}, \quad I = \int_{0}^{1} (x^3 - 3x)\,dx
\]
\choiceB{$3$.}{$7$.}{$9$.}{$5$.}
}{
\begin{tikzpicture}
    \begin{axis}[
        axis lines=middle,
        width=6.5cm, height=4.2cm,
        xmin=-2, xmax=2,
        ymin=-3.5, ymax=2.5,
        xlabel={$x$}, ylabel={$y$}
    ]
    \addplot[thick, blue, domain=-1.65:1.65] {x^3 - 3*x};
    \end{axis}
\end{tikzpicture}
}

\textbf{Câu 2.} Cho hàm số $f(x)$ có bảng biến thiên như hình vẽ dưới đây. Số điểm cực trị của hàm số $g(x) = 2 - \dfrac{1}{x^4} f^3(2x+1)$ là:
\begin{center}
\begin{tikzpicture}
    \tkzTabInit[lgt=1.2, espcl=1.6, deltacl=0.6]
      {$x$ / 0.8, $f'(x)$ / 0.8, $f(x)$ / 2}
      {$-\infty$, $-1$, $0$, $1$, $+\infty$}
    \tkzTabLine{ , -, 0, +, 0, -, 0, +, }
    \tkzTabVar{+/ $+\infty$, -/ $-2$, +/ $3$, -/ $-2$, +/ $+\infty$}
\end{tikzpicture}
\end{center}
\choiceC{$4$}{$6$}{$7$}{$5$}

\textbf{Câu 3.} Bảng thống kê ghép nhóm chiều cao của học sinh lớp 12A:
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
Nhóm chiều cao (cm) & $[150; 155)$ & $[155; 160)$ & $[160; 165)$ & $[165; 170)$ \\
\hline
Tần số ($n_i$) & $5$ & $12$ & $18$ & $7$ \\
\hline
\end{tabular}
\end{center}

\end{document}
"""


def main() -> None:
    st.set_page_config(
        page_title="LaTeX to Word Unified Converter (Native DrawingML & Math Tables)",
        page_icon="📐",
        layout="wide",
    )

    st.title("📐 Hệ Thống Chuyển Đổi LaTeX → Word (.docx) (Native DrawingML & Bảng Toán)")
    st.caption(
        "Đồng bộ chung một nhân xử lý `backend/core_converter.py` cho cả CLI, FastAPI và Streamlit — "
        "Chuyển hình học TikZ/PGFPlots thành **Word Shapes (DrawingML)** chỉnh sửa từng nét và hỗ trợ đầy đủ **Bảng Biến Thiên, Bảng Trắc Nghiệm, Bảng Dữ Liệu Gộp Ô**."
    )

    has_soffice = bool(shutil.which("soffice") or shutil.which("libreoffice"))

    with st.sidebar:
        st.header("⚙️ Cấu hình Chuyển đổi")
        conversion_mode = st.radio(
            "Chế độ ưu tiên:",
            options=[
                "Mode 2: Native Word Shapes (DrawingML) & Editable Math Tables",
                "Mode 1: Standard Layout Mode",
            ],
            index=0,
        )
        mode_int = 2 if "Mode 2" in conversion_mode else 1

        dpi = st.select_slider(
            "Độ phân giải (DPI):",
            options=[200, 300, 400, 600],
            value=400,
        )
        enable_qa_pdf = st.checkbox(
            "Xuất kèm PDF xem trước (--qa-pdf / Layer 5)",
            value=has_soffice,
        )

        st.divider()
        st.subheader("🔍 Trạng thái Nhân Xử Lý")
        st.write("✅ **Unified Core Engine (`backend/core_converter.py`)**")
        st.write("✅ **Native DrawingML Shape Compiler (`<wpg:wgp>`)**")
        st.write("✅ **Native Math Tables (`tkz-tab`, `\\choice`, `tabular`, `\\immini`)**")
        st.write("✅ **Native OMML Math (`<m:oMath>`)**")

    tab_paste, tab_upload = st.tabs(
        ["✍️ Cách 1: Dán trực tiếp mã LaTeX / TikZ", "📂 Cách 2: Tải lên file .tex hoặc gói .zip"]
    )

    with tab_paste:
        tex_input_code = st.text_area(
            "Dán mã LaTeX có chứa hình học TikZ, đồ thị PGFPlots, bảng biến thiên `tkz-tab`, trắc nghiệm `\\choice`:",
            value=DEFAULT_LATEX_TEMPLATE,
            height=380,
        )
        convert_from_paste = st.button("🚀 Chuyển đổi sang Word (.docx)", type="primary", use_container_width=True)

    with tab_upload:
        uploaded_file = st.file_uploader(
            "Chọn file `.tex` đơn lẻ hoặc gói `.zip`:",
            type=["tex", "zip"],
        )
        main_tex_name = st.text_input(
            "Tên file .tex chính (nếu tải lên gói .zip):",
            value="document.tex",
        )
        convert_from_upload = st.button(
            "🚀 Chuyển đổi File tải lên sang Word",
            type="primary",
            use_container_width=True,
            disabled=(uploaded_file is None),
        )

    if convert_from_paste or convert_from_upload:
        with tempfile.TemporaryDirectory(prefix="latex2word_unified_") as tmp_str:
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
                        st.error("Không tìm thấy file `.tex` trong gói `.zip`.")
                        return
                    matched = [p for p in tex_candidates if p.name == main_tex_name]
                    target_tex = matched[0] if matched else tex_candidates[0]
                else:
                    target_tex = workdir / uploaded_file.name
                    target_tex.write_bytes(uploaded_file.getvalue())

            output_docx = workdir / f"{target_tex.stem}.docx"
            raw_tex_source = target_tex.read_text(encoding="utf-8", errors="replace")

            with st.status("Đang thực thi Nhân Xử Lý Hợp Nhất (DrawingML + OMML + Math Tables)...", expanded=True) as status_box:
                report = convert_latex_file_unified(
                    target_tex,
                    output_docx,
                    dpi=dpi,
                    qa_pdf=enable_qa_pdf,
                    mode=mode_int,
                )
                visual_info = {"enabled": False}
                if enable_qa_pdf and report.get("qa_rendered_pdf"):
                    visual_info, _ = generate_visual_comparison(
                        target_tex.parent, target_tex, Path(report["qa_rendered_pdf"])
                    )
                layers = run_all_layers(raw_tex_source, output_docx, visual_info)
                status_box.update(
                    label=(
                        f"Hoàn tất! Đã tạo {report.get('drawingml_groups', 0)} nhóm hình DrawingML "
                        f"({report.get('drawingml_shapes', 0)} nét vẽ), "
                        f"{report.get('tkz_tab_tables', 0)} bảng biến thiên và "
                        f"{report.get('word_tables', 0)} bảng Word!"
                    ),
                    state="complete",
                    expanded=False,
                )

            st.success(
                f"✅ Đã tạo thành công `{output_docx.name}` với **{report.get('drawingml_groups', 0)}** nhóm hình vẽ gốc Word (DrawingML), "
                f"**{report.get('tkz_tab_tables', 0)}** bảng biến thiên Word và **{report.get('equation_nodes', 0)}** công thức OMML!"
            )

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
                st.download_button(
                    label="📊 Tải báo cáo kiểm thử (.report.json)",
                    data=json.dumps({"engine_report": report, "layers": layers}, ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name=f"{output_docx.stem}.report.json",
                    mime="application/json",
                    use_container_width=True,
                )

            st.subheader("🛡️ Báo cáo Kiểm thử Tự động 5 Tầng (5 Validation Layers)")
            for res in layers:
                badge = {"PASS": "✅ PASS", "WARN": "⚠️ WARN", "FAIL": "❌ FAIL", "SKIP": "⏭️ SKIP"}.get(
                    res["status"], res["status"]
                )
                with st.expander(f"{badge} | **Layer {res['layer']} — {res['name']}** — {res['summary']}"):
                    st.json(res["details"])


if __name__ == "__main__":
    main()
