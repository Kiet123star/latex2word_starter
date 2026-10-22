#!/usr/bin/env python3
"""Giao diện Web (Streamlit) đồng bộ với Nhân Xử Lý Duy Nhất (`backend/core_converter.py`).

Hỗ trợ chuyển đổi trực tiếp cả:
- File `.pdf` (Đề thi Toán, tài liệu Toán học PDF — tích hợp Bộ giải mã 2D Pure-Python `zlib` không phụ thuộc thư viện ngoài)
- File `.tex` (Mã nguồn LaTeX / TikZ / PGFPlots / tkz-tab)
- Gói `.zip` (chứa `.pdf` hoặc `.tex`)

Chạy ứng dụng bằng lệnh:
    streamlit run app.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

# Tự động đồng bộ PATH từ Windows Registry (giúp tìm thấy LibreOffice, Poppler, MiKTeX nếu có)
if sys.platform == "win32":
    try:
        import winreg

        reg_paths: list[str] = []
        for hkey in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            sub = (
                r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
                if hkey == winreg.HKEY_LOCAL_MACHINE
                else r"Environment"
            )
            try:
                with winreg.OpenKey(hkey, sub) as key:
                    val, _ = winreg.QueryValueEx(key, "Path")
                    if val:
                        reg_paths.append(val)
            except Exception:
                pass
        if reg_paths:
            combined = ";".join(reg_paths) + ";" + os.environ.get("PATH", "")
            seen: set[str] = set()
            clean_p: list[str] = []
            for pt in combined.split(";"):
                pc = pt.strip()
                if pc and pc.lower() not in seen:
                    seen.add(pc.lower())
                    clean_p.append(pc)
            os.environ["PATH"] = ";".join(clean_p)
    except Exception:
        pass

from backend.core_converter import convert_latex_file_unified, decompile_pdf_to_latex
from backend.smart_paste import convert_pasted_latex_unified  # chỉ dùng cho nhánh "Dán mã LaTeX"
from backend.validators import generate_visual_comparison, run_all_layers

BASE_DIR = Path(__file__).resolve().parent
EXAMPLE_TEX = BASE_DIR / "example.tex"

DEFAULT_LATEX_TEMPLATE = r"""\documentclass[12pt,a4paper]{article}
\usepackage{amsmath,amssymb}
\usepackage{tikz,pgfplots,tkz-tab}
\usepackage{booktabs,multirow}
\pgfplotsset{compat=1.18}

\title{Mẫu Kiểm Thử PDF \& LaTeX sang Word (Native DrawingML Shapes \& Bảng Toán)}
\author{Nhân Xử Lý Hợp Nhất (Unified Core Engine v2.4 Universal)}
\date{\today}

\begin{document}
\maketitle

\textbf{Câu 1.} Cho hàm số bậc ba $y = f(x)$ có đồ thị như hình vẽ bên. Hỏi phương trình $f(f(x)) = 0$ có bao nhiêu nghiệm thực phân biệt?
\immini{
Khảo sát hàm số hợp $g(x) = f(f(x))$ dựa trên đồ thị hàm số $y = f(x)$:
\[
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}, \quad \binom{n}{k} = \frac{n!}{k!(n-k)!}
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
        page_title="PDF & LaTeX to Word Unified Converter (Native DrawingML & Math Tables)",
        page_icon="📐",
        layout="wide",
    )

    st.title("📐 Hệ Thống Chuyển Đổi PDF & LaTeX → Word (.docx)")
    st.caption(
        "Đồng bộ chung một nhân xử lý `backend/core_converter.py` (v2.4 Universal) cho cả CLI, FastAPI và Streamlit — "
        "Hỗ trợ tải trực tiếp **file `.pdf`** (tích hợp **2D Pure-Python Math PDF Decompiler** không cần cài đặt công cụ ngoài), "
        "**`.tex`** hoặc **`.zip`**, chuyển hình học TikZ/PGFPlots thành **Word Shapes (DrawingML)** và hỗ trợ đầy đủ **Bảng Biến Thiên, Hệ phương trình/ngoặc vuông, Công thức Tổ hợp `\\binom{n}{k}`**."
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
        st.write("✅ **2D Pure-Python Math PDF Decompiler (`zlib` + `CMEX10/CMSY10/CMMI10`)**")
        st.write("✅ **Native DrawingML Shape Compiler (`<wpg:wgp>`)**")
        st.write("✅ **Native Math Tables (`tkz-tab`, `\\choice`, `tabular`, `\\immini`)**")
        st.write("✅ **Native OMML Math (`<m:oMath>`, `\\heva`, `\\hoac`, `\\xrightarrow`, `\\binom`)**")

    tab_paste, tab_upload = st.tabs(
        ["✍️ Cách 1: Dán trực tiếp mã LaTeX / TikZ", "📂 Cách 2: Tải lên file .pdf, .tex hoặc gói .zip"]
    )

    with tab_paste:
        tex_input_code = st.text_area(
            "Dán mã LaTeX có chứa hình học TikZ, đồ thị PGFPlots, bảng biến thiên `tkz-tab`, trắc nghiệm `\\choice`, tổ hợp `\\binom{n}{k}`:",
            value=DEFAULT_LATEX_TEMPLATE,
            height=380,
        )
        convert_from_paste = st.button("🚀 Chuyển đổi sang Word (.docx)", type="primary", use_container_width=True)

    with tab_upload:
        uploaded_file = st.file_uploader(
            "Chọn file `.pdf`, `.tex` đơn lẻ hoặc gói `.zip` (chứa `.pdf` / `.tex`):",
            type=["pdf", "tex", "zip"],
        )
        main_tex_name = st.text_input(
            "Tên file .tex chính (tùy chọn nếu tải lên gói .zip chứa nhiều file):",
            value="",
            placeholder="document.tex (để trống nếu tải file .pdf)",
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
                lower_name = uploaded_file.name.lower()
                if lower_name.endswith(".zip"):
                    zip_path = workdir / uploaded_file.name
                    zip_path.write_bytes(uploaded_file.getvalue())
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        zf.extractall(workdir)
                    tex_candidates = [p for p in workdir.rglob("*.tex") if "__MACOSX" not in p.parts]
                    pdf_candidates = [p for p in workdir.rglob("*.pdf") if "__MACOSX" not in p.parts]
                    if tex_candidates:
                        matched = [p for p in tex_candidates if p.name == main_tex_name] if main_tex_name else []
                        target_tex = matched[0] if matched else tex_candidates[0]
                    elif pdf_candidates:
                        pdf_file = pdf_candidates[0]
                        decompiled = decompile_pdf_to_latex(
                            pdf_file,
                            search_dirs=[workdir, BASE_DIR, BASE_DIR / "jobs"],
                        )
                        target_tex = workdir / f"{pdf_file.stem}.tex"
                        target_tex.write_text(decompiled, encoding="utf-8")
                    else:
                        st.error("Không tìm thấy file `.tex` hoặc `.pdf` nào trong gói `.zip`.")
                        return
                elif lower_name.endswith(".pdf"):
                    pdf_path = workdir / uploaded_file.name
                    pdf_path.write_bytes(uploaded_file.getvalue())
                    decompiled = decompile_pdf_to_latex(
                        pdf_path,
                        search_dirs=[workdir, BASE_DIR, BASE_DIR / "jobs"],
                    )
                    target_tex = workdir / f"{pdf_path.stem}.tex"
                    target_tex.write_text(decompiled, encoding="utf-8")
                else:
                    target_tex = workdir / uploaded_file.name
                    target_tex.write_bytes(uploaded_file.getvalue())

            output_docx = workdir / f"{target_tex.stem}.docx"
            raw_tex_source = target_tex.read_text(encoding="utf-8", errors="replace")

            with st.status("Đang thực thi Nhân Xử Lý Hợp Nhất (PDF/LaTeX → DrawingML + OMML + Math Tables)...", expanded=True) as status_box:
                # Nhánh dán mã → tiền xử lý thông minh (macro lạ, \cau, \bonpa, môi trường lạ...).
                # Nhánh tải file → giữ nguyên engine gốc, không thay đổi gì.
                convert_fn = convert_pasted_latex_unified if convert_from_paste else convert_latex_file_unified
                report = convert_fn(
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
                        f"Hoàn tất! Đã tạo {report.get('equation_nodes', 0)} công thức OMML, "
                        f"{report.get('drawingml_groups', 0)} nhóm hình DrawingML "
                        f"({report.get('drawingml_shapes', 0)} nét vẽ), "
                        f"{report.get('tkz_tab_tables', 0)} bảng biến thiên và "
                        f"{report.get('word_tables', 0)} bảng Word!"
                    ),
                    state="complete",
                    expanded=False,
                )

            smart_info = report.get("smart_paste") or {}
            if smart_info.get("status") == "ok":
                auto_cmds = smart_info.get("unknown_commands_auto_handled") or {}
                auto_envs = smart_info.get("unknown_environments_auto_handled") or {}
                with st.expander(
                    f"🧠 Tiền xử lý thông minh: {smart_info.get('stats', {}).get('macros_expanded', 0)} macro được mở rộng, "
                    f"{len(auto_cmds)} lệnh lạ + {len(auto_envs)} môi trường lạ được xử lý tự động"
                ):
                    if auto_cmds or auto_envs:
                        st.write("Lệnh/môi trường lạ (đã giữ nội dung, bỏ tên lệnh):",
                                 ", ".join(f"\\{k}" for k in auto_cmds), ", ".join(auto_envs))
                    for w in smart_info.get("warnings", [])[:20]:
                        st.write("⚠️ " + w)
                    st.json(smart_info.get("stats", {}))
            elif smart_info.get("status") == "fallback_original_engine":
                st.warning("Tiền xử lý thông minh gặp lỗi nên đã dùng engine gốc: " + str(smart_info.get("error", "")))

            # ── Lớp bảo hiểm phủ 100% (chỉ có ở nhánh dán mã) ──
            uf = smart_info.get("universal_fallback") or {}
            cov = uf.get("do_phu") or {}
            if cov:
                c1, c2, c3 = st.columns(3)
                c1.metric("Phủ nguyên bản (sửa được trong Word)",
                          f"{cov.get('do_phu_nguyen_ban_phan_tram', 0)}%")
                c2.metric("Phủ tổng (giống bản gốc)",
                          f"{cov.get('do_phu_tong_phan_tram', 0)}%")
                c3.metric("Khối do TeX dựng (ảnh vector)",
                          f"{(uf.get('ket_qua') or {}).get('so_anh_chen', 0)}/{cov.get('so_khoi_chup_anh', 0)}")
                with st.expander("🛡️ Chi tiết lớp bảo hiểm phủ 100% (dán mã)"):
                    st.write(uf.get("ly_do", ""))
                    if cov.get("lenh_la"):
                        st.write("Lệnh chưa có trong sổ đăng ký:", ", ".join(f"\\{k}" for k in cov["lenh_la"]))
                    if cov.get("moi_truong_la"):
                        st.write("Môi trường chưa có trong sổ đăng ký:", ", ".join(cov["moi_truong_la"]))
                    for w in (uf.get("canh_bao") or [])[:10]:
                        st.write("⚠️ " + str(w))
                    for w in ((uf.get("ket_qua") or {}).get("canh_bao") or [])[:10]:
                        st.write("⚠️ " + str(w))
                    st.json({"tex": uf.get("tex"), "cong_cu": uf.get("cong_cu")})
            elif uf and not uf.get("bat"):
                st.info("🛡️ Lớp bảo hiểm phủ 100% chưa chạy: " + str(uf.get("ly_do", "")))

            st.success(
                f"✅ Đã tạo thành công `{output_docx.name}` với **{report.get('equation_nodes', 0)}** công thức toán OMML chỉnh sửa trực tiếp, "
                f"**{report.get('drawingml_groups', 0)}** nhóm hình vẽ gốc Word (DrawingML), "
                f"**{report.get('tkz_tab_tables', 0)}** bảng biến thiên Word và **{report.get('word_tables', 0)}** bảng Word!"
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

            with st.expander("📄 Xem mã nguồn LaTeX đã giải mã / chuẩn hóa từ file đầu vào"):
                st.code(raw_tex_source, language="latex")

            st.subheader("🛡️ Báo cáo Kiểm thử Tự động 5 Tầng (5 Validation Layers)")
            for res in layers:
                badge = {"PASS": "✅ PASS", "WARN": "⚠️ WARN", "FAIL": "❌ FAIL", "SKIP": "⏭️ SKIP"}.get(
                    res["status"], res["status"]
                )
                with st.expander(f"{badge} | **Layer {res['layer']} — {res['name']}** — {res['summary']}"):
                    st.json(res["details"])


if __name__ == "__main__":
    main()




