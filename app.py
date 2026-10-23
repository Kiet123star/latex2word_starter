#!/usr/bin/env python3
"""Giao diện Streamlit — dùng CHUNG pipeline (`backend/pipeline.py`) với CLI và FastAPI.

Hỗ trợ: dán mã LaTeX (đi qua cùng pipeline; khi máy có tiền xử lý thông minh thì pipeline gọi
`convert_pasted_latex_unified` cho cả nhánh dán mã lẫn tệp .tex/.zip), tải `.pdf`, `.tex`, `.zip`.

Chạy:  streamlit run app.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
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
            sub = (r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
                   if hkey == winreg.HKEY_LOCAL_MACHINE else r"Environment")
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

from backend.capabilities import probe_capabilities
from backend.pipeline import ConversionError, convert_bytes, convert_paste

BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = BASE_DIR / "jobs"
EXAMPLE_TEX = BASE_DIR / "example.tex"

DEFAULT_LATEX_TEMPLATE = r"""\documentclass[12pt,a4paper]{article}
\usepackage{amsmath,amssymb}
\usepackage{tikz,pgfplots,tkz-tab}
\pgfplotsset{compat=1.18}
\begin{document}

\textbf{Câu 1.} Cho hàm số bậc ba $y = f(x)$ có đồ thị như hình vẽ bên. Hỏi phương trình $f(f(x)) = 0$ có bao nhiêu nghiệm thực phân biệt?
\immini{
Khảo sát hàm số hợp $g(x) = f(f(x))$:
\[
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}, \quad \binom{n}{k} = \frac{n!}{k!(n-k)!}
\]
\choiceB{$3$.}{$7$.}{$9$.}{$5$.}
}{
\begin{tikzpicture}
    \begin{axis}[axis lines=middle, width=6.5cm, height=4.2cm, xmin=-2, xmax=2, ymin=-3.5, ymax=2.5,
                 xlabel={$x$}, ylabel={$y$}]
    \addplot[thick, blue, domain=-1.65:1.65] {x^3 - 3*x};
    \end{axis}
\end{tikzpicture}
}

\textbf{Câu 2.} Cho hàm số $f(x)$ có bảng biến thiên:
\begin{center}
\begin{tikzpicture}
    \tkzTabInit[lgt=1.2, espcl=1.6]{$x$ / 0.8, $f'(x)$ / 0.8, $f(x)$ / 2}{$-\infty$, $-1$, $0$, $1$, $+\infty$}
    \tkzTabLine{ , -, 0, +, 0, -, 0, +, }
    \tkzTabVar{+/ $+\infty$, -/ $-2$, +/ $3$, -/ $-2$, +/ $+\infty$}
\end{tikzpicture}
\end{center}
\choiceC{$4$}{$6$}{$7$}{$5$}

\end{document}
"""

STATUS_BADGE = {"passed": "✅ passed", "review_required": "⚠️ review_required", "failed": "❌ failed"}
LAYER_BADGE = {"PASS": "✅ PASS", "WARN": "⚠️ WARN", "FAIL": "❌ FAIL", "SKIP": "⏭️ SKIP",
               "NOT_MEASURED": "⏸️ NOT_MEASURED"}


def _capability_panel() -> None:
    caps = probe_capabilities()
    st.subheader("🔍 Năng lực trên máy này (dò thật)")
    labels = {
        "tex_upload": "Engine dựng Word từ TeX", "paste_smart": "Tiền xử lý thông minh",
        "pdf_reader": "Bộ đọc PDF", "pdf_vector_crop": "Cắt vùng PDF dạng vector (SVG)",
        "ocr": "OCR trang scan", "math_ocr": "Nhận dạng công thức từ ảnh",
        "visual_qa": "Render QA (DOCX→PDF)", "tex_reference_compile": "Biên dịch PDF tham chiếu từ TeX",
    }
    for key, label in labels.items():
        item = caps["derived"][key]
        st.write(("✅ " if item["available"] else "⚠️ ") + f"**{label}** — {item['detail']}")


def _show_report(report: dict, filename: str) -> None:
    q = report.get("quality") or {}
    status = report.get("quality_status")
    done = report.get("conversion_completed")
    if not done:
        st.error("❌ Không tạo được gói Word hợp lệ.")
    elif status == "passed":
        st.success(f"Đã tạo tệp. Kiểm định: {STATUS_BADGE[status]} — {q.get('scope_note', '')}")
    else:
        st.warning(f"Đã tạo tệp. Kiểm định: {STATUS_BADGE.get(status, status)} — cần xem các mục bên dưới.")
    bs = report.get("blocks_summary") or {}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Khối dựng native", bs.get("native", 0), help="text_native, math_omml, table_native, drawingml, layout_table")
    c2.metric("Khối giữ bằng ảnh", bs.get("as_image", 0), help="vector_image, raster_image, page_image")
    c3.metric("Khối unresolved", bs.get("unresolved", 0))
    ocr_blocks = sum(1 for b in (report.get("blocks") or {}).get("items", []) if b.get("ocr"))
    c4.metric("Khối từ OCR", ocr_blocks)
    st.caption("Mẫu số = tổng khối trong manifest của job (" + str(bs.get("total", 0)) + "). "
               "Đây là thống kê phương thức dựng, KHÔNG phải độ chính xác thị giác đã đo.")
    if q.get("not_performed"):
        st.info("Kiểm tra bắt buộc CHƯA thực hiện:\n\n- " + "\n- ".join(q["not_performed"]))
    if q.get("reasons"):
        with st.expander(f"Lý do trạng thái ({len(q['reasons'])})", expanded=status != "passed"):
            for r in q["reasons"]:
                st.write("• " + str(r))
    out = report.get("output") or {}
    d1, d2, d3 = st.columns(3)
    if out.get("docx") and Path(out["docx"]).exists():
        d1.download_button(f"📥 Tải Word ({Path(out['docx']).name})", Path(out["docx"]).read_bytes(),
                           file_name=Path(out["docx"]).name, type="primary", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if out.get("report") and Path(out["report"]).exists():
        d2.download_button("📊 Tải báo cáo (.report.json)", Path(out["report"]).read_bytes(),
                           file_name=Path(out["report"]).name, mime="application/json", use_container_width=True)
    if out.get("qa_pdf") and Path(out["qa_pdf"]).exists():
        d3.download_button("🖨️ PDF xem trước (render LibreOffice)", Path(out["qa_pdf"]).read_bytes(),
                           file_name=Path(out["qa_pdf"]).name, mime="application/pdf", use_container_width=True)
    st.subheader("🛡️ Kiểm định 5 nhóm (A–E)")
    for res in report.get("layers", []):
        badge = LAYER_BADGE.get(res["status"], res["status"])
        with st.expander(f"{badge} | Layer {res['layer']} — {res['name']} — {res['summary']}"):
            st.json(res["details"])
    issues = report.get("issues") or []
    if issues:
        with st.expander(f"Cảnh báo / vấn đề ({len(issues)})"):
            for i in issues[:200]:
                st.write(f"[{i.get('severity')}] {i.get('code')}: {i.get('message')}")
    er = report.get("engine_report") or {}
    sp = er.get("smart_paste") or {}
    uf = sp.get("universal_fallback") or {}
    cov = uf.get("do_phu") or {}
    if cov:
        with st.expander("🛡️ Lớp TeX dựng ảnh (universal_fallback)"):
            st.write("Tỷ lệ dưới đây là phân loại khối khi QUÉT mã nguồn (mẫu số = số khối quét được), "
                     "không phải độ giống bản gốc đã đo.")
            st.json({"do_phu": cov, "ket_qua": uf.get("ket_qua"), "ly_do": uf.get("ly_do")})
    with st.expander("Chi tiết khối (manifest)"):
        st.json((report.get("blocks") or {}).get("items", [])[:300])


def main() -> None:
    st.set_page_config(page_title="PDF & LaTeX → Word", page_icon="📐", layout="wide")
    st.title("📐 Chuyển PDF / LaTeX (.tex, .zip, dán mã) → Word (.docx)")
    st.caption("Một pipeline chung cho CLI, FastAPI và Streamlit. Thành phần không tái tạo chắc chắn được giữ bằng "
               "ảnh vùng gốc và được báo cáo; không cam kết giống/sửa được 100%.")
    from backend.capabilities import find_tool
    has_soffice = bool(find_tool("soffice", "libreoffice"))
    with st.sidebar:
        st.header("⚙️ Cấu hình")
        conversion_mode = st.radio("Chế độ:", ["Mode 2: Ưu tiên đối tượng Word sửa được", "Mode 1: Bố cục chuẩn"], index=0)
        mode_int = 2 if conversion_mode.startswith("Mode 2") else 1
        dpi = st.select_slider("Độ phân giải ảnh cắt/chụp (DPI):", options=[200, 300, 400, 600], value=400)
        enable_qa_pdf = st.checkbox("Render Word → PDF và so với nguồn (nhóm E)", value=has_soffice)
        answer_label = st.radio("Đáp án:", ["Giữ như nguồn", "Bản học sinh (không lộ đáp án)"], index=0)
        answer_mode = "student" if answer_label.startswith("Bản học sinh") else "source"
        show_solutions = st.checkbox("Hiện lời giải (\\loigiai)", value=answer_mode == "source")
        figure_policy = st.selectbox("Hình TikZ:", ["auto", "native", "snapshot"], index=0,
                                     help="auto: thử DrawingML rồi đối chiếu với ảnh TeX; native: luôn DrawingML; "
                                          "snapshot: luôn ảnh TeX")
        preprocess = st.selectbox("Tiền xử lý TeX:", ["engine", "smart"], index=0,
                                  help="engine (mặc định): pipeline quản lý khối hình/ảnh + báo cáo từng khối; "
                                       "smart: dùng smart_paste/latex_preprocessor như nhánh dán mã cũ "
                                       "(không có danh sách khối trong báo cáo)")
        strict = st.checkbox("Nghiêm ngặt (báo lỗi nếu chưa 'passed')", value=False)
        st.divider()
        _capability_panel()
    tab_paste, tab_upload = st.tabs(["✍️ Dán mã LaTeX", "📂 Tải .pdf / .tex / .zip"])
    with tab_paste:
        tex_input_code = st.text_area("Mã LaTeX:", value=DEFAULT_LATEX_TEMPLATE, height=380)
        convert_from_paste = st.button("🚀 Chuyển sang Word", type="primary", use_container_width=True)
    with tab_upload:
        uploaded_file = st.file_uploader("Chọn tệp `.pdf`, `.tex` hoặc gói `.zip`:", type=["pdf", "tex", "zip"])
        main_tex_name = st.text_input("Tệp chính trong gói .zip (đường dẫn tương đối, để trống để tự xác định):",
                                      value="", placeholder="vd. chuong/main.tex")
        convert_from_upload = st.button("🚀 Chuyển tệp sang Word", type="primary", use_container_width=True,
                                        disabled=(uploaded_file is None))
    if not (convert_from_paste or convert_from_upload):
        return
    opts = dict(jobs_root=JOBS_DIR, mode=mode_int, dpi=dpi, qa_pdf=enable_qa_pdf, strict=strict,
                answer_mode=answer_mode, show_solutions=show_solutions, figure_policy=figure_policy,
                preprocess=preprocess)
    with st.status("Đang chuyển đổi…", expanded=True) as box:
        try:
            if convert_from_paste:
                filename = "document.tex"
                report = convert_paste(tex_input_code, **opts)
            else:
                assert uploaded_file is not None
                filename = uploaded_file.name
                report = convert_bytes(filename, uploaded_file.getvalue(), main_file=main_tex_name or None, **opts)
        except ConversionError as exc:
            box.update(label="Không chuyển đổi được", state="error")
            st.error(str(exc))
            cands = exc.extra.get("candidates")
            if cands:
                st.write("Các tệp có thể chọn làm tệp chính (nhập vào ô 'Tệp chính'):")
                st.json(cands)
            return
        box.update(label=f"Xong — {STATUS_BADGE.get(report.get('quality_status'), report.get('quality_status'))}",
                   state="complete", expanded=False)
    if strict and report.get("quality_status") != "passed":
        st.error("Chế độ nghiêm ngặt: kết quả chưa đạt 'passed'. Tệp vẫn được tạo để bạn kiểm tra.")
    _show_report(report, filename)
    with st.expander("Báo cáo đầy đủ (JSON)"):
        st.code(json.dumps({k: v for k, v in report.items() if k not in ("blocks",)}, ensure_ascii=False,
                           indent=2, default=str)[:200000], language="json")


if __name__ == "__main__":
    main()
