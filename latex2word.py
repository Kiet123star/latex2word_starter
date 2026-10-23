#!/usr/bin/env python3
"""LaTeX / PDF → Word CLI (`latex2word.py`) — dùng CHUNG pipeline với FastAPI và Streamlit.

    python latex2word.py de_thi.tex -o de_thi.docx --qa-pdf
    python latex2word.py du_an.zip --main chuong/main.tex -o ket_qua.docx
    python latex2word.py tai_lieu.pdf -o tai_lieu.docx --qa-pdf

Mã thoát: 0 = tạo được .docx (và nếu --strict thì phải `passed`); 2 = --strict mà chưa đạt; 1 = lỗi đầu vào/chuyển đổi.
Báo cáo đầy đủ được ghi cạnh tệp kết quả: <tên>.report.json
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from backend.pipeline import ConversionError, convert_path

ENGINE_KEYS = ("equation_nodes", "word_tables", "drawingml_groups", "drawingml_shapes", "tkz_tab_tables",
               "choice_tables", "data_tables", "immini_layouts", "embedded_media_files", "ooxml_stats",
               "qa_rendered_pdf", "warnings")


def convert(
    tex: Path,
    output: Path,
    *,
    engine: str = "xelatex",
    dpi: int = 400,
    reference: Path | None = None,
    strict: bool = False,
    qa_pdf: bool = False,
    mode: int = 2,
    main_file: str | None = None,
    answer_mode: str = "source",
    show_solutions: bool = True,
    figure_policy: str = "auto",
    ocr: str = "auto",
    ocr_lang: str | None = None,
    jobs_dir: Path | None = None,
    preprocess: str | None = None,
) -> dict:
    """Chuyển .tex/.zip/.pdf → .docx qua pipeline chung; chép kết quả về `output` (+ .report.json).

    Giữ chữ ký cũ (tex, output, engine, dpi, reference, strict, qa_pdf, mode). Trả về báo cáo hợp nhất; các khoá
    số liệu cũ của engine (equation_nodes, word_tables, …) được nhân bản ở cấp đầu cho caller cũ — đều là số đếm
    trên tệp CUỐI CÙNG."""
    tex, output = Path(tex), Path(output)
    own_tmp = None
    if jobs_dir is None:
        own_tmp = tempfile.mkdtemp(prefix="l2w_cli_jobs_")
        jobs_dir = Path(own_tmp)
    report = convert_path(tex, jobs_root=jobs_dir, engine=engine, dpi=dpi, strict=strict, qa_pdf=qa_pdf, mode=mode,
                          main_file=main_file, answer_mode=answer_mode, show_solutions=show_solutions,
                          figure_policy=figure_policy, ocr=ocr, ocr_lang=ocr_lang, preprocess=preprocess)
    if reference is not None:
        report.setdefault("issues", []).append({"severity": "info", "code": "reference_docx_ignored",
                                                "message": "--reference chưa được pipeline áp dụng (giữ tham số cho tương thích)."})
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(report["output"]["docx"], output)
    qa = report["output"].get("qa_pdf")
    if qa:
        qa_dst = output.with_name(output.stem + "_qa.pdf")
        shutil.copy2(qa, qa_dst)
        report["output"]["qa_pdf_copy"] = str(qa_dst)
    report["output"]["docx_copy"] = str(output)
    for k in ENGINE_KEYS:
        if k in (report.get("engine_report") or {}):
            report[k] = report["engine_report"][k]
    output.with_suffix(".report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str),
                                                  encoding="utf-8")
    if own_tmp:
        report["job_dir_note"] = f"thư mục job tạm: {own_tmp}"
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("tex", type=Path, help="Tệp .tex, gói .zip hoặc .pdf")
    ap.add_argument("-o", "--output", type=Path, help="Tệp .docx đích")
    ap.add_argument("--engine", choices=["pdflatex", "xelatex", "lualatex"], default="xelatex")
    ap.add_argument("--dpi", type=int, default=400, help="Độ phân giải ảnh cắt/chụp (mặc định 400)")
    ap.add_argument("--reference", type=Path, help="(giữ cho tương thích) reference.docx")
    ap.add_argument("--strict", action="store_true", help="Mã thoát 2 nếu kết quả chưa 'passed'")
    ap.add_argument("--qa-pdf", action="store_true", help="Render Word → PDF và so với nguồn (nhóm E)")
    ap.add_argument("--mode", type=int, choices=[1, 2], default=2)
    ap.add_argument("--main", dest="main_file", help="Tệp chính trong gói .zip (đường dẫn tương đối)")
    ap.add_argument("--answer-mode", choices=["source", "student"], default="source",
                    help="student: không làm lộ đáp án (bỏ \\True, \\choiceA..D, cột Đ/S)")
    ap.add_argument("--no-solutions", action="store_true", help="Ẩn lời giải (\\loigiai)")
    ap.add_argument("--figure-policy", choices=["auto", "native", "snapshot"], default="auto")
    ap.add_argument("--ocr", choices=["auto", "off"], default="auto")
    ap.add_argument("--ocr-lang", default=None, help="vd. vie+eng hoặc eng")
    ap.add_argument("--jobs-dir", type=Path, default=None, help="Thư mục chứa job (mặc định: thư mục tạm)")
    ap.add_argument("--preprocess", choices=["engine", "smart", "auto"], default=None,
                    help="engine (mặc định): pipeline quản lý khối + provenance; smart: dùng smart_paste/latex_preprocessor")
    args = ap.parse_args()
    out = args.output or args.tex.with_suffix(".docx")
    try:
        result = convert(args.tex, out, engine=args.engine, dpi=args.dpi, reference=args.reference, strict=args.strict,
                         qa_pdf=args.qa_pdf, mode=args.mode, main_file=args.main_file, answer_mode=args.answer_mode,
                         show_solutions=not args.no_solutions, figure_policy=args.figure_policy, ocr=args.ocr,
                         ocr_lang=args.ocr_lang, jobs_dir=args.jobs_dir, preprocess=args.preprocess)
    except ConversionError as error:
        print("ERROR:", error, file=sys.stderr)
        cands = error.extra.get("candidates")
        if cands:
            print("Ứng viên:", json.dumps(cands, ensure_ascii=False), file=sys.stderr)
        return 1
    except (RuntimeError, OSError, UnicodeError, zipfile.BadZipFile) as error:
        print("ERROR:", error, file=sys.stderr)
        return 1
    summary = {k: result.get(k) for k in ("job_id", "conversion_completed", "quality_status", "blocks_summary")}
    summary["layers"] = [f"L{l['layer']} {l['status']}: {l['summary']}" for l in result.get("layers", [])]
    summary["reasons"] = (result.get("quality") or {}).get("reasons", [])[:10]
    summary["output"] = str(out)
    summary["report"] = str(out.with_suffix(".report.json"))
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    if args.strict and result.get("quality_status") != "passed":
        return 2
    return 0 if result.get("conversion_completed") else 1


if __name__ == "__main__":
    raise SystemExit(main())
