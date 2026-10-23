"""`backend/pipeline.py` — ĐIỀU PHỐI CHUNG cho mọi đầu vào: dán LaTeX, .tex, .zip, .pdf.

    nhận bytes ─► job riêng (input/ chỉ đọc, work/, out/) ─► nhận diện loại theo NỘI DUNG
      ├─ TeX / ZIP chứa TeX: giải nén an toàn → chọn tệp chính → nạp include/tài nguyên (source_resolver)
      │     → MẶC ĐỊNH: [macro/danh sách/môi trường + khối hình/ảnh (tex_blocks) → engine → khôi phục khối]
      │       (preprocess="smart": giao cho smart_paste/latex_preprocessor — KHÔNG có provenance từng khối)
      ├─ PDF: phân tích trang/vùng/chữ/hình/bảng/công thức (pdf_reader) → dựng Word (pdf_docx_builder)
      └─► hậu xử lý chung: sửa dấu OMML, chuẩn hoá thứ tự XML, kiểm tra gói, đếm lại trên tệp CUỐI
          → render QA + so với nguồn THẬT → 5 nhóm kiểm định → quality_status

API cũ (`convert_latex_file_unified`, `convert_pasted_latex_unified`) vẫn giữ nguyên; module này là lớp
điều phối mà CLI / Streamlit / FastAPI cùng gọi.
"""
from __future__ import annotations

import inspect
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import time
import traceback
from pathlib import Path
from typing import Any

from .job_workspace import (InputKindError, JobWorkspace, ZipSafetyError, safe_extract_zip, sanitize_upload_name,
                            sniff_kind)
from .source_resolver import (MainSelectionError, SourceEncodingError, decode_tex_bytes, flatten_project,
                              select_main)

__all__ = ["convert_bytes", "convert_path", "convert_paste", "render_docx_to_pdf", "ConversionError",
           "REPORT_SCHEMA"]

REPORT_SCHEMA = "l2w-report/2"
STATIC_LIMITATIONS = [
    "Không cam kết giống 100% hay sửa được 100%: ảnh SVG/PNG nhúng là ảnh, không phải OMML/Word Shapes.",
    "‘passed’ chỉ áp dụng cho các tiêu chí tự động đã nêu trong báo cáo, không có nghĩa đúng mọi tài liệu.",
    "Render QA dùng LibreOffice; Microsoft Word có thể ngắt dòng/trang khác.",
    "Nhận dạng công thức từ PDF là heuristic (chỉ phân số/chỉ số đơn giản); công thức không chắc → ảnh vùng.",
]


class ConversionError(RuntimeError):
    def __init__(self, message: str, *, code: str = "conversion_error", http_status: int = 400,
                 extra: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.http_status = http_status
        self.extra = extra or {}


# ───────────────────────────── tiện ích ─────────────────────────────
def _call_with_supported_kwargs(fn, *args: Any, **kwargs: Any) -> Any:
    try:
        params = inspect.signature(fn).parameters
        if not any(p.kind == p.VAR_KEYWORD for p in params.values()):
            kwargs = {k: v for k, v in kwargs.items() if k in params}
    except (TypeError, ValueError):
        pass
    return fn(*args, **kwargs)


def render_docx_to_pdf(docx_path: Path, out_dir: Path, *, timeout: int = 240) -> dict[str, Any]:
    """DOCX → PDF bằng LibreOffice với hồ sơ người dùng RIÊNG cho job (tránh xung đột khi chạy đồng thời)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    from .capabilities import find_tool
    lo = find_tool("soffice", "libreoffice")
    if lo:
        prof = out_dir / "lo_profile"
        t0 = time.time()
        try:
            cp = subprocess.run([lo, "--headless", "-env:UserInstallation=" + prof.resolve().as_uri(),
                                 "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
                                capture_output=True, timeout=timeout)
            pdf = out_dir / (docx_path.stem + ".pdf")
            if pdf.exists():
                return {"ok": True, "pdf": str(pdf), "engine": "LibreOffice", "seconds": round(time.time() - t0, 1)}
            return {"ok": False, "error": (cp.stderr or b"").decode("utf-8", "replace")[-400:] or "không tạo được PDF",
                    "engine": "LibreOffice"}
        except subprocess.TimeoutExpired:
            return {"ok": False, "error": f"LibreOffice quá {timeout}s", "engine": "LibreOffice"}
        finally:
            shutil.rmtree(prof, ignore_errors=True)
    if sys.platform == "win32":
        try:
            import docx2pdf  # type: ignore
            pdf = out_dir / (docx_path.stem + ".pdf")
            docx2pdf.convert(str(docx_path), str(pdf))
            if pdf.exists():
                return {"ok": True, "pdf": str(pdf), "engine": "Microsoft Word (docx2pdf)"}
        except Exception as exc:
            return {"ok": False, "error": f"docx2pdf: {type(exc).__name__}: {exc}", "engine": "Microsoft Word"}
    return {"ok": False, "error": "máy không có LibreOffice/Microsoft Word để render", "engine": None}


def _environment() -> dict[str, Any]:
    from .capabilities import probe_capabilities
    caps = probe_capabilities()
    return {"python": platform.python_version(), "platform": platform.platform(),
            "libraries": {k: v.get("version") for k, v in caps["libraries"].items() if v.get("available")},
            "tools": {k: v.get("version") for k, v in caps["tools"].items() if v.get("available")},
            "tesseract_languages": caps.get("tesseract_languages", [])}


def _postprocess_final_docx(docx_path: Path) -> dict[str, Any]:
    """Sửa dấu OMML (lớp bảo hiểm) + chuẩn hoá thứ tự phần tử thuộc tính. Ghi đè tệp KẾT QUẢ (không phải nguồn)."""
    out: dict[str, Any] = {}
    try:
        from .omml_fix import fix_math_accents
        out["accent_fix"] = fix_math_accents(docx_path)
    except Exception as exc:
        out["accent_fix"] = {"error": f"{type(exc).__name__}: {exc}"}
    try:
        import docx as _docx
        from .ooxml_tools import normalize_xml_order
        d = _docx.Document(str(docx_path))
        n = normalize_xml_order(d.element)
        for sec in d.sections:
            for hf in (sec.header, sec.footer):
                try:
                    n += normalize_xml_order(hf._element)
                except Exception:
                    pass
        if n:
            d.save(str(docx_path))
        out["xml_order_fixed"] = n
    except Exception as exc:
        out["xml_order_fixed"] = {"error": f"{type(exc).__name__}: {exc}"}
    return out


def _new_report(ws: JobWorkspace, kind: str, name: str) -> dict[str, Any]:
    return {
        "schema": REPORT_SCHEMA, "job_id": ws.job_id, "started_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "input": {"name": name, "kind": kind, **(ws.inputs[-1] if ws.inputs else {})},
        "conversion_completed": False, "quality_status": "failed", "status": "failed", "overall": "failed",
        "issues": [], "steps": [], "limitations": list(STATIC_LIMITATIONS),
    }


def _step(report: dict[str, Any], name: str, t0: float, **info: Any) -> None:
    report["steps"].append({"step": name, "seconds": round(time.time() - t0, 2), **info})


def _issue(report: dict[str, Any], severity: str, code: str, message: str, **kw: Any) -> None:
    report["issues"].append({"severity": severity, "code": code, "message": message, **kw})


# ───────────────────────────── nhánh TeX ─────────────────────────────
def _convert_tex_project(ws: JobWorkspace, project_root: Path, report: dict[str, Any], *, main_file: str | None,
                         opts: dict[str, Any], out_stem: str) -> Path:
    t0 = time.time()
    sel = select_main(project_root, main_file)
    report["source"] = {"main": sel["main"], "selection": sel["how"],
                        "candidates": [c["path"] for c in sel["candidates"]]}
    flat = flatten_project(project_root, sel["main"])
    report["source"].update(flat.summary())
    for iss in flat.issues:
        _issue(report, iss["severity"], iss["code"], iss["message"], file=iss.get("file"), line=iss.get("line"))
    _step(report, "resolve_source", t0, files=len(flat.files))
    main_path = (project_root / sel["main"]).resolve()
    main_dir = main_path.parent
    out_docx = ws.out_dir / f"{out_stem}.docx"
    from .capabilities import probe_capabilities
    caps = probe_capabilities()
    pre_mode = opts.get("preprocess", "engine")
    smart_ok = caps["derived"]["paste_smart"]["available"]
    use_smart = pre_mode == "smart" or (pre_mode == "auto" and smart_ok)
    if pre_mode == "smart" and not smart_ok:
        raise ConversionError("Yêu cầu tiền xử lý thông minh nhưng không nạp được smart_paste: "
                              + caps["derived"]["paste_smart"]["detail"], code="smart_unavailable", http_status=503)
    if not caps["derived"]["tex_upload"]["available"]:
        raise ConversionError("Không có engine dựng Word từ TeX: " + caps["derived"]["tex_upload"]["detail"],
                              code="engine_unavailable", http_status=503)
    from . import math_tables as mt
    from . import omml_math as om
    from .core_converter import convert_latex_file_unified
    report["preprocess"] = {"mode": "smart" if use_smart else "engine", "requested": pre_mode,
                            "reason": (("theo yêu cầu (preprocess=" + pre_mode + ")") if use_smart else
                                       ("mặc định: nhánh engine của pipeline — có quản lý khối hình/ảnh, provenance "
                                        "từng khối, thông báo khối thiếu" if pre_mode == "engine" else
                                        caps["derived"]["paste_smart"]["detail"]))}
    if use_smart:
        _issue(report, "info", "smart_mode_no_block_provenance",
               "Chế độ preprocess=smart: khối hình/ảnh do smart_paste/universal_fallback xử lý — báo cáo KHÔNG có "
               "danh sách khối (provenance) và kiểm định nhóm C chỉ dựa trên tệp Word cuối.")
    answer_tok = mt.ANSWER_MODE.set(opts.get("answer_mode", "source"))
    events = mt.start_table_events()
    mstats = om.start_math_stats()
    try:
        t1 = time.time()
        if use_smart:
            try:
                flat_tex = main_dir / f"{main_path.stem}.l2w_flat.tex"
                flat_tex.write_text(flat.text, encoding="utf-8")
                from .smart_paste import convert_pasted_latex_unified
                engine_report = _call_with_supported_kwargs(
                    convert_pasted_latex_unified, flat_tex, out_docx, engine=opts["engine"], dpi=opts["dpi"],
                    strict=opts["strict"], qa_pdf=False, mode=opts["mode"], show_solutions=opts["show_solutions"],
                    answer_mode=opts["answer_mode"])
                validation_text = flat.text
                report["blocks"] = {"items": [], "note": "khối do lớp universal_fallback quản lý (xem engine_report.smart_paste)"}
            except ConversionError:
                raise
            except Exception as exc:
                if pre_mode == "smart":
                    raise
                # Nhánh thông minh lỗi → KHÔNG dừng job: chạy nhánh engine của pipeline và ghi rõ lý do.
                err = f"{type(exc).__name__}: {exc}"
                _issue(report, "warning", "smart_path_failed",
                       "Tiền xử lý thông minh (smart_paste) lỗi → đã chuyển bằng nhánh engine của pipeline: " + err[:400])
                report["preprocess"] = {"mode": "engine", "reason": "smart_paste lỗi: " + err[:300],
                                        "smart_trace": traceback.format_exc()[-1500:]}
                use_smart = False
        if not use_smart:
            from .tex_blocks import inject_engine_blocks, prepare_tex_for_engine
            prep = prepare_tex_for_engine(flat, project_root, ws.work_dir, main_rel=sel["main"], dpi=opts["dpi"],
                                          figure_policy=opts["figure_policy"], answer_mode=opts["answer_mode"],
                                          show_solutions=opts["show_solutions"], tex_engine=opts["engine"])
            _step(report, "prepare_blocks", t1, blocks=len(prep.blocks))
            eng_tex = main_dir / f"{main_path.stem}.l2w_engine.tex"
            eng_tex.write_text(prep.text, encoding="utf-8")
            t2 = time.time()
            engine_report = _call_with_supported_kwargs(
                convert_latex_file_unified, eng_tex, out_docx, engine=opts["engine"], dpi=opts["dpi"],
                strict=opts["strict"], qa_pdf=False, mode=opts["mode"])
            _step(report, "engine", t2)
            t3 = time.time()
            inj = inject_engine_blocks(out_docx, prep)
            _step(report, "inject_blocks", t3, **inj.get("blocks_by_status", {}))
            for iss in prep.issues:
                _issue(report, iss["severity"], iss["code"], iss["message"], block_id=iss.get("block_id"),
                       source=iss.get("source"))
            for b in prep.blocks:
                if b.status == "unresolved" and not any(i.get("block_id") == b.block_id for i in prep.issues):
                    _issue(report, "warning", f"{b.kind}_unresolved", f"{b.block_id}: {b.reason}", block_id=b.block_id,
                           source=b.source)
            report["transforms"] = prep.transforms
            report["blocks"] = {"items": [b.to_dict() for b in prep.blocks], "injection": inj}
            validation_text = prep.validation_text
            report["_normalized_text"] = prep.normalized_text
        engine_report = dict(engine_report or {})
        engine_report.pop("qa_rendered_pdf", None)
        report["engine_report"] = engine_report
        report["table_events"] = events[:500]
        report["math_stats"] = mstats
    finally:
        mt.ANSWER_MODE.reset(answer_tok)
    report["_validation_text"] = validation_text
    report["_reference_tex"] = str(main_path)
    for ev in events:
        if ev["kind"] == "bbt_unresolved":
            _issue(report, "warning", "bbt_unresolved", "Bảng biến thiên chưa dựng được: " + ev.get("reason", ""),
                   source_excerpt=ev.get("source", "")[:300])
        elif ev["kind"] == "table_row_overflow":
            _issue(report, "warning", "table_row_overflow",
                   f"Hàng {ev.get('row')} của một bảng có nhiều ô hơn số cột khai báo: {ev.get('cells')} ô dư được ghép "
                   "vào ô cuối (không bỏ nội dung) — cần kiểm tra bố cục bảng.")
    if mstats.get("unknown_commands"):
        _issue(report, "warning", "math_unknown_commands",
               "Lệnh toán engine chưa hiểu (in ra dạng chữ): " + ", ".join(sorted(mstats["unknown_commands"]))[:400])
    if mstats.get("unverified_constructs"):
        _issue(report, "info", "math_unverified_construct",
               "Ký hiệu cung/ngoặc nhóm (m:groupChr) chưa kiểm chứng hiển thị trên Microsoft Word: "
               + ", ".join(sorted(mstats["unverified_constructs"])))
    return out_docx


# ───────────────────────────── nhánh PDF ─────────────────────────────
def _convert_pdf(ws: JobWorkspace, pdf_path: Path, report: dict[str, Any], *, opts: dict[str, Any],
                 out_stem: str) -> Path:
    t0 = time.time()
    from .pdf_reader import PdfInputError, analyze_pdf
    from .pdf_docx_builder import build_docx_from_manifest
    try:
        manifest = analyze_pdf(pdf_path, ws.work_dir / "pdf", dpi=opts["dpi"], ocr=opts.get("ocr", "auto"),
                               ocr_lang=opts.get("ocr_lang"))
    except PdfInputError as exc:
        raise ConversionError(str(exc), code=exc.code, http_status=422) from exc
    _step(report, "analyze_pdf", t0, pages=manifest["page_count"], blocks=len(manifest["blocks"]))
    out_docx = ws.out_dir / f"{out_stem}.docx"
    t1 = time.time()
    build = build_docx_from_manifest(manifest, out_docx)
    _step(report, "build_docx", t1)
    report["source"] = {"pdf": pdf_path.name, "pages": manifest["page_count"], "page_types": manifest["page_types"],
                        "document": manifest.get("document", {})}
    report["pdf_manifest_summary"] = manifest["summary"]
    report["blocks"] = {"items": manifest["blocks"], "build": build}
    report["_pdf_manifest"] = manifest
    for iss in manifest.get("issues", []):
        _issue(report, iss["severity"], iss["code"], iss["message"], page=iss.get("page"), block_id=iss.get("block_id"))
    report["engine_report"] = {"engine": "pdf_reader + pdf_docx_builder", **build.get("counts", {})}
    return out_docx


# ───────────────────────────── chung ─────────────────────────────
def _finish(ws: JobWorkspace, report: dict[str, Any], out_docx: Path, *, opts: dict[str, Any]) -> dict[str, Any]:
    from .ooxml_tools import docx_inventory, validate_package
    from .validators import generate_visual_comparison, run_all_layers, run_pdf_layers, summarize_quality
    t0 = time.time()
    report["postprocess"] = _postprocess_final_docx(out_docx)
    pkg = validate_package(out_docx)
    report["package"] = pkg
    report["conversion_completed"] = bool(pkg["ok"])
    inv = docx_inventory(out_docx) if pkg["ok"] else {}
    report["counts"] = {k: v for k, v in inv.items() if k not in ("tables", "image_names")}
    er = report.get("engine_report") or {}
    if inv:
        # số liệu báo cáo phải là của tệp CUỐI CÙNG
        er["engine_counts_before_postprocess"] = {k: er.get(k) for k in (
            "equation_nodes", "word_tables", "drawingml_groups", "drawingml_shapes", "embedded_media_files")}
        er.update({"equation_nodes": inv["equation_nodes"], "word_tables": inv["word_tables"],
                   "drawingml_groups": inv["drawingml_groups"], "drawingml_shapes": inv["drawingml_shapes"],
                   "embedded_media_files": inv["logical_images"]})
        report["engine_report"] = er
    _step(report, "package_check", t0, ok=pkg["ok"])
    visual_info: dict[str, Any] = {"enabled": False, "measured": False, "similarity_percent": None}
    is_pdf = report["input"]["kind"] == "pdf" or report.get("_pdf_manifest") is not None
    if opts["qa_pdf"] and pkg["ok"]:
        t1 = time.time()
        qa_dir = ws.out_dir / "qa"
        rr = render_docx_to_pdf(out_docx, qa_dir)
        report["qa_render"] = {k: v for k, v in rr.items()}
        from .validators import renderer_supports_omml
        om_ok = renderer_supports_omml(ws.work_dir / "qa_probe")
        report["qa_render"]["omml_rendered"] = om_ok
        if om_ok.get("checked") and not om_ok.get("supported"):
            _issue(report, "warning", "qa_renderer_drops_omml",
                   "Trình render QA (LibreOffice trên máy này) KHÔNG vẽ công thức OMML (thiếu module Math): PDF xem "
                   "trước sẽ thiếu công thức dù .docx có OMML. Nhóm E chỉ đo chữ và bố cục.")
        if rr.get("ok"):
            report["qa_rendered_pdf"] = rr["pdf"]
            er["qa_rendered_pdf"] = rr["pdf"]
            omml_words = None
            if om_ok.get("checked") and not om_ok.get("supported"):
                from .validators import omml_words_in_docx
                omml_words = omml_words_in_docx(out_docx)
            if is_pdf:
                ref = Path(report["_pdf_source"])
                visual_info, pages = generate_visual_comparison(ws.work_dir / "qa", None, Path(rr["pdf"]),
                                                                reference_pdf=ref, mode="pdf", omml_words=omml_words)
            else:
                visual_info, pages = generate_visual_comparison(ws.work_dir / "qa", Path(report["_reference_tex"]),
                                                                Path(rr["pdf"]), compile_reference=True, mode="tex",
                                                                omml_words=omml_words)
            report["visual_pages"] = pages[:300]
        else:
            visual_info["reason"] = "Không render được Word → PDF: " + str(rr.get("error"))
        _step(report, "visual_qa", t1, measured=visual_info.get("measured"))
    elif not opts["qa_pdf"]:
        visual_info["reason"] = "QA hiển thị bị tắt (qa_pdf=False) → nhóm E chưa thực hiện."
    if pkg["ok"]:
        if is_pdf:
            layers = run_pdf_layers(report["_pdf_manifest"], out_docx, visual_info)
        else:
            fig_blocks = [b for b in (report.get("blocks") or {}).get("items", [])
                          if b.get("kind") in ("figure", "includegraphics", "bbt")]
            layers = run_all_layers(report["_validation_text"], out_docx, visual_info,
                                    manifest={"figure_blocks": [b for b in fig_blocks if b.get("kind") != "bbt"]},
                                    normalized_tex=report.get("_normalized_text"))
    else:
        layers = []
    report["layers"] = layers
    items = (report.get("blocks") or {}).get("items", [])
    unresolved = [b for b in items if b.get("status") == "unresolved" or b.get("method") == "unresolved"]
    by_method: dict[str, int] = {}
    for b in items:
        by_method[b.get("method") or "?"] = by_method.get(b.get("method") or "?", 0) + 1
    report["blocks_summary"] = {
        "total": len(items), "by_method": by_method, "unresolved": len(unresolved),
        "native": sum(v for k, v in by_method.items() if k in ("text_native", "math_omml", "table_native",
                                                               "drawingml", "layout_table")),
        "as_image": sum(v for k, v in by_method.items() if k in ("vector_image", "raster_image", "page_image")),
    }
    blocking = []
    review = [i["message"] for i in report["issues"] if i["severity"] in ("error", "warning")][:30]
    if inv and inv.get("marker_leaks_total"):
        review.append(f"{inv['marker_leaks_total']} chỗ đánh dấu nội bộ còn trong Word")
    q = summarize_quality(layers, conversion_completed=report["conversion_completed"], package_ok=pkg["ok"],
                          unresolved_blocks=len(unresolved), blocking_issues=blocking, review_issues=review)
    report["quality"] = q
    report["quality_status"] = q["quality_status"]
    report["status"] = q["quality_status"]
    report["overall"] = q["quality_status"]
    report["input_preserved"] = ws.verify_inputs_unchanged()
    if not report["input_preserved"]:
        report["quality_status"] = report["status"] = report["overall"] = "failed"
        _issue(report, "error", "input_modified", "Tệp đầu vào bị thay đổi trong lúc xử lý (lỗi nội bộ).")
    report["output"] = {"docx": str(out_docx), "report": str(out_docx.with_suffix(".report.json")),
                        "qa_pdf": report.get("qa_rendered_pdf")}
    report["environment"] = _environment()
    from .capabilities import capability_summary
    report["capabilities"] = capability_summary()
    caps_d = report["capabilities"]["derived"]
    if not caps_d["ocr"]["available"]:
        report["limitations"].append("OCR chưa có trên máy → trang scan giữ ảnh trang.")
    elif not caps_d["ocr"].get("vietnamese", "vie" in report["capabilities"].get("tesseract_languages", [])):
        report["limitations"].append("OCR có nhưng thiếu gói tiếng Việt (vie.traineddata) → trang scan tiếng Việt giữ ảnh.")
    if not caps_d["paste_smart"]["available"]:
        report["limitations"].append("Tiền xử lý thông minh (latex_preprocessor/tex_truth) không nạp được trên máy này "
                                     "→ nhánh TeX dùng bộ chuẩn hoá của pipeline + engine.")
    report["finished_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    clean = {k: v for k, v in report.items() if not k.startswith("_")}
    Path(report["output"]["report"]).write_text(json.dumps(clean, ensure_ascii=False, indent=2, default=str),
                                                encoding="utf-8")
    return clean


def _opts(**kw: Any) -> dict[str, Any]:
    o = {"mode": 2, "dpi": 400, "strict": False, "qa_pdf": True, "engine": "xelatex", "answer_mode": "source",
         "show_solutions": True, "figure_policy": "auto",
         "preprocess": os.environ.get("L2W_PREPROCESS", "engine"), "ocr": "auto", "ocr_lang": None}
    o.update({k: v for k, v in kw.items() if v is not None})
    if o["preprocess"] not in ("engine", "smart", "auto"):
        raise ConversionError(f"preprocess không hợp lệ: {o['preprocess']} (engine | smart | auto)", code="bad_option")
    if o["answer_mode"] not in ("source", "student"):
        raise ConversionError(f"answer_mode không hợp lệ: {o['answer_mode']}", code="bad_option")
    if o["figure_policy"] not in ("auto", "native", "snapshot"):
        raise ConversionError(f"figure_policy không hợp lệ: {o['figure_policy']}", code="bad_option")
    o["dpi"] = max(72, min(int(o["dpi"]), 1200))
    return o


def convert_bytes(filename: str | None, data: bytes, *, jobs_root: Path, job_id: str | None = None,
                  main_file: str | None = None, **options: Any) -> dict[str, Any]:
    """Điểm vào chung cho tải tệp (.pdf/.tex/.zip). Ném ConversionError (có http_status) khi đầu vào không dùng được."""
    opts = _opts(**options)
    ws = JobWorkspace(jobs_root, job_id)
    stored = ws.store_input(filename, data)
    sniff = sniff_kind(data, stored.name)
    report = _new_report(ws, sniff["kind"], stored.name)
    report["input"]["detection"] = sniff
    report["options"] = opts
    for w in sniff["warnings"]:
        _issue(report, "warning", "extension_mismatch", w)
    out_stem = Path(stored.name).stem or "document"
    try:
        src_root = ws.work_dir / "src"
        src_root.mkdir(parents=True, exist_ok=True)
        if sniff["kind"] == "zip":
            t0 = time.time()
            ex = safe_extract_zip(stored, src_root)
            report["zip"] = ex.to_dict()
            _step(report, "safe_extract", t0, files=len(ex.files))
            for w in ex.warnings:
                _issue(report, "warning", "zip_name", w)
            tex_files = [f for f in ex.files if f.lower().endswith((".tex", ".ltx", ".latex"))]
            pdf_files = [f for f in ex.files if (src_root / f).read_bytes()[:1024].find(b"%PDF-") >= 0]
            want = (main_file or "").strip()
            if want.lower().endswith(".pdf") or (not tex_files and pdf_files):
                if want:
                    cand = [f for f in pdf_files if f == want or Path(f).name == want]
                    if len(cand) != 1:
                        raise ConversionError(f"Không xác định được PDF '{want}' trong gói.", code="main_file",
                                              http_status=409, extra={"candidates": pdf_files})
                    pdf_rel = cand[0]
                elif len(pdf_files) == 1:
                    pdf_rel = pdf_files[0]
                else:
                    raise ConversionError("Gói có nhiều PDF: " + ", ".join(pdf_files) + ". Hãy chọn tệp chính.",
                                          code="main_file_ambiguous", http_status=409, extra={"candidates": pdf_files})
                report["input"]["kind"] = "pdf"
                report["_pdf_source"] = str(src_root / pdf_rel)
                docx_out = _convert_pdf(ws, src_root / pdf_rel, report, opts=opts, out_stem=Path(pdf_rel).stem)
                if tex_files:
                    _issue(report, "info", "zip_has_tex_and_pdf",
                           "Gói có cả .tex và .pdf; đã chuyển PDF theo lựa chọn. Không tự coi .tex là nguồn của PDF.")
            elif tex_files:
                report["input"]["kind"] = "zip_tex"
                docx_out = _convert_tex_project(ws, src_root, report, main_file=main_file or None, opts=opts,
                                                out_stem=out_stem)
                if pdf_files:
                    report.setdefault("source", {})["pdf_in_package"] = pdf_files
            else:
                raise ConversionError("Không tìm thấy tệp .tex hoặc .pdf nào trong gói .zip.", code="zip_empty")
        elif sniff["kind"] == "pdf":
            pdf_copy = ws.copy_to_work(stored, "src/" + stored.name)
            report["_pdf_source"] = str(pdf_copy)
            docx_out = _convert_pdf(ws, pdf_copy, report, opts=opts, out_stem=out_stem)
        elif sniff["kind"] == "tex":
            name = stored.name if stored.suffix.lower() in (".tex", ".ltx", ".latex") else stored.stem + ".tex"
            shutil.copy2(stored, src_root / name)
            docx_out = _convert_tex_project(ws, src_root, report, main_file=name, opts=opts, out_stem=out_stem)
        else:
            raise ConversionError("Không nhận diện được tệp (không phải PDF, ZIP hay mã TeX).", code="unknown_kind",
                                  http_status=415)
    except MainSelectionError as exc:
        raise ConversionError(str(exc), code=exc.code, http_status=409,
                              extra={"candidates": exc.candidates, "job_id": ws.job_id}) from exc
    except (ZipSafetyError, SourceEncodingError) as exc:
        raise ConversionError(str(exc), code=type(exc).__name__, http_status=400, extra={"job_id": ws.job_id}) from exc
    except ConversionError as exc:
        exc.extra.setdefault("job_id", ws.job_id)
        raise
    except Exception as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"
        report["trace"] = traceback.format_exc()[-4000:]
        (ws.out_dir / "error.report.json").write_text(json.dumps(
            {k: v for k, v in report.items() if not k.startswith("_")}, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8")
        raise ConversionError(f"Lỗi khi chuyển đổi: {type(exc).__name__}: {exc}", code="internal_error",
                              http_status=500, extra={"job_id": ws.job_id}) from exc
    return _finish(ws, report, docx_out, opts=opts)


def convert_path(path: Path, *, jobs_root: Path, **kw: Any) -> dict[str, Any]:
    path = Path(path)
    return convert_bytes(path.name, path.read_bytes(), jobs_root=jobs_root, **kw)


def convert_paste(latex_source: str, *, jobs_root: Path, filename: str = "document.tex", **kw: Any) -> dict[str, Any]:
    """Nhánh dán mã đi qua CÙNG pipeline (đọc nguồn → chuẩn hoá → engine → hậu xử lý → kiểm định)."""
    text = latex_source.replace("\r\n", "\n").replace("\r", "\n")
    return convert_bytes(filename, text.encode("utf-8"), jobs_root=jobs_root, **kw)
