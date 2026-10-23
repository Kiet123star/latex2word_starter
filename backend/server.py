"""FastAPI REST API (`backend/server.py`) — dùng CHUNG pipeline với CLI và Streamlit (`backend/pipeline.py`).

Giữ nguyên các endpoint và tham số cũ:
  GET  /api/status, GET /api/jobs, POST /api/convert/paste, POST /api/convert,
  GET  /api/validate/{job_id}, GET /api/download/{job_id}/{docx|pdf|report}, GET /
Tham số mới (tuỳ chọn): answer_mode (source|student), show_solutions, figure_policy (auto|native|snapshot),
ocr (auto|off), ocr_lang, preprocess (engine|smart|auto; mặc định engine hoặc biến môi trường L2W_PREPROCESS).

Thay đổi so với bản cũ:
* năng lực dò THẬT (`capabilities.py`), không còn hằng True;
* tệp tải lên lưu nguyên bytes trong job riêng; nhận diện theo nội dung; ZIP giải nén có kiểm soát;
* `main_tex` sai → lỗi 409 kèm danh sách ứng viên (không âm thầm chọn tệp đầu tiên);
* PDF đọc trực tiếp (không tìm .tex trùng tên trong BASE_DIR/JOBS_DIR);
* dán mã và tệp TeX cùng đi qua pipeline; khi máy có tiền xử lý thông minh, pipeline gọi
  `convert_pasted_latex_unified` (smart_paste) cho cả hai nhánh;
* chuyển đổi chạy trong luồng riêng (`asyncio.to_thread`) để không chặn event loop;
* trạng thái: conversion_completed + quality_status (passed/review_required/failed); SKIP/NOT_MEASURED ≠ đạt;
* strict=True mà chưa `passed` → HTTP 422 (tệp vẫn tải được).
"""
from __future__ import annotations

import asyncio
import json
import sys
import time
from pathlib import Path
from typing import Any

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .capabilities import capability_summary, legacy_tools_view, probe_capabilities
from .job_workspace import MAX_UPLOAD_BYTES, JobWorkspace
from .pipeline import ConversionError, convert_bytes, convert_paste

app = FastAPI(
    title="LaTeX & PDF to Word Converter API",
    version="3.0.0",
    description="REST API chuyển PDF / LaTeX (.tex, .zip, dán mã) → Word (.docx), báo cáo khối native/ảnh/unresolved.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def add_no_cache_for_static(request, call_next):
    response = await call_next(request)
    if request.url.path.startswith("/static") or request.url.path == "/":
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


BASE_DIR = Path(__file__).resolve().parent.parent
CONVERTER_SCRIPT = BASE_DIR / "latex2word.py"
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

FRONTEND_DIR = BASE_DIR / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

_jobs: dict[str, dict[str, Any]] = {}


def _check_tools() -> dict[str, bool]:
    """(Giữ tên cũ) — nay lấy từ dò năng lực thật."""
    return legacy_tools_view(probe_capabilities())


@app.get("/api/status")
async def system_status():
    caps = await asyncio.to_thread(probe_capabilities)
    tools = legacy_tools_view(caps)
    d = caps["derived"]
    ready = d["tex_upload"]["available"] or d["pdf_reader"]["available"]
    return {
        "status": "ready" if ready else "degraded",
        "tools": tools,
        "capabilities": capability_summary(caps),
        "jobs_count": len(_jobs),
    }


@app.get("/api/jobs")
async def list_jobs():
    summaries = []
    for jid, j in sorted(_jobs.items(), key=lambda x: x[1].get("created", 0), reverse=True):
        summaries.append({"job_id": jid, "status": j.get("status"), "quality_status": j.get("quality_status"),
                          "filename": j.get("filename"), "mode": j.get("mode"), "created": j.get("created")})
    return {"jobs": summaries[:50]}


def _response_payload(report: dict[str, Any], filename: str, mode: int) -> dict[str, Any]:
    er = report.get("engine_report") or {}
    return {
        "job_id": report["job_id"],
        "status": "completed" if report.get("conversion_completed") else "failed",
        "conversion_completed": report.get("conversion_completed"),
        "quality_status": report.get("quality_status"),
        "overall": report.get("quality_status"),
        "filename": filename,
        "mode": mode,
        "ooxml_stats": er.get("ooxml_stats", {}),
        "engine_report": er,
        "layers": report.get("layers", []),
        "visual_pages_count": len(report.get("visual_pages", [])),
        "blocks_summary": report.get("blocks_summary", {}),
        "quality": report.get("quality", {}),
        "issues": report.get("issues", [])[:60],
        "source": {k: v for k, v in (report.get("source") or {}).items() if k in ("main", "selection", "candidates", "pdf", "pages", "page_types")},
        "preprocess": report.get("preprocess"),
        "input_preserved": report.get("input_preserved"),
    }


def _error_response(exc: ConversionError, filename: str) -> JSONResponse:
    content = {"job_id": exc.extra.get("job_id"), "status": "failed", "error": str(exc), "code": exc.code,
               "filename": filename}
    if exc.extra.get("candidates") is not None:
        content["candidates"] = exc.extra["candidates"]
    return JSONResponse(status_code=exc.http_status, content=content)


def _register(report: dict[str, Any], filename: str, mode: int) -> None:
    _jobs[report["job_id"]] = {
        "status": "completed" if report.get("conversion_completed") else "failed",
        "quality_status": report.get("quality_status"),
        "filename": filename, "mode": mode, "created": time.time(),
        "job_dir": str(JOBS_DIR / report["job_id"]),
        "docx_path": (report.get("output") or {}).get("docx"),
        "report_path": (report.get("output") or {}).get("report"),
        "qa_pdf": (report.get("output") or {}).get("qa_pdf"),
        "layers": report.get("layers", []),
    }


def _finish_response(report: dict[str, Any], filename: str, mode: int, strict: bool) -> JSONResponse:
    _register(report, filename, mode)
    payload = _response_payload(report, filename, mode)
    if strict and report.get("quality_status") != "passed":
        payload["strict_violation"] = True
        payload["error"] = ("Chế độ nghiêm ngặt: kết quả chưa đạt (" + str(report.get("quality_status")) + "). "
                            + "; ".join((report.get("quality") or {}).get("reasons", [])[:3]))
        return JSONResponse(status_code=422, content=payload)
    return JSONResponse(content=payload)


@app.post("/api/convert/paste")
async def convert_from_paste(
    latex_source: str = Form(...),
    mode: int = Form(2),
    dpi: int = Form(400),
    strict: bool = Form(False),
    qa_pdf: bool = Form(True),
    answer_mode: str = Form("source"),
    show_solutions: bool = Form(True),
    figure_policy: str = Form("auto"),
    preprocess: str = Form(""),
):
    try:
        report = await asyncio.to_thread(
            convert_paste, latex_source, jobs_root=JOBS_DIR, mode=mode, dpi=dpi, strict=strict, qa_pdf=qa_pdf,
            answer_mode=answer_mode, show_solutions=show_solutions, figure_policy=figure_policy,
            preprocess=preprocess or None)
    except ConversionError as exc:
        return _error_response(exc, "document.tex")
    return _finish_response(report, "document.tex", mode, strict)


@app.post("/api/convert")
async def convert_upload(
    file: UploadFile = File(...),
    main_tex: str = Form(""),
    mode: int = Form(2),
    dpi: int = Form(400),
    strict: bool = Form(False),
    qa_pdf: bool = Form(True),
    answer_mode: str = Form("source"),
    show_solutions: bool = Form(True),
    figure_policy: str = Form("auto"),
    ocr: str = Form("auto"),
    ocr_lang: str = Form(""),
    preprocess: str = Form(""),
):
    content = await file.read(MAX_UPLOAD_BYTES + 1)
    filename = file.filename or "upload"
    if len(content) > MAX_UPLOAD_BYTES:
        return JSONResponse(status_code=413, content={"status": "failed", "filename": filename,
                                                      "error": f"Tệp vượt giới hạn {MAX_UPLOAD_BYTES // (1024 * 1024)} MB."})
    try:
        report = await asyncio.to_thread(
            convert_bytes, filename, content, jobs_root=JOBS_DIR, main_file=main_tex or None, mode=mode, dpi=dpi,
            strict=strict, qa_pdf=qa_pdf, answer_mode=answer_mode, show_solutions=show_solutions,
            figure_policy=figure_policy, ocr=ocr, ocr_lang=ocr_lang or None, preprocess=preprocess or None)
    except ConversionError as exc:
        return _error_response(exc, filename)
    return _finish_response(report, filename, mode, strict)


@app.get("/api/validate/{job_id}")
async def validate_job(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        root = JobWorkspace.resolve_existing(JOBS_DIR, job_id)
        reports = sorted((root / "out").glob("*.report.json")) if root else []
        if not reports:
            raise HTTPException(404, "Job not found")
        rep = json.loads(reports[0].read_text(encoding="utf-8"))
        return {"job_id": job_id, "layers": rep.get("layers", []), "quality_status": rep.get("quality_status")}
    return {"job_id": job_id, "layers": job.get("layers", []), "quality_status": job.get("quality_status")}


def _inside(root: Path, p: Path) -> bool:
    try:
        p.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


@app.get("/api/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    root = JobWorkspace.resolve_existing(JOBS_DIR, job_id)
    if root is None:
        raise HTTPException(404, "Job not found")
    job = _jobs.get(job_id, {})
    out = root / "out"
    if file_type == "docx":
        target = Path(job["docx_path"]) if job.get("docx_path") else None
        if target is None or not target.exists():
            files = sorted(out.glob("*.docx")) if out.exists() else sorted(root.glob("*.docx"))
            target = files[0] if files else None
        if target is None or not _inside(root, target):
            raise HTTPException(404, "DOCX not found")
        return FileResponse(target, filename=target.name,
                            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if file_type == "report":
        target = Path(job["report_path"]) if job.get("report_path") else None
        if target is None or not target.exists():
            files = sorted(out.glob("*.report.json")) if out.exists() else []
            target = files[0] if files else None
        if target is None or not _inside(root, target):
            raise HTTPException(404, "Report not found")
        return FileResponse(target, filename=target.name, media_type="application/json")
    if file_type == "pdf":
        target = Path(job["qa_pdf"]) if job.get("qa_pdf") else None
        if target is None or not target.exists():
            files = sorted((out / "qa").glob("*.pdf")) if (out / "qa").exists() else []
            target = files[0] if files else None
        if target is None or not _inside(root, target):
            raise HTTPException(404, "QA PDF not found")
        return FileResponse(target, filename=target.name, media_type="application/pdf")
    raise HTTPException(400, "Invalid file_type. Use: docx, pdf, report")


@app.get("/")
async def serve_frontend():
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(index)
    return {"message": "Frontend not found. Place files in /frontend directory."}
