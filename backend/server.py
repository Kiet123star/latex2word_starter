"""FastAPI REST API Backend for LaTeX-to-Word Unified Converter (`backend/server.py`).

Synchronized with `latex2word.py` and `app.py` to use `backend.core_converter`.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import time
import uuid
import zipfile
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

from .core_converter import convert_docx_to_pdf, convert_latex_file_unified
from .validators import generate_visual_comparison, run_all_layers

app = FastAPI(
    title="LaTeX to Word Converter API (Native DrawingML + Math Tables)",
    version="2.0.0",
    description="REST API cho Hệ thống chuyển đổi LaTeX → Word (.docx) với Native Word Shapes (DrawingML) & Bảng Toán học",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
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
    has_lo = bool(shutil.which("soffice") or shutil.which("libreoffice"))
    has_d2p = False
    if sys.platform == "win32":
        try:
            import docx2pdf

            has_d2p = True
        except Exception:
            pass
    return {
        "latex2word_py": CONVERTER_SCRIPT.exists(),
        "drawingml_engine": True,
        "math_tables_engine": True,
        "omml_math_engine": True,
        "pandoc": shutil.which("pandoc") is not None,
        "xelatex": shutil.which("xelatex") is not None,
        "pdflatex": shutil.which("pdflatex") is not None,
        "pdftocairo": shutil.which("pdftocairo") is not None,
        "libreoffice": has_lo or has_d2p,
        "word_installed": has_d2p,
    }


@app.get("/api/status")
async def system_status():
    """Kiểm tra trạng thái hệ thống (Nhân Native DrawingML + OMML luôn sẵn sàng)."""
    tools = _check_tools()
    return {
        "status": "ready",
        "tools": tools,
        "jobs_count": len(_jobs),
    }


@app.get("/api/jobs")
async def list_jobs():
    summaries = []
    for jid, j in sorted(_jobs.items(), key=lambda x: x[1].get("created", 0), reverse=True):
        summaries.append(
            {
                "job_id": jid,
                "status": j.get("status"),
                "filename": j.get("filename"),
                "mode": j.get("mode"),
                "created": j.get("created"),
            }
        )
    return {"jobs": summaries[:50]}


def normalize_latex(text: str) -> str:
    text = text.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n{3,}", "\n\n", text)


@app.post("/api/convert/paste")
async def convert_from_paste(
    latex_source: str = Form(...),
    mode: int = Form(2),
    dpi: int = Form(400),
    strict: bool = Form(False),
    qa_pdf: bool = Form(True),
):
    job_id = str(uuid.uuid4())[:8]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    tex_file = job_dir / "document.tex"
    clean_src = normalize_latex(latex_source)
    tex_file.write_text(clean_src, encoding="utf-8")

    return await _run_conversion(job_id, job_dir, tex_file, "document.tex", mode, dpi, strict, qa_pdf)


@app.post("/api/convert")
async def convert_upload(
    file: UploadFile = File(...),
    main_tex: str = Form(""),
    mode: int = Form(2),
    dpi: int = Form(400),
    strict: bool = Form(False),
    qa_pdf: bool = Form(True),
):
    job_id = str(uuid.uuid4())[:8]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    content = await file.read()
    filename = file.filename or "upload.tex"

    if filename.endswith(".zip"):
        zip_path = job_dir / filename
        zip_path.write_bytes(content)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(job_dir)
        candidates = list(job_dir.rglob("*.tex"))
        if not candidates:
            raise HTTPException(400, "No .tex file found in ZIP archive")
        if main_tex:
            matched = [p for p in candidates if p.name == main_tex]
            tex_file = matched[0] if matched else candidates[0]
        else:
            tex_file = candidates[0]
        cleaned = normalize_latex(tex_file.read_text(encoding="utf-8", errors="replace"))
        tex_file.write_text(cleaned, encoding="utf-8")
    else:
        tex_file = job_dir / filename
        clean_content = normalize_latex(content.decode("utf-8", errors="replace"))
        tex_file.write_text(clean_content, encoding="utf-8")

    return await _run_conversion(job_id, job_dir, tex_file, filename, mode, dpi, strict, qa_pdf)


async def _run_conversion(
    job_id: str,
    job_dir: Path,
    tex_file: Path,
    orig_filename: str,
    mode: int,
    dpi: int,
    strict: bool,
    qa_pdf: bool,
) -> JSONResponse:
    output_docx = job_dir / f"{tex_file.stem}.docx"

    _jobs[job_id] = {
        "status": "converting",
        "filename": orig_filename,
        "mode": mode,
        "created": time.time(),
        "job_dir": str(job_dir),
    }

    tools = _check_tools()
    if qa_pdf and not tools["libreoffice"]:
        qa_pdf = False

    try:
        engine_report = convert_latex_file_unified(
            tex_file,
            output_docx,
            dpi=dpi,
            strict=strict,
            qa_pdf=qa_pdf,
            mode=mode,
        )
    except Exception as exc:
        _jobs[job_id]["status"] = "failed"
        _jobs[job_id]["error"] = str(exc)
        return JSONResponse(
            status_code=500,
            content={"job_id": job_id, "status": "failed", "error": str(exc)},
        )

    ooxml_stats = engine_report.get("ooxml_stats", {})

    visual_info = {"enabled": False}
    visual_pages = []
    if qa_pdf:
        qa_dir = job_dir / f"{tex_file.stem}_qa"
        qa_pdf_file = qa_dir / f"{tex_file.stem}.pdf"
        if not qa_pdf_file.exists():
            convert_docx_to_pdf(output_docx, qa_pdf_file)
        if qa_pdf_file.exists():
            engine_report["qa_rendered_pdf"] = str(qa_pdf_file)
            visual_info, visual_pages = generate_visual_comparison(
                tex_file.parent, tex_file, qa_pdf_file
            )

    tex_content = tex_file.read_text(encoding="utf-8", errors="ignore")
    layers = run_all_layers(tex_content, output_docx, visual_info)
    statuses = [lr["status"] for lr in layers]
    overall = "passed" if all(s in ("PASS", "SKIP") for s in statuses) else "review_required"

    _jobs[job_id].update(
        {
            "status": "completed",
            "overall": overall,
            "docx_path": str(output_docx),
            "report": engine_report,
            "ooxml_stats": ooxml_stats,
            "layers": layers,
            "visual_pages": visual_pages,
        }
    )

    return JSONResponse(
        content={
            "job_id": job_id,
            "status": "completed",
            "overall": overall,
            "filename": orig_filename,
            "mode": mode,
            "ooxml_stats": ooxml_stats,
            "engine_report": engine_report,
            "layers": layers,
            "visual_pages_count": len(visual_pages),
        }
    )


@app.get("/api/validate/{job_id}")
async def validate_job(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] != "completed":
        raise HTTPException(400, f"Job status: {job['status']}")
    return {"job_id": job_id, "layers": job.get("layers", [])}


@app.get("/api/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    job_dir = Path(job["job_dir"])
    if file_type == "docx":
        files = list(job_dir.glob("*.docx"))
        if not files:
            raise HTTPException(404, "DOCX not found")
        return FileResponse(
            files[0],
            filename=files[0].name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    elif file_type == "report":
        files = list(job_dir.glob("*.report.json"))
        if not files:
            raise HTTPException(404, "Report not found")
        return FileResponse(files[0], filename=files[0].name, media_type="application/json")
    elif file_type == "pdf":
        pdf_files = list(job_dir.rglob("*_qa/*.pdf"))
        if not pdf_files:
            raise HTTPException(404, "QA PDF not found")
        return FileResponse(pdf_files[0], filename=pdf_files[0].name, media_type="application/pdf")
    else:
        raise HTTPException(400, "Invalid file_type. Use: docx, pdf, report")


@app.get("/")
async def serve_frontend():
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(index)
    return {"message": "Frontend not found. Place files in /frontend directory."}
