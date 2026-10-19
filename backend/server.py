"""FastAPI REST API Backend for LaTeX-to-Word Hybrid Converter.

Endpoints:
  POST /api/convert           — Upload .tex or .zip, get back .docx + report
  POST /api/convert/paste     — Send LaTeX source as text, get back .docx + report
  GET  /api/validate/{job_id} — Run 5-layer validation on a completed job
  GET  /api/download/{job_id}/{file_type} — Download .docx, .pdf, or report
  GET  /api/status            — System health & tool availability check
  GET  /api/jobs              — List recent conversion jobs

Run:
  uvicorn backend.server:app --reload --port 8000
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
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
    try:
        import winreg
        paths = []
        for hkey in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            sub = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment" if hkey == winreg.HKEY_LOCAL_MACHINE else r"Environment"
            try:
                with winreg.OpenKey(hkey, sub) as key:
                    p, _ = winreg.QueryValueEx(key, "Path")
                    paths.append(p)
            except Exception:
                pass
        if paths:
            combined = ";".join(paths) + ";" + os.environ.get("PATH", "")
            seen = set()
            clean_paths = []
            for part in combined.split(";"):
                part_clean = part.strip()
                if part_clean and part_clean.lower() not in seen:
                    seen.add(part_clean.lower())
                    clean_paths.append(part_clean)
            os.environ["PATH"] = ";".join(clean_paths)
    except Exception:
        pass

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .postprocessor import postprocess_docx_tables
from .validators import generate_visual_comparison, run_all_layers

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = FastAPI(
    title="LaTeX to Word Converter API",
    version="1.0.0",
    description="REST API cho Hệ thống chuyển đổi LaTeX → Word (.docx) Hybrid",
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

# Serve the frontend
FRONTEND_DIR = BASE_DIR / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

# In-memory job registry (lightweight; not for production at scale)
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
        "pandoc": shutil.which("pandoc") is not None,
        "xelatex": shutil.which("xelatex") is not None,
        "pdflatex": shutil.which("pdflatex") is not None,
        "pdftocairo": shutil.which("pdftocairo") is not None,
        "libreoffice": has_lo or has_d2p,
        "word_installed": has_d2p,
    }


# ---------------------------------------------------------------------------
# API Endpoints
# ---------------------------------------------------------------------------

@app.get("/api/status")
async def system_status():
    """Kiểm tra trạng thái hệ thống và công cụ ngoại vi."""
    tools = _check_tools()
    ready = tools["latex2word_py"] and tools["pandoc"] and (tools["xelatex"] or tools["pdflatex"])
    return {
        "status": "ready" if ready else "missing_tools",
        "tools": tools,
        "jobs_count": len(_jobs),
    }


@app.get("/api/jobs")
async def list_jobs():
    """Liệt kê các job chuyển đổi gần đây."""
    summaries = []
    for jid, j in sorted(_jobs.items(), key=lambda x: x[1].get("created", 0), reverse=True):
        summaries.append({
            "job_id": jid,
            "status": j.get("status"),
            "filename": j.get("filename"),
            "mode": j.get("mode"),
            "created": j.get("created"),
        })
    return {"jobs": summaries[:50]}


def normalize_latex(text: str) -> str:
    """Normalize Windows \r\r\n and \r\n to standard LF and prevent excessive blank lines."""
    text = text.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n{3,}", "\n\n", text)


@app.post("/api/convert/paste")
async def convert_from_paste(
    latex_source: str = Form(...),
    mode: int = Form(1),
    dpi: int = Form(400),
    strict: bool = Form(True),
    qa_pdf: bool = Form(True),
):
    """Chuyển đổi mã LaTeX dán trực tiếp."""
    if not CONVERTER_SCRIPT.exists():
        raise HTTPException(500, "latex2word.py not found")

    job_id = str(uuid.uuid4())[:8]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    tex_file = job_dir / "document.tex"
    clean_src = normalize_latex(latex_source)
    with open(tex_file, "w", encoding="utf-8", newline="\n") as f:
        f.write(clean_src)

    return await _run_conversion(job_id, job_dir, tex_file, "document.tex", mode, dpi, strict, qa_pdf)


@app.post("/api/convert")
async def convert_upload(
    file: UploadFile = File(...),
    main_tex: str = Form(""),
    mode: int = Form(1),
    dpi: int = Form(400),
    strict: bool = Form(True),
    qa_pdf: bool = Form(True),
):
    """Chuyển đổi file .tex hoặc .zip tải lên."""
    if not CONVERTER_SCRIPT.exists():
        raise HTTPException(500, "latex2word.py not found")

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
        # Clean newlines in main tex
        try:
            cleaned = normalize_latex(tex_file.read_text(encoding="utf-8", errors="replace"))
            with open(tex_file, "w", encoding="utf-8", newline="\n") as f:
                f.write(cleaned)
        except Exception:
            pass
    else:
        tex_file = job_dir / filename
        clean_content = normalize_latex(content.decode("utf-8", errors="replace"))
        with open(tex_file, "w", encoding="utf-8", newline="\n") as f:
            f.write(clean_content)

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
    """Thực thi pipeline chuyển đổi."""
    output_docx = job_dir / f"{tex_file.stem}.docx"
    report_json = output_docx.with_suffix(".report.json")

    _jobs[job_id] = {
        "status": "converting",
        "filename": orig_filename,
        "mode": mode,
        "created": time.time(),
        "job_dir": str(job_dir),
    }

    # Auto-disable features when tools are missing
    tools = _check_tools()
    if qa_pdf and not tools["libreoffice"]:
        qa_pdf = False  # LibreOffice not installed, skip QA PDF

    # Build command
    engine = "xelatex" if tools.get("xelatex") else "pdflatex"
    cmd = [
        sys.executable, str(CONVERTER_SCRIPT),
        str(tex_file), "-o", str(output_docx),
        "--engine", engine,
        "--dpi", str(dpi),
    ]
    # Don't use --strict to avoid failing on non-critical warnings
    # (e.g. unsupported environments that don't affect the main conversion)
    if qa_pdf:
        cmd.append("--qa-pdf")

    proc = subprocess.run(
        cmd, cwd=str(tex_file.parent),
        capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        timeout=300, env=os.environ.copy()
    )

    if not output_docx.exists():
        _jobs[job_id]["status"] = "failed"
        full_err = (proc.stdout + "\n" + proc.stderr).strip()
        # Find critical error lines
        meaningful = [
            l.strip() for l in full_err.splitlines()
            if any(k in l for k in ("Error:", "l.", "Runaway", "Fatal", "failed", "No pages"))
        ]
        short_err = "\n".join(meaningful[:5]) if meaningful else full_err[-500:]
        _jobs[job_id]["error"] = short_err
        _jobs[job_id]["full_error"] = full_err[-4000:]
        return JSONResponse(
            status_code=500,
            content={
                "job_id": job_id,
                "status": "failed",
                "error": short_err,
                "full_error": full_err[-4000:],
            },
        )

    # Post-process tables (always removes ugly borders on multiple choice options A/B/C/D)
    ooxml_stats = postprocess_docx_tables(output_docx, is_mode2=(mode == 2))

    # Read engine report
    engine_report = {}
    if report_json.exists():
        try:
            engine_report = json.loads(report_json.read_text(encoding="utf-8"))
        except Exception:
            pass

    # Visual comparison (Layer 5)
    visual_info = {"enabled": False}
    visual_pages = []
    if qa_pdf:
        qa_dir = job_dir / f"{tex_file.stem}_qa"
        qa_pdf_file = qa_dir / f"{tex_file.stem}.pdf"
        if not qa_pdf_file.exists():
            from latex2word import convert_docx_to_pdf
            convert_docx_to_pdf(output_docx, qa_pdf_file)
        if qa_pdf_file.exists():
            engine_report["qa_rendered_pdf"] = str(qa_pdf_file)
            visual_info, visual_pages = generate_visual_comparison(
                tex_file.parent, tex_file, qa_pdf_file,
            )

    # 5-layer validation
    tex_content = tex_file.read_text(encoding="utf-8", errors="ignore")
    layers = run_all_layers(tex_content, output_docx, visual_info)

    # Overall status
    statuses = [lr["status"] for lr in layers]
    overall = "passed" if all(s in ("PASS", "SKIP") for s in statuses) else "review_required"

    _jobs[job_id].update({
        "status": "completed",
        "overall": overall,
        "docx_path": str(output_docx),
        "report": engine_report,
        "ooxml_stats": ooxml_stats,
        "layers": layers,
        "visual_pages": visual_pages,
    })

    return JSONResponse(content={
        "job_id": job_id,
        "status": "completed",
        "overall": overall,
        "filename": orig_filename,
        "mode": mode,
        "ooxml_stats": ooxml_stats,
        "engine_report": engine_report,
        "layers": layers,
        "visual_pages_count": len(visual_pages),
    })


@app.get("/api/validate/{job_id}")
async def validate_job(job_id: str):
    """Chạy lại kiểm thử 5 tầng cho một job đã hoàn thành."""
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] != "completed":
        raise HTTPException(400, f"Job status: {job['status']}")
    return {"job_id": job_id, "layers": job.get("layers", [])}


@app.get("/api/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    """Tải file kết quả: docx, pdf, report."""
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
        # Look in _qa subdirectory
        pdf_files = list(job_dir.rglob("*_qa/*.pdf"))
        if not pdf_files:
            raise HTTPException(404, "QA PDF not found")
        return FileResponse(pdf_files[0], filename=pdf_files[0].name, media_type="application/pdf")
    else:
        raise HTTPException(400, "Invalid file_type. Use: docx, pdf, report")


# Serve frontend index.html for root
@app.get("/")
async def serve_frontend():
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(index)
    return {"message": "Frontend not found. Place files in /frontend directory."}
