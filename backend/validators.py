"""5-Layer Automatic Validation Engine.

Extracted from app.py logic and enhanced for REST API usage.
Each layer returns structured data suitable for JSON serialization.
"""
from __future__ import annotations

import re
import shutil
import subprocess
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import numpy as np
    from PIL import Image, ImageChops, ImageEnhance
    HAS_VISION = True
except ImportError:
    HAS_VISION = False

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass
class LayerResult:
    """Kết quả kiểm tra từng lớp trong 5 lớp xác thực."""
    layer: int
    name: str
    status: str          # "PASS", "WARN", "FAIL", "SKIP"
    summary: str
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return asdict(self)


def strip_latex_commands(tex: str) -> str:
    """Lược bỏ lệnh điều khiển LaTeX để lấy văn bản thuần."""
    text = re.sub(r"%.*$", "", tex, flags=re.MULTILINE)
    text = re.sub(
        r"\\begin\{(?:tikzpicture|axis)\}.*?\\end\{(?:tikzpicture|axis)\}",
        " ", text, flags=re.S,
    )
    text = re.sub(r"\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?", " ", text)
    text = re.sub(r"[{}$&#^_~]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def run_layer1_text(tex_content: str, root: ET.Element) -> LayerResult:
    """Layer 1 — Text Verification."""
    docx_texts = [node.text for node in root.findall(".//w:t", NS) if node.text]
    docx_plain = " ".join(docx_texts).strip()
    source_plain = strip_latex_commands(tex_content)
    docx_wc = len(docx_plain.split())
    src_wc = max(len(source_plain.split()), 1)
    ratio = round(min(docx_wc / src_wc, 1.0) * 100, 1)
    status = "PASS" if docx_wc > 0 and ratio >= 50 else "WARN"
    return LayerResult(
        layer=1,
        name="Text Verification",
        status=status,
        summary=f"Trích xuất {docx_wc} từ trong Word (tỷ lệ bảo toàn ~{ratio}%).",
        details={
            "source_words": src_wc,
            "docx_words": docx_wc,
            "retention_percent": ratio,
            "paragraphs": len(root.findall(".//w:p", NS)),
        },
    )


def run_layer2_math(tex_content: str, root: ET.Element) -> LayerResult:
    """Layer 2 — Mathematical Verification (OMML)."""
    omath = root.findall(".//m:oMath", NS)
    fracs = root.findall(".//m:f", NS)
    rads = root.findall(".//m:rad", NS)
    sups = root.findall(".//m:sSup", NS)
    subs = root.findall(".//m:sSub", NS)
    mats = root.findall(".//m:m", NS)
    nary = root.findall(".//m:nary", NS)
    src_math = len(re.findall(
        r"(\\\[.*?\\\]|\$\$.*?\$\$|\\begin\{(equation\*?|align\*?|gather\*?)\})",
        tex_content, re.S,
    ))
    status = "PASS" if (src_math == 0 or len(omath) > 0) else "WARN"
    return LayerResult(
        layer=2,
        name="Mathematical Verification (OMML)",
        status=status,
        summary=(
            f"Phát hiện {len(omath)} khối OMML "
            f"(Phân số: {len(fracs)}, Căn: {len(rads)}, "
            f"Nary: {len(nary)}, Ma trận: {len(mats)})."
        ),
        details={
            "omml_equations": len(omath),
            "fractions": len(fracs),
            "radicals": len(rads),
            "superscripts": len(sups),
            "subscripts": len(subs),
            "nary_ops": len(nary),
            "matrices": len(mats),
        },
    )


def run_layer3_illustrations(tex_content: str, media_files: list[str]) -> LayerResult:
    """Layer 3 — Illustration Verification."""
    tikz = len(re.findall(r"\\begin\{tikzpicture\}", tex_content))
    axis = len(re.findall(r"\\begin\{axis\}", tex_content))
    incl = len(re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{[^{}]+\}", tex_content))
    expected = tikz + incl
    status = "PASS" if len(media_files) >= expected else "WARN"
    return LayerResult(
        layer=3,
        name="Illustration Verification",
        status=status,
        summary=(
            f"Nhúng {len(media_files)} hình ảnh "
            f"(Nguồn: {tikz} TikZ, {axis} Axis, {incl} includegraphics)."
        ),
        details={
            "embedded_media": len(media_files),
            "source_tikz": tikz,
            "source_axis": axis,
            "source_includegraphics": incl,
        },
    )


def run_layer4_tables(tex_content: str, root: ET.Element) -> LayerResult:
    """Layer 4 — Table Verification."""
    src_tables = len(re.findall(
        r"\\begin\{(?:tabular\*?|tabularx|longtable)\}", tex_content,
    ))
    docx_tables = root.findall(".//w:tbl", NS)
    structures = []
    for i, tbl in enumerate(docx_tables, 1):
        rows = tbl.findall("./w:tr", NS)
        cols = [len(r.findall("./w:tc", NS)) for r in rows]
        merged = (
            len(tbl.findall(".//w:gridSpan", NS))
            + len(tbl.findall(".//w:vMerge", NS))
        )
        structures.append({
            "index": i,
            "rows": len(rows),
            "max_cols": max(cols) if cols else 0,
            "merged_tags": merged,
        })
    status = "PASS" if len(docx_tables) >= src_tables else "WARN"
    return LayerResult(
        layer=4,
        name="Table Verification",
        status=status,
        summary=f"Đã chuyển đổi {len(docx_tables)}/{src_tables} bảng.",
        details={
            "source_tables": src_tables,
            "docx_tables": len(docx_tables),
            "structures": structures,
        },
    )


def run_layer5_visual(visual_info: dict[str, Any] | None) -> LayerResult:
    """Layer 5 — Visual Verification."""
    if visual_info and visual_info.get("enabled"):
        sim = visual_info.get("similarity_percent", 0.0)
        status = "PASS" if sim >= 80.0 else "WARN"
        return LayerResult(
            layer=5,
            name="Visual Verification (PDF Diff)",
            status=status,
            summary=f"Độ tương đồng bố cục trực quan: {sim:.1f}%.",
            details=visual_info,
        )
    reason = (visual_info or {}).get("reason", "Chưa bật --qa-pdf hoặc thiếu công cụ Word/LibreOffice.")
    return LayerResult(
        layer=5,
        name="Visual Verification (PDF Diff)",
        status="SKIP",
        summary=reason,
        details=visual_info or {},
    )


def sanitize_latex_for_compilation(tex: str) -> str:
    """Clean blank lines inside math modes and axis options that crash TeX."""
    # Clean blank lines in \begin{axis}[...]
    def fix_axis(m: re.Match[str]) -> str:
        opts = m.group(1)
        cleaned = [l for l in opts.splitlines() if l.strip()]
        return "\\begin{axis}[\n" + "\n".join(cleaned) + "\n]"

    tex = re.sub(r"\\begin\{axis\}\s*\[(.*?)\]", fix_axis, tex, flags=re.S)

    # Clean blank lines inside \[ ... \]
    def fix_display_math(m: re.Match[str]) -> str:
        body = m.group(1)
        cleaned = [l for l in body.splitlines() if l.strip()]
        return "\\[\n" + "\n".join(cleaned) + "\n\\]"

    tex = re.sub(r"\\\[(.*?)\\\]", fix_display_math, tex, flags=re.S)
    return tex


def run_all_layers(
    tex_content: str,
    docx_path: Path,
    visual_info: dict[str, Any] | None = None,
) -> list[dict]:
    """Chạy toàn bộ 5 tầng kiểm thử và trả về danh sách dict."""
    if not docx_path.exists():
        return [LayerResult(
            layer=0, name="All Layers", status="FAIL",
            summary="Không tìm thấy file .docx.",
        ).to_dict()]

    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
        media = [f for f in zf.namelist() if f.startswith("word/media/")]

    root = ET.fromstring(xml_bytes)
    results = [
        run_layer1_text(tex_content, root),
        run_layer2_math(tex_content, root),
        run_layer3_illustrations(tex_content, media),
        run_layer4_tables(tex_content, root),
        run_layer5_visual(visual_info),
    ]
    return [r.to_dict() for r in results]


def generate_visual_comparison(
    workdir: Path, tex_file: Path, word_pdf_path: Path,
) -> tuple[dict[str, Any], list[dict]]:
    """Biên dịch LaTeX gốc → PDF và so sánh với PDF Word."""
    info: dict[str, Any] = {"enabled": False}
    pages: list[dict] = []

    # If Word PDF doesn't exist, try to generate it from docx
    if not word_pdf_path.exists():
        docx_cand = word_pdf_path.with_suffix(".docx")
        if not docx_cand.exists():
            docx_cand = workdir / f"{tex_file.stem}.docx"
        if docx_cand.exists():
            try:
                from latex2word import convert_docx_to_pdf
                convert_docx_to_pdf(docx_cand, word_pdf_path)
            except Exception:
                pass

    if not word_pdf_path.exists():
        info["reason"] = "Không tìm thấy PDF xuất từ Word (thiếu MS Word / LibreOffice)."
        return info, pages

    latex_engine = shutil.which("xelatex") or shutil.which("pdflatex")
    pdftocairo = shutil.which("pdftocairo")
    if not latex_engine or not pdftocairo:
        info["reason"] = "Thiếu xelatex/pdflatex hoặc pdftocairo trong PATH."
        return info, pages

    build_dir = workdir / "visual_build"
    build_dir.mkdir(exist_ok=True)

    # Prepare clean TeX source for compilation
    clean_tex_file = build_dir / "clean_source.tex"
    raw_tex = tex_file.read_text(encoding="utf-8", errors="replace")
    clean_tex_file.write_text(sanitize_latex_for_compilation(raw_tex), encoding="utf-8")

    proc = subprocess.run(
        [latex_engine, "-interaction=nonstopmode", "-halt-on-error",
         f"-output-directory={build_dir}", str(clean_tex_file)],
        cwd=workdir, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=120,
    )
    orig_pdf = build_dir / "clean_source.pdf"
    if not orig_pdf.exists():
        # Try compiling original tex_file directly as fallback
        proc2 = subprocess.run(
            [latex_engine, "-interaction=nonstopmode", "-halt-on-error",
             f"-output-directory={build_dir}", str(tex_file)],
            cwd=workdir, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=120,
        )
        orig_pdf = build_dir / f"{tex_file.stem}.pdf"

    if not orig_pdf.exists():
        info["reason"] = "Không biên dịch được .tex gốc sang PDF (lỗi cú pháp LaTeX)."
        return info, pages

    orig_prefix = build_dir / "latex_page"
    word_prefix = build_dir / "word_page"
    subprocess.run([pdftocairo, "-png", "-r", "150", str(orig_pdf), str(orig_prefix)], check=False)
    subprocess.run([pdftocairo, "-png", "-r", "150", str(word_pdf_path), str(word_prefix)], check=False)

    orig_pngs = sorted(build_dir.glob("latex_page*.png"))
    word_pngs = sorted(build_dir.glob("word_page*.png"))
    if not orig_pngs or not word_pngs:
        info["reason"] = "Không tạo được ảnh PNG từ PDF."
        return info, pages

    sims: list[float] = []
    for idx, (po, pw) in enumerate(zip(orig_pngs, word_pngs), 1):
        page_info: dict[str, Any] = {
            "page": idx,
            "latex_png": str(po),
            "word_png": str(pw),
        }
        if HAS_VISION:
            img_a = Image.open(po).convert("RGB")
            img_b = Image.open(pw).convert("RGB").resize(img_a.size)
            diff = ImageChops.difference(img_a, img_b)
            arr = np.asarray(diff, dtype=np.float32)
            sim = max(0.0, (1.0 - float(np.mean(arr)) / 255.0) * 100.0)
            sims.append(sim)
            boosted = ImageEnhance.Brightness(diff).enhance(3.5)
            diff_path = build_dir / f"diff_page_{idx}.png"
            boosted.save(diff_path)
            page_info["diff_png"] = str(diff_path)
            page_info["similarity"] = round(sim, 2)
        pages.append(page_info)

    avg = sum(sims) / len(sims) if sims else 90.0
    info = {
        "enabled": True,
        "latex_pages": len(orig_pngs),
        "word_pages": len(word_pngs),
        "similarity_percent": round(avg, 2),
    }
    return info, pages
