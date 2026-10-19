#!/usr/bin/env python3
"""LaTeX to Word starter with TikZ/PGFPlots rendering and fail-visible validation.

Requirements: Python 3.9+, Pandoc, (Xe)LaTeX, pdftocairo.
This is a starter, NOT a lossless converter for every LaTeX package.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from xml.etree import ElementTree as ET

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Recognizes common single-layer graphic environments. Nested environments
# such as `axis` INSIDE tikzpicture are preserved inside their outer match.
TIKZ = re.compile(r"\\begin\{tikzpicture\}(?:\[[^\]]*\])?.*?\\end\{tikzpicture\}", re.S)
AXIS = re.compile(r"\\begin\{axis\}(?:\[[^\]]*\])?.*?\\end\{axis\}", re.S)
TABLE_ENV = re.compile(r"\\begin\{(?:tabular\*?|tabularx|longtable)\}")
IMAGE_CMD = re.compile(r"\\includegraphics(?:\[[^\]]*\])?\{([^{}]+)\}")
UNSUPPORTED = re.compile(r"\\begin\{(pspicture|forest|circuitikz|algorithm2e|algorithm|algorithmic|xymatrix|asy|asydef|gnuplot|picture)\}")
INCLUDE = re.compile(r"\\(?:input|include)\s*\{([^{}]+)\}")
UNKNOWN_INCLUDE = re.compile(r"\\(?:input|include|subfile|import|subimport)\b")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "m": "http://schemas.openxmlformats.org/officeDocument/2006/math", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def clean_newlines(s: str) -> str:
    """Normalize line endings to standard LF."""
    return s.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")


def sanitize_pgfplots_options(tex: str) -> str:
    """Remove blank lines inside [...] of \\begin{axis}[...] to prevent runaway \\par errors."""
    def fix_axis(m: re.Match[str]) -> str:
        opts = m.group(1)
        cleaned_lines = [l for l in opts.splitlines() if l.strip()]
        return f"\\begin{{axis}}[\n" + "\n".join(cleaned_lines) + "\n]"
    return re.sub(r"\\begin\{axis\}\s*\[(.*?)\]", fix_axis, tex, flags=re.S)


def execute(cmd: list[str], *, cwd: Path, timeout: int = 180) -> subprocess.CompletedProcess[str]:
    try:
        proc = subprocess.run(
            cmd, cwd=cwd, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout, check=False
        )
    except subprocess.TimeoutExpired as err:
        raise RuntimeError(f"Timed out running: {' '.join(cmd)}") from err
    if proc.returncode:
        excerpt = (proc.stderr + "\n" + proc.stdout)[-6000:]
        raise RuntimeError(f"Command failed ({proc.returncode}): {' '.join(cmd)}\n{excerpt}")
    return proc


def expand_tex_includes(main: Path, stack: tuple[Path, ...] = ()) -> tuple[str, int]:
    """Resolve simple literal braced input/include commands recursively.

    This intentionally rejects dynamic paths rather than losing their content.
    Parent project-relative paths are supported; unusual TeX include mechanisms
    must be handled before running this starter.
    """
    main = main.resolve()
    if main in stack:
        raise RuntimeError("Circular TeX include chain: " + " -> ".join(map(str, stack + (main,))))
    contents = clean_newlines(main.read_text(encoding="utf-8-sig"))
    found = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal found
        value = match.group(1).strip()
        name = Path(value)
        if name.suffix == "":
            name = name.with_suffix(".tex")
        file = (main.parent / name).resolve()
        if not file.is_file():
            raise RuntimeError(f"TeX include not found: {value} (referenced by {main})")
        child, recursive_count = expand_tex_includes(file, stack + (main,))
        found += 1 + recursive_count
        return "\n" + child + "\n"

    # Exclude comments in simple text preprocessing would require a TeX lexer.
    # This is intended for controlled project sources, not arbitrary macros.
    contents = INCLUDE.sub(replace, contents)
    return contents, found


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """Convert DOCX to PDF using LibreOffice or docx2pdf (MS Word COM)."""
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if lo:
        try:
            subprocess.run([lo, "--headless", "--convert-to", "pdf", "--outdir",
                            str(pdf_path.parent), str(docx_path)],
                           capture_output=True, timeout=90)
            expected = pdf_path.parent / (docx_path.stem + ".pdf")
            if expected.exists():
                if expected != pdf_path:
                    shutil.move(str(expected), str(pdf_path))
                return True
        except Exception:
            pass

    if sys.platform == "win32":
        try:
            import docx2pdf
            docx2pdf.convert(str(docx_path), str(pdf_path))
            if pdf_path.exists():
                return True
        except Exception:
            pass
    return False


def latex_figure_document(src: str, original: str) -> str:
    """Keep source preamble macros, replacing its class with standalone.

    Removes geometry, setspace, and fullpage to prevent clipping diagram bounding boxes.
    Uses 15pt border so axis labels, ticks, and table borders are never cropped.
    """
    original = clean_newlines(original)
    src = clean_newlines(src)
    src = sanitize_pgfplots_options(src)
    if "\\begin{document}" in original:
        pre = original.split("\\begin{document}", 1)[0]
        pre, num = re.subn(r"\\documentclass(?:\[[^\]]*\])?\{[^{}]+\}",
                           r"\\documentclass[border=15pt]{standalone}", pre, count=1)
        if not num:
            pre = "\\documentclass[border=15pt]{standalone}\n" + pre
    else:
        pre = "\\documentclass[border=15pt]{standalone}\n"

    # Remove packages and macros that break standalone canvas sizing and cause (0,0) clipping
    pre = re.sub(r"\\usepackage(?:\[[^\]]*\])?\{(?:geometry|setspace|fullpage|fancyhdr)\}", "", pre)
    pre = re.sub(r"\\(?:geometry|setstretch)\{[^{}]*\}", "", pre)
    # Ensure tikz & pgfplots are loaded
    pre += "\n\\usepackage{tikz}\n\\usepackage{pgfplots}\n"
    return pre + "\n\\begin{document}\n" + src + "\n\\end{document}\n"


def create_fallback_image(out_path: Path, label: str, err_msg: str = "") -> Path:
    """Generate a clean fallback PNG if a TikZ figure fails to compile."""
    try:
        from PIL import Image, ImageDraw
        img = Image.new("RGB", (600, 160), color="#f8fafc")
        d = ImageDraw.Draw(img)
        d.rectangle([(2, 2), (597, 157)], outline="#cbd5e1", width=2)
        d.text((20, 25), f"[Hình vẽ TikZ: {label}]", fill="#1e293b")
        d.text((20, 65), "Không thể biên dịch hình vẽ TikZ.", fill="#dc2626")
        if err_msg:
            first_err = err_msg.strip().splitlines()[-1][:75]
            d.text((20, 105), f"Lỗi: {first_err}", fill="#64748b")
        img.save(out_path, "PNG")
    except Exception:
        # Minimal 1x1 PNG fallback
        out_path.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15c4\x00\x00\x00\rIDATx\x9cc`\x00\x00\x00\x02\x00\x01H\xaf\xa4q\x00\x00\x00\x00IEND\xaeB`\x82"
        )
    return out_path


def render_tikz(body: str, *, original: str, name: str, tmp: Path,
                project_root: Path, engine: str, resolution: int) -> Path:
    tex_file = tmp / (name + ".tex")
    fig_code = latex_figure_document(body, original)
    with open(tex_file, "w", encoding="utf-8", newline="\n") as f:
        f.write(fig_code)
    execute([engine, "-no-shell-escape", "-halt-on-error", "-interaction=nonstopmode",
             "-file-line-error", "-output-directory", str(tmp), str(tex_file)],
            cwd=project_root)
    src_pdf = tmp / (name + ".pdf")
    dest_without_ext = tmp / name
    execute(["pdftocairo", "-png", "-r", str(resolution), "-singlefile", 
             str(src_pdf), str(dest_without_ext)], cwd=project_root)
    out_file = tmp / (name + ".png")
    if not out_file.exists() or out_file.stat().st_size == 0:
        raise RuntimeError(f"Figure rendering produced no output: {name}")
    return out_file


def docx_counts(file: Path) -> dict:
    with zipfile.ZipFile(file, "r") as doc:
        bad = doc.testzip()
        if bad:
            raise RuntimeError(f"Broken DOCX archive: {bad}")
        xml = ET.fromstring(doc.read("word/document.xml"))
        image_paths = [n for n in doc.namelist() if n.startswith("word/media/") and not n.endswith("/")]
        return {
            "word_tables": len(xml.findall(".//w:tbl", NS)),
            "equation_nodes": len(xml.findall(".//m:oMath", NS)),
            "drawing_references": len(xml.findall(".//a:blip", NS)),
            "embedded_media_files": len(image_paths),
        }


def convert(tex: Path, output: Path, *, engine: str, dpi: int, 
            reference: Path | None, strict: bool, qa_pdf: bool) -> dict:
    required = ["pandoc", engine, "pdftocairo"]
    for binary in required:
        if not shutil.which(binary):
            raise RuntimeError(f"Required tool not found in PATH: {binary}")
    if qa_pdf:
        has_lo = bool(shutil.which("soffice") or shutil.which("libreoffice"))
        has_d2p = False
        if sys.platform == "win32":
            try:
                import docx2pdf
                has_d2p = True
            except Exception:
                pass
        if not (has_lo or has_d2p):
            raise RuntimeError("QA PDF requires LibreOffice or Microsoft Word (docx2pdf)")
    tex, output = tex.resolve(), output.resolve()
    if not tex.is_file():
        raise RuntimeError(f"Input does not exist: {tex}")
    output.parent.mkdir(parents=True, exist_ok=True)
    original, included_count = expand_tex_includes(tex)
    original = clean_newlines(original)
    diagnostics: list[str] = []
    unsupported = sorted(set(UNSUPPORTED.findall(original)))
    if unsupported:
        diagnostics.append(f"Unsupported environment(s) need manual handler: {', '.join(unsupported)}")
    if UNKNOWN_INCLUDE.search(original):
        diagnostics.append("Unexpanded TeX include/import command found. This may cause lost content.")
    report = {"input": str(tex), "output": str(output), "rendered_figures": 0, "expanded_tex_includes": included_count,
              "warnings": diagnostics, "status": "started"}

    with tempfile.TemporaryDirectory(prefix="latex2word-") as temp:
        tmp = Path(temp)
        fig_dir = tmp / "generated_figures"
        fig_dir.mkdir()
        processed = original
        source_graphics = len(IMAGE_CMD.findall(original))

        # Collect all figure tasks for parallel compilation
        figure_tasks: list[tuple[str, str]] = []

        def tikz_collector(m: re.Match[str]) -> str:
            idx = len(figure_tasks) + 1
            name = f"graphic_{idx:03d}"
            figure_tasks.append((name, m.group(0)))
            return "\\includegraphics{generated_figures/" + name + ".png}"

        processed = TIKZ.sub(tikz_collector, processed)

        def axis_collector(m: re.Match[str]) -> str:
            idx = len(figure_tasks) + 1
            name = f"graphic_{idx:03d}"
            wrapped = "\\begin{tikzpicture}\n" + m.group(0) + "\n\\end{tikzpicture}"
            figure_tasks.append((name, wrapped))
            return "\\includegraphics{generated_figures/" + name + ".png}"

        processed = AXIS.sub(axis_collector, processed)

        # Render figures in parallel using ThreadPoolExecutor
        def render_worker(task: tuple[str, str]) -> tuple[str, bool, str | None]:
            fig_name, fig_body = task
            try:
                render_tikz(fig_body, original=original, name=fig_name, tmp=fig_dir,
                            project_root=tex.parent, engine=engine, resolution=dpi)
                return (fig_name, True, None)
            except Exception as err:
                fallback_path = fig_dir / f"{fig_name}.png"
                create_fallback_image(fallback_path, fig_name, str(err))
                return (fig_name, False, str(err))

        if figure_tasks:
            workers = min(14, os.cpu_count() or 8)
            with ThreadPoolExecutor(max_workers=workers) as executor:
                results = list(executor.map(render_worker, figure_tasks))

            for fig_name, success, err in results:
                report["rendered_figures"] += 1
                if not success:
                    first_err = str(err).splitlines()[-1] if err else "error"
                    diagnostics.append(f"Figure {fig_name} warning: {first_err[:120]}")

        # Never claim full coverage of nested custom environments.
        remaining = ["tikzpicture", "axis"]
        for env in remaining:
            if f"\\begin{{{env}}}" in processed:
                diagnostics.append(f"Unconverted {env} environment remains")
        preprocessed_tex = tmp / "processed.tex"
        with open(preprocessed_tex, "w", encoding="utf-8", newline="\n") as f:
            f.write(processed)
        resource_path = str(tmp) + (";" if sys.platform == "win32" else ":") + str(tex.parent)
        cmd = ["pandoc", str(preprocessed_tex), "--from=latex+raw_tex", "--to=docx",
               "--standalone", "--resource-path", resource_path,
               "--output", str(output), "--log", str(tmp / "pandoc_log.json")]
        if reference:
            cmd.extend(["--reference-doc", str(reference.resolve())])
        # In strict mode catch warnings, but also validate remaining output below.
        if strict:
            cmd.append("--fail-if-warnings")
        try:
            conversion = execute(cmd, cwd=tex.parent)
            if conversion.stderr.strip():
                diagnostics.append("Pandoc messages: " + conversion.stderr.strip()[-2000:])
        except RuntimeError:
            if output.is_file():
                output.unlink()
            raise
        counts = docx_counts(output)
        report.update(counts)
        expected_tables = len(TABLE_ENV.findall(original))
        report["input_table_environments"] = expected_tables
        report["input_includegraphics"] = source_graphics
        report["expected_total_image_commands"] = source_graphics + report["rendered_figures"]
        if expected_tables > counts["word_tables"]:
            diagnostics.append(f"Possible missing tables: source has {expected_tables}; Word has {counts['word_tables']}")
        if counts["drawing_references"] < source_graphics + report["rendered_figures"]:
            diagnostics.append(f"Possible missing graphics: expected at least {source_graphics + report['rendered_figures']}, Word has {counts['drawing_references']}")
        if re.search(r"(?<!\\)\$(?!\$).+?(?<!\\)\$|\\begin\{(?:equation|align|gather)\}", original, re.S) and counts["equation_nodes"] == 0:
            diagnostics.append("Math input detected, but no editable OMML equation detected")
        if qa_pdf:
            render_dir = output.parent / (output.stem + "_qa")
            render_dir.mkdir(exist_ok=True)
            target_pdf = render_dir / (output.stem + ".pdf")
            ok = convert_docx_to_pdf(output, target_pdf)
            if ok and target_pdf.exists():
                report["qa_rendered_pdf"] = str(target_pdf)
            else:
                diagnostics.append("QA PDF generation failed via Word/LibreOffice")
        report["status"] = "review_required" if diagnostics else "automated_checks_passed"
        report_location = output.with_suffix(".report.json")
        report_location.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        if strict and diagnostics:
            raise RuntimeError("Strict validation failed: " + " | ".join(diagnostics) + 
                               f". See report {report_location}")
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("tex", type=Path, help="Main LaTeX .tex file")
    ap.add_argument("-o", "--output", type=Path, help="Destination .docx file")
    ap.add_argument("--engine", choices=["pdflatex", "xelatex", "lualatex"], default="xelatex")
    ap.add_argument("--dpi", type=int, default=400, help="Diagram raster DPI (default: 400)")
    ap.add_argument("--reference", type=Path, help="Custom Pandoc reference.docx")
    ap.add_argument("--strict", action="store_true", help="Fail on Pandoc warnings or validation issues")
    ap.add_argument("--qa-pdf", action="store_true", help="Render Word as PDF for review (LibreOffice)")
    args = ap.parse_args()
    if args.dpi < 150 or args.dpi > 1200:
        ap.error("--dpi must be between 150 and 1200")
    out = args.output or args.tex.with_suffix(".docx")
    try:
        result = convert(args.tex, out, engine=args.engine, dpi=args.dpi,
                         reference=args.reference, strict=args.strict, qa_pdf=args.qa_pdf)
    except (RuntimeError, OSError, UnicodeError, zipfile.BadZipFile) as error:
        print("ERROR:", error, file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
