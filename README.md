# LaTeX → Word converter: starter tool

A **tested Python command-line starter** for turning a LaTeX `.tex` document into an editable Microsoft Word `.docx`. The core converter is **Pandoc**, with a preprocessing step that renders TikZ/PGFPlots diagrams into high-resolution images instead of discarding them.

> **Scope:** This is not a guarantee of lossless conversion for *every* LaTeX package or document. The tool is intentionally designed to report unsupported constructs rather than silently claim success. Passing automated checks still does **not** replace visual review in Microsoft Word.

## Installation

Install the following executables and ensure they are available on your command-line `PATH`:

1. **Python 3.9+** (no Python packages are required by the core script).
2. **Pandoc** — document and OMML math conversion: https://pandoc.org/installing.html
3. **A LaTeX distribution** containing `xelatex` (or `pdflatex`, `lualatex`), `standalone`, TikZ, and PGFPlots. On Windows, MiKTeX or TeX Live will work.
4. **Poppler utilities** including `pdftocairo` — turns compiled vector figures into 400-DPI PNG files for safe embedding into Word.
5. **Optional:** LibreOffice for a PDF preview of the DOCX (`--qa-pdf`). Use Word itself for final compatibility checks.

Verify installation in PowerShell / Terminal:

```shell
python --version
pandoc --version
xelatex --version
pdftocairo -v
```

## Usage

Place the `.tex` entry file and linked assets in a project directory, then run:

```shell
python latex2word.py example.tex -o converted.docx --strict --qa-pdf
```

If XeLaTeX is unavailable, switch to `--engine pdflatex`, but Unicode fonts and characters might require a different setup.

If your source includes TikZ plots with fine labels, increase raster resolution to 600 DPI:

```shell
python latex2word.py main.tex -o result.docx --engine xelatex --dpi 600 --qa-pdf
```

If you have a template with Word page styles, use a Pandoc `reference.docx`:

```shell
python latex2word.py main.tex -o result.docx --reference reference.docx --strict
```

The tool writes:

- `result.docx` — editable Word document
- `result.report.json` — counts and warnings from automated validation
- `result_qa/result.pdf` — optional PDF to inspect page breaks, diagrams and tables

## Expected behavior

| Source material | Destination | Editable in Word? |
| --- | --- | --- |
| Ordinary paragraphs, headings, numbers, basic formatting | Native text and styles | Yes |
| Common math (`$...$`, `equation`, etc.) | OMML Office Math | Yes, when parsed |
| Common `tabular` / `longtable` elements | Word tables | Yes, when parsed |
| Standard PNG/JPEG images | Embedded images | Image manipulation only |
| TikZ and PGFPlots | 400-DPI PNG after actual LaTeX rendering | **No** (the pixels are embedded) |
| Complicated custom macros, uncommon tables and environments | May require a custom handler | Not guaranteed |

If you need **editable Word charts** or Word shapes, use a second converter stage that parses source data and rebuilds the chart as an Office chart (Open XML), instead of embedding its image. That is a separate nontrivial feature.

## Current limitations / engineering notes

- TikZ/PGFPlots compilation extracts the preamble and uses the `standalone` LaTeX class. Some custom document classes or command definitions are incompatible with this approach.
- Simple braced `\\input{file}` and `\\include{file}` references are inlined recursively. Advanced include mechanisms, dynamic macro paths, commented instructions, and TeX verbatim may need a proper TeX parser.
- The TikZ matcher is intentionally scoped to the usual single-level `tikzpicture` outer environment. Exotic constructs and custom macros may need custom Lua filters or a parsing layer.
- Complex tables (`tabularray`, multi-page spans, nested tables, rotated text, merged cells) need specialized reconstruction.
- This starter does not yet resolve bibliography/cross-reference numbering, line-breaking decisions, all float placements, arbitrary `\\graphicspath`, or dynamically generated data plots. Pandoc supports citeproc; it can be added with `--citeproc` and `--bibliography`.
- Conversion to Word will not match TeX's line and page breaks exactly. Pandoc's `reference.docx` controls Word styles, not the precise TeX typesetting algorithm.
- A table/image/mathematics count is a **heuristic**, not proof that all content is correct.
- Treat untrusted LaTeX files as potentially unsafe. `-no-shell-escape` reduces one risk but is **not** a security sandbox. Run the conversion inside a locked-down container when supporting unknown uploads.

## Recommended production architecture

`Upload → Save assets → Validate / compile original PDF → Resolve includes and macros → Parse LaTeX → Route text & math & tables & graphics → Pandoc/OOXML → Structural audit → Render DOCX PDF → Visual + semantic comparison with original PDF → Deliver or produce exception report`

Suggested extensibility:

- **Web app:** Streamlit/FastAPI upload/download interface; isolated worker queue.
- **Editable math:** Pandoc to OMML plus a fallback pipeline with LaTeX/MathML conversion for unsupported formulas.
- **High-fidelity figures:** compile each graphic separately; SVG with compatibility fallback where supported, or a high-resolution PNG.
- **Full table fidelity:** custom LaTeX table AST parser and OOXML table writer.
- **Editable graphs:** parse PGFPlots/CSV points and construct native Office charts, otherwise use rendered figures.
- **Validation:** compare source and output text/numbers, OOXML object counts, and page-level image differences. Fail with a review report rather than asserting 100% correctness.

## Built-in smoke test

`example.tex` includes Unicode, numbers, native OMML math, one PGFPlots chart, and a LaTeX table. It was successfully converted using the script. `example.docx` is included for inspection.
