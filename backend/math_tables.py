"""Comprehensive Native Word Table Engine for Mathematics (`backend/math_tables.py`).

Handles all table types required in Mathematics documents and exams:
1. Variation & Sign Tables (`tkz-tab`: \\tkzTabInit, \\tkzTabLine, \\tkzTabVar, and hand-drawn TikZ BBT) ->
   Compact fixed-grid Native Word Table (<w:tbl>) with exact column widths, OMML equations (<m:oMath>),
   and native DrawingML variation arrows (<wps:wsp>) aligned directly between extrema values.
2. Smart Multiple-Choice Tables (\\choiceA..D, \\choice, \\boncot, \\haicot, \\motcot) -> Auto 1x4, 2x2, or 4x1
   borderless Word table with OMML math and bold/underlined correct answer key.
3. True/False 2025 Exam Tables (\\choiceTF) -> Structured 3-column table (Mệnh đề | Đúng | Sai).
4. Data, Frequency & Grouped Statistics Tables (`tabular`, `tabularx`, `longtable`, `array`) -> Full support for
   \\multicolumn (<w:gridSpan>), \\multirow (<w:vMerge>), \\diagbox (<w:tl2br>), \\hline, \\cline, and booktabs.
5. Side-by-Side Question & Figure Layout (\\immini{question}{tikz}) -> Borderless 2-column layout table.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

from .omml_math import append_rich_latex_text_to_paragraph

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def set_cell_margins(cell, top: int = 80, bottom: int = 80, left: int = 120, right: int = 120) -> None:
    """Set internal cell padding in dxa (1/20th of a point)."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcMar"))
    if existing is not None:
        tcPr.remove(existing)
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_borders(
    cell,
    *,
    top: tuple[str, int, str] | None = None,
    bottom: tuple[str, int, str] | None = None,
    left: tuple[str, int, str] | None = None,
    right: tuple[str, int, str] | None = None,
    tl2br: tuple[str, int, str] | None = None,
) -> None:
    """Set individual borders on a Word table cell. Each spec is (val, sz, color_hex)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    specs = {"top": top, "bottom": bottom, "left": left, "right": right, "tl2br": tl2br}
    for side, spec in specs.items():
        if spec is None:
            continue
        existing = tcBorders.find(qn(f"w:{side}"))
        if existing is not None:
            tcBorders.remove(existing)
        val, sz, col = spec
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), val)
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), col)
        tcBorders.append(b)


def set_table_borderless(table) -> None:
    """Make an entire Word table completely borderless."""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        return
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    tblPr.append(borders)


def set_table_width_pct(table, pct: int = 5000) -> None:
    """Set table width in fiftieths of a percent (5000 = 100%)."""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        return
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(pct))
    tblW.set(qn("w:type"), "pct")


def apply_fixed_table_grid(table, col_widths_dxa: list[int]) -> None:
    """Lock a Word table to exact fixed column widths in dxa (1/20th of a point) so Word AutoFit never shifts cells."""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        return
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(sum(col_widths_dxa)))
    tblW.set(qn("w:type"), "dxa")

    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")

    old_grid = table._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        table._tbl.remove(old_grid)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths_dxa:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)
    tblPr.addnext(tblGrid)

    for row in table.rows:
        for c_i, cell in enumerate(row.cells):
            if c_i < len(col_widths_dxa):
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(qn("w:w"), str(col_widths_dxa[c_i]))
                tcW.set(qn("w:type"), "dxa")


def keep_table_together(table) -> None:
    if len(table.rows) <= 6:
        for row in table.rows[:-1]:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.keep_with_next = True


def prevent_row_split(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


# ==============================================================================
# 1. BẢNG XÉT DẤU & BẢNG BIẾN THIÊN (TKZ-TAB & TIKZ BBT -> COMPACT FIXED WORD TABLE)
# ==============================================================================
def _create_variation_arrow_drawingml(
    going_down: bool, width_emu: int = 440000, height_emu: int = 270000, doc_id: int = 900
):
    """Create a native Word DrawingML diagonal arrow (<wps:wsp>) for a variation table interval cell."""
    flip_v = "" if going_down else ' flipV="1"'
    xml = f"""
    <w:r xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}">
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="{width_emu}" cy="{height_emu}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="{doc_id}" name="VarArrow_{doc_id}"/>
          <wp:cNvGraphicFramePr/>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp>
                <wps:cNvSpPr/>
                <wps:spPr>
                  <a:xfrm{flip_v}>
                    <a:off x="0" y="0"/>
                    <a:ext cx="{width_emu}" cy="{height_emu}"/>
                  </a:xfrm>
                  <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                  <a:ln w="12700">
                    <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
                    <a:tailEnd type="triangle" w="med" len="med"/>
                  </a:ln>
                </wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
    """
    return parse_xml(xml)


def is_tkz_tab_block(tikz_code: str) -> bool:
    """Return True if a tikzpicture block is a tkz-tab variation/sign table or a hand-drawn TikZ variation table."""
    if r"\tkzTabInit" in tikz_code or r"\tkzTabLine" in tikz_code or r"\tkzTabVar" in tikz_code:
        return True
    has_rect = "rectangle" in tikz_code
    has_inf = r"-\infty" in tikz_code and r"+\infty" in tikz_code
    has_xy = bool(re.search(r"\\node.*\{\$x\$\}", tikz_code)) and bool(
        re.search(r"\\node.*\{\$(?:y|f\(x\))\$\}", tikz_code)
    )
    return has_rect and has_inf and has_xy


def _read_braced_groups(s: str, start_idx: int, count: int) -> list[str]:
    """Extract `count` consecutive `{...}` braced groups starting from `start_idx`."""
    groups: list[str] = []
    pos = start_idx
    n = len(s)
    while pos < n and len(groups) < count:
        while pos < n and s[pos] != "{":
            if s[pos] == "[":
                depth_b = 1
                pos += 1
                while pos < n and depth_b > 0:
                    if s[pos] == "[":
                        depth_b += 1
                    elif s[pos] == "]":
                        depth_b -= 1
                    pos += 1
            else:
                pos += 1
        if pos >= n or s[pos] != "{":
            break
        pos += 1
        depth = 1
        g_start = pos
        while pos < n and depth > 0:
            if s[pos] == "{":
                depth += 1
            elif s[pos] == "}":
                depth -= 1
                if depth == 0:
                    groups.append(s[g_start:pos].strip())
                    pos += 1
                    break
            pos += 1
    return groups


def build_tkz_tab_word_table(doc, tikz_code: str, table_id: int = 500):
    """Convert a `tkz-tab` (or hand-drawn TikZ BBT) block into a compact, fixed-grid Native Editable Word Table."""
    if r"\tkzTabInit" not in tikz_code:
        return _build_handdrawn_tikz_bbt_table(doc, tikz_code, table_id=table_id)

    m_init = re.search(r"\\tkzTabInit", tikz_code)
    init_groups = _read_braced_groups(tikz_code, m_init.end(), 2) if m_init else []
    if len(init_groups) < 2:
        return None

    row_specs_raw = [r.strip() for r in init_groups[0].split(",") if r.strip()]
    row_headers: list[str] = [rs.split("/")[0].strip() for rs in row_specs_raw]

    x_vals = [x.strip() for x in init_groups[1].split(",") if x.strip()]
    num_pts = len(x_vals)
    num_grid_cols = max(1, 2 * num_pts - 1)
    total_cols = 1 + num_grid_cols

    line_tokens: list[str] = []
    m_line = re.search(r"\\tkzTabLine", tikz_code)
    if m_line:
        lg = _read_braced_groups(tikz_code, m_line.end(), 1)
        if lg:
            line_tokens = [t.strip() for t in lg[0].split(",")]

    var_tokens: list[str] = []
    m_var = re.search(r"\\tkzTabVar", tikz_code)
    if m_var:
        vg = _read_braced_groups(tikz_code, m_var.end(), 1)
        if vg:
            var_tokens = [t.strip() for t in vg[0].split(",") if t.strip()]

    num_rows = len(row_headers)
    table = doc.add_table(rows=num_rows, cols=total_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Compact fixed-grid widths in dxa (1 cm ≈ 567 dxa):
    # Col 0 (header): 760 dxa (~1.34 cm)
    # Odd cols (point/value columns): 520 dxa (~0.92 cm)
    # Even cols (interval/arrow columns): 780 dxa (~1.38 cm)
    col_widths_dxa = [760] + [520 if (c % 2 == 1) else 780 for c in range(1, total_cols)]
    apply_fixed_table_grid(table, col_widths_dxa)

    b_solid = ("single", 6, "0F172A")
    b_none = ("none", 0, "auto")

    for r_idx, row in enumerate(table.rows):
        prevent_row_split(row)
        for c_idx, cell in enumerate(row.cells):
            top_b = b_solid
            bot_b = b_solid
            left_b = b_solid if c_idx in (0, 1) else b_none
            right_b = b_solid if c_idx in (0, total_cols - 1) else b_none
            set_cell_borders(cell, top=top_b, bottom=bot_b, left=left_b, right=right_b)
            set_cell_margins(cell, top=40, bottom=40, left=15, right=15)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Populate Column 0 headers ($x$, $y'$, $y$)
    for r_idx, h_text in enumerate(row_headers):
        cell = table.cell(r_idx, 0)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        append_rich_latex_text_to_paragraph(p, h_text, font_size_pt=11.0)

    # Populate Row 0 ($x$ values at odd column indices 1, 3, 5, ...)
    for i, xv in enumerate(x_vals):
        col_i = 1 + 2 * i
        if col_i < total_cols:
            cell = table.cell(0, col_i)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            append_rich_latex_text_to_paragraph(p, xv, font_size_pt=11.0)

    # Populate Row 1 ($y'$ / $f'(x)$ signs and roots)
    if num_rows >= 2 and line_tokens:
        for c_offset in range(num_grid_cols):
            col_i = 1 + c_offset
            if c_offset >= len(line_tokens):
                break
            tok = line_tokens[c_offset].strip()
            if not tok:
                continue
            cell = table.cell(1, col_i)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if tok in ("d", r"\doubleline", "||"):
                append_rich_latex_text_to_paragraph(p, r"$\Vert$", font_size_pt=11.0)
            elif tok == "z":
                append_rich_latex_text_to_paragraph(p, "$0$", font_size_pt=11.0)
            else:
                math_tok = tok if tok.startswith("$") else f"${tok}$"
                append_rich_latex_text_to_paragraph(p, math_tok, font_size_pt=11.0)

    # Populate Variation Row ($y$ / $f(x)$) if present
    var_row_idx = num_rows - 1
    if var_tokens and var_row_idx >= 1:
        parsed_pts: list[tuple[str, str]] = []
        for vt in var_tokens:
            if "/" in vt:
                pos_code, val_str = vt.split("/", 1)
                parsed_pts.append((pos_code.strip(), val_str.strip()))
            else:
                parsed_pts.append(("", vt.strip()))

        for i, (pos_code, val_str) in enumerate(parsed_pts):
            col_i = 1 + 2 * i
            if col_i >= total_cols:
                break
            cell = table.cell(var_row_idx, col_i)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            is_top = pos_code.startswith("+")
            if is_top:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(22)
            else:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
                p.paragraph_format.space_before = Pt(22)
                p.paragraph_format.space_after = Pt(2)

            append_rich_latex_text_to_paragraph(p, val_str, font_size_pt=11.0)

            if i + 1 < len(parsed_pts) and (col_i + 1) < total_cols:
                next_pos_code, _ = parsed_pts[i + 1]
                next_is_top = next_pos_code.startswith("+")
                going_down = is_top and not next_is_top
                arrow_cell = table.cell(var_row_idx, col_i + 1)
                arrow_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                ap = arrow_cell.paragraphs[0]
                ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ap.paragraph_format.space_before = Pt(2)
                ap.paragraph_format.space_after = Pt(2)
                ap._p.append(
                    _create_variation_arrow_drawingml(
                        going_down=going_down,
                        width_emu=460000,
                        height_emu=280000,
                        doc_id=table_id * 10 + i,
                    )
                )

    keep_table_together(table)
    if hasattr(doc, "add_paragraph"):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(2)
    return table


def _build_handdrawn_tikz_bbt_table(doc, tikz_code: str, table_id: int = 550):
    """Convert a hand-drawn TikZ variation table (such as Câu 12 with vertical asymptote x=1) into a 7-column fixed-grid Word table."""
    has_double = "double" in tikz_code
    table = doc.add_table(rows=2, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Symmetric 7-column fixed grid:
    # Col 0 (header x/y): 680 dxa
    # Col 1, 2, 3 (left of asymptote): 480 (start val), 740 (arrow), 520 (end val before ||)
    # Col 4, 5, 6 (right of asymptote): 520 (start val after ||), 740 (arrow), 480 (end val)
    col_w = [680, 480, 740, 520, 520, 740, 480]
    apply_fixed_table_grid(table, col_w)

    b_solid = ("single", 6, "0F172A")
    b_none = ("none", 0, "auto")
    b_double = ("double", 6, "0F172A")

    for r_idx, row in enumerate(table.rows):
        prevent_row_split(row)
        for c_idx, cell in enumerate(row.cells):
            left_b = b_solid if c_idx in (0, 1) else (b_double if (c_idx == 4 and r_idx == 1 and has_double) else b_none)
            right_b = b_solid if c_idx in (0, 6) else (b_double if (c_idx == 3 and r_idx == 1 and has_double) else b_none)
            set_cell_borders(cell, top=b_solid, bottom=b_solid, left=left_b, right=right_b)
            set_cell_margins(cell, top=40, bottom=40, left=15, right=15)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Merge (0, 3) and (0, 4) (each 520 dxa) so x = 1 sits centered directly over the double border between col 3 & 4
    merged_x1 = table.cell(0, 3).merge(table.cell(0, 4))
    tcPr_m = merged_x1._tc.get_or_add_tcPr()
    tcW_m = tcPr_m.find(qn("w:tcW"))
    if tcW_m is not None:
        tcW_m.set(qn("w:w"), str(col_w[3] + col_w[4]))
    set_cell_borders(merged_x1, top=b_solid, bottom=b_solid, left=b_none, right=b_none)

    # Row 0: x | -\infty | (empty) | 1 (centered over ||) | (empty) | +\infty
    for col_i, txt in [(0, "$x$"), (1, r"$-\infty$"), (3, "$1$"), (6, r"$+\infty$")]:
        p = table.cell(0, col_i).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        append_rich_latex_text_to_paragraph(p, txt, font_size_pt=11.0)

    # Row 1: y | 2 (top) | \ arrow | -\infty (bottom) || +\infty (top) | \ arrow | 2 (bottom)
    p_y = table.cell(1, 0).paragraphs[0]
    p_y.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_y.paragraph_format.space_before = Pt(2)
    p_y.paragraph_format.space_after = Pt(2)
    append_rich_latex_text_to_paragraph(p_y, "$y$", font_size_pt=11.0)

    c1 = table.cell(1, 1)
    c1.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(22)
    append_rich_latex_text_to_paragraph(p1, "$2$", font_size_pt=11.0)

    c2 = table.cell(1, 2)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(2)
    p2._p.append(_create_variation_arrow_drawingml(going_down=True, width_emu=430000, height_emu=260000, doc_id=table_id + 1))

    c3 = table.cell(1, 3)
    c3.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    p3 = c3.paragraphs[0]
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p3.paragraph_format.space_before = Pt(22)
    p3.paragraph_format.space_after = Pt(2)
    append_rich_latex_text_to_paragraph(p3, r"$-\infty$", font_size_pt=10.5)

    c4 = table.cell(1, 4)
    c4.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p4 = c4.paragraphs[0]
    p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p4.paragraph_format.space_before = Pt(2)
    p4.paragraph_format.space_after = Pt(22)
    append_rich_latex_text_to_paragraph(p4, r"$+\infty$", font_size_pt=10.5)

    c5 = table.cell(1, 5)
    c5.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p5 = c5.paragraphs[0]
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(2)
    p5.paragraph_format.space_after = Pt(2)
    p5._p.append(_create_variation_arrow_drawingml(going_down=True, width_emu=430000, height_emu=260000, doc_id=table_id + 2))

    c6 = table.cell(1, 6)
    c6.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    p6 = c6.paragraphs[0]
    p6.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p6.paragraph_format.space_before = Pt(22)
    p6.paragraph_format.space_after = Pt(2)
    append_rich_latex_text_to_paragraph(p6, "$2$", font_size_pt=11.0)

    keep_table_together(table)
    if hasattr(doc, "add_paragraph"):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(2)
    return table


# ==============================================================================
# 2. BẢNG ĐÁP ÁN TRẮC NGHIỆM (4 PHƯƠNG ÁN & ĐÚNG/SAI THPT 2025)
# ==============================================================================
def _estimate_option_visual_length(opt_tex: str) -> int:
    s = re.sub(r"\\(?:dfrac|frac)\s*\{([^{}]*)\}\s*\{([^{}]*)\}", r"\1/\2", opt_tex)
    s = re.sub(r"\\[a-zA-Z]+", "", s)
    s = re.sub(r"[{}$_^]", "", s).strip()
    return len(s)


def build_multiple_choice_table(
    doc,
    options: list[str],
    correct_index: int | None = None,
    force_cols: int | None = None,
):
    """Build a clean, borderless Word table for 4 multiple-choice options (A, B, C, D)."""
    labels = ["A.", "B.", "C.", "D."]
    lengths = [_estimate_option_visual_length(o) for o in options]
    max_len = max(lengths) if lengths else 10

    if force_cols in (1, 2, 4):
        num_cols = force_cols
    elif max_len <= 14:
        num_cols = 4
    elif max_len <= 42:
        num_cols = 2
    else:
        num_cols = 1

    num_rows = (4 + num_cols - 1) // num_cols
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borderless(table)
    set_table_width_pct(table, 5000)

    col_pct = 5000 // num_cols
    for r_idx in range(num_rows):
        row = table.rows[r_idx]
        prevent_row_split(row)
        for c_idx in range(num_cols):
            opt_idx = r_idx * num_cols + c_idx
            cell = row.cells[c_idx]
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_pct))
            tcW.set(qn("w:type"), "pct")
            tcPr.append(tcW)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            set_cell_borders(
                cell,
                top=("none", 0, "auto"),
                bottom=("none", 0, "auto"),
                left=("none", 0, "auto"),
                right=("none", 0, "auto"),
            )
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            if opt_idx < len(options):
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15

                lbl_run = p.add_run(labels[opt_idx] + " ")
                lbl_run.bold = True
                lbl_run.font.size = Pt(12)
                if correct_index == opt_idx:
                    lbl_run.underline = True

                append_rich_latex_text_to_paragraph(p, options[opt_idx].strip(), font_size_pt=12.0)

    keep_table_together(table)
    return table


def build_true_false_table(doc, statements: list[tuple[str, str]]):
    """Build a 2025 Vietnamese High School Math True/False (Đúng/Sai) 4-statement table."""
    table = doc.add_table(rows=1 + len(statements), cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width_pct(table, 5000)

    headers = ["Mệnh đề", "Đúng", "Sai"]
    col_pcts = [3800, 600, 600]
    border_spec = ("single", 4, "94A3B8")

    h_row = table.rows[0]
    prevent_row_split(h_row)
    for c_i, h_txt in enumerate(headers):
        cell = h_row.cells[c_i]
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(col_pcts[c_i]))
        tcW.set(qn("w:type"), "pct")
        tcPr.append(tcW)
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>'))
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        set_cell_borders(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt)
        r.bold = True
        r.font.size = Pt(11)

    sub_labels = ["a)", "b)", "c)", "d)"]
    for idx, (stmt_tex, ans) in enumerate(statements):
        row = table.rows[idx + 1]
        prevent_row_split(row)
        for c_i in range(3):
            cell = row.cells[c_i]
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_pcts[c_i]))
            tcW.set(qn("w:type"), "pct")
            tcPr.append(tcW)
            set_cell_margins(cell, top=70, bottom=70, left=120, right=120)
            set_cell_borders(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p0 = row.cells[0].paragraphs[0]
        r_lbl = p0.add_run(f"{sub_labels[idx % 4]} ")
        r_lbl.bold = True
        r_lbl.font.size = Pt(11.5)
        append_rich_latex_text_to_paragraph(p0, stmt_tex, font_size_pt=11.5)

        ans_up = ans.strip().upper()
        if ans_up in ("Đ", "D", "TRUE", "1"):
            p1 = row.cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run("Đ")
            r1.bold = True
            r1.font.size = Pt(11.5)
        elif ans_up in ("S", "FALSE", "0"):
            p2 = row.cells[2].paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run("S")
            r2.bold = True
            r2.font.size = Pt(11.5)

    keep_table_together(table)
    return table


# ==============================================================================
# 3. BẢNG DỮ LIỆU, THỐNG KÊ TẦN SỐ, GHÉP NHÓM (TABULAR / MULTICOLUMN / MULTIROW)
# ==============================================================================
@dataclass
class _ParsedTableCell:
    text: str
    colspan: int = 1
    rowspan: int = 1
    align: str | None = None
    diagbox: tuple[str, str] | None = None


def _split_latex_table_cells(row_str: str) -> list[str]:
    cells: list[str] = []
    cur: list[str] = []
    depth = 0
    i = 0
    n = len(row_str)
    while i < n:
        ch = row_str[i]
        if ch == "\\":
            cur.append(row_str[i : i + 2])
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        if ch == "&" and depth == 0:
            cells.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
        i += 1
    cells.append("".join(cur).strip())
    return cells


def _split_latex_table_rows(body_str: str) -> list[str]:
    rows: list[str] = []
    cur: list[str] = []
    depth = 0
    i = 0
    n = len(body_str)
    while i < n:
        if body_str[i] == "{":
            depth += 1
            cur.append("{")
            i += 1
        elif body_str[i] == "}":
            depth = max(0, depth - 1)
            cur.append("}")
            i += 1
        elif body_str.startswith(r"\\", i) and depth == 0:
            rows.append("".join(cur).strip())
            cur = []
            i += 2
            while i < n and body_str[i] in " \t":
                i += 1
            if i < n and body_str[i] == "[":
                while i < n and body_str[i] != "]":
                    i += 1
                if i < n:
                    i += 1
        else:
            cur.append(body_str[i])
            i += 1
    tail = "".join(cur).strip()
    if tail:
        rows.append(tail)
    return rows


def build_latex_data_table(doc, tabular_tex: str):
    """Parse any LaTeX `tabular`, `tabularx`, or `longtable` environment into a styled Word table."""
    if "tabularx" in tabular_tex and re.search(r"(?:\\underline\{\\textbf\{[A-D]\.\}\}|[A-D]\.)\s*", tabular_tex):
        raw_cells = [
            c.strip()
            for r in _split_latex_table_rows(
                re.sub(r"^.*?\\begin\{tabularx\}\{[^{}]*\}\{[^{}]*\}", "", tabular_tex, flags=re.S).replace(
                    r"\end{tabularx}", ""
                )
            )
            for c in _split_latex_table_cells(r)
            if c.strip()
        ]
        if len(raw_cells) == 4:
            opts: list[str] = []
            correct_idx: int | None = None
            for idx, rc in enumerate(raw_cells):
                if r"\underline" in rc or r"\textbf" in rc:
                    correct_idx = idx
                cleaned_opt = re.sub(
                    r"^(?:\\underline\{\\textbf\{[A-D]\.\}\}|\\textbf\{\\underline\{[A-D]\.\}\}|[A-D]\.)\s*",
                    "",
                    rc,
                ).strip()
                opts.append(cleaned_opt)
            return build_multiple_choice_table(doc, opts, correct_index=correct_idx)

    m_begin = re.search(
        r"\\begin\{(?:tabular\*?|tabularx|longtable)\}(?:\{[^{}]*\})?\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}",
        tabular_tex,
        re.S,
    )
    col_spec_raw = m_begin.group(1) if m_begin else "ccc"
    body_start = m_begin.end() if m_begin else 0
    body_end = re.search(r"\\end\{(?:tabular\*?|tabularx|longtable)\}", tabular_tex)
    body = tabular_tex[body_start : body_end.start()] if body_end else tabular_tex[body_start:]

    has_vertical_bars = "|" in col_spec_raw
    col_aligns: list[str] = []
    for ch in re.sub(r"\{[^{}]*\}", "", col_spec_raw):
        if ch in "lcrpXm":
            col_aligns.append("center" if ch == "c" else ("right" if ch == "r" else "left"))

    raw_rows = _split_latex_table_rows(body)
    parsed_rows: list[list[_ParsedTableCell]] = []

    for r_str in raw_rows:
        r_clean = re.sub(
            r"\\(?:hline|toprule|midrule|bottomrule|cline\{[^{}]*\})\s*",
            "",
            r_str,
        ).strip()
        if not r_clean:
            continue
        raw_cells = _split_latex_table_cells(r_clean)
        row_cells: list[_ParsedTableCell] = []
        for c_str in raw_cells:
            c_clean = re.sub(
                r"\\(?:hline|toprule|midrule|bottomrule|cline\{[^{}]*\})\s*",
                "",
                c_str,
            ).strip()
            colspan = 1
            rowspan = 1
            c_align = None
            diag: tuple[str, str] | None = None

            m_mc = re.match(r"^\\multicolumn\s*\{\s*(\d+)\s*\}\s*\{([^{}]*)\}\s*\{(.*)\}\s*$", c_clean, re.S)
            if m_mc:
                colspan = max(1, int(m_mc.group(1)))
                spec_mc = m_mc.group(2)
                if "c" in spec_mc:
                    c_align = "center"
                elif "r" in spec_mc:
                    c_align = "right"
                elif "l" in spec_mc:
                    c_align = "left"
                c_clean = m_mc.group(3).strip()

            m_mr = re.match(r"^\\multirow\s*\{\s*(\d+)\s*\}\s*\{[^{}]*\}\s*\{(.*)\}\s*$", c_clean, re.S)
            if m_mr:
                rowspan = max(1, int(m_mr.group(1)))
                c_clean = m_mr.group(2).strip()

            m_db = re.match(r"^\\diagbox\s*\{([^{}]*)\}\s*\{([^{}]*)\}\s*$", c_clean, re.S)
            if m_db:
                diag = (m_db.group(1).strip(), m_db.group(2).strip())

            row_cells.append(
                _ParsedTableCell(text=c_clean, colspan=colspan, rowspan=rowspan, align=c_align, diagbox=diag)
            )
        if row_cells:
            parsed_rows.append(row_cells)

    if not parsed_rows:
        return None

    num_rows = len(parsed_rows)
    num_cols = max(sum(c.colspan for c in r) for r in parsed_rows)
    while len(col_aligns) < num_cols:
        col_aligns.append("center")

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width_pct(table, 4800)

    border_color = "475569" if has_vertical_bars else "94A3B8"
    b_single = ("single", 4, border_color)
    b_header_bottom = ("single", 8, "1E293B")
    b_none = ("none", 0, "auto")

    vmerge_remaining: dict[int, int] = {}

    for r_idx, r_cells in enumerate(parsed_rows):
        row = table.rows[r_idx]
        prevent_row_split(row)
        is_header = r_idx == 0

        col_cursor = 0
        for p_cell in r_cells:
            while col_cursor < num_cols and vmerge_remaining.get(col_cursor, 0) > 0:
                cov_cell = row.cells[col_cursor]
                cov_tcPr = cov_cell._tc.get_or_add_tcPr()
                cov_tcPr.append(OxmlElement("w:vMerge"))
                set_cell_borders(
                    cov_cell,
                    top=b_none,
                    bottom=b_single if r_idx == num_rows - 1 else b_none,
                    left=b_single if has_vertical_bars else b_none,
                    right=b_single if has_vertical_bars else b_none,
                )
                vmerge_remaining[col_cursor] -= 1
                col_cursor += 1

            if col_cursor >= num_cols:
                break

            target_cell = row.cells[col_cursor]
            if p_cell.colspan > 1 and col_cursor + p_cell.colspan - 1 < num_cols:
                end_cell = row.cells[col_cursor + p_cell.colspan - 1]
                target_cell = target_cell.merge(end_cell)

            if p_cell.rowspan > 1:
                tcPr = target_cell._tc.get_or_add_tcPr()
                vm = OxmlElement("w:vMerge")
                vm.set(qn("w:val"), "restart")
                tcPr.append(vm)
                for c_span in range(col_cursor, min(num_cols, col_cursor + p_cell.colspan)):
                    vmerge_remaining[c_span] = p_cell.rowspan - 1

            set_cell_margins(target_cell, top=80, bottom=80, left=120, right=120)
            target_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            if is_header:
                tcPr = target_cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

            set_cell_borders(
                target_cell,
                top=b_single if (r_idx == 0 or has_vertical_bars) else b_none,
                bottom=b_header_bottom
                if is_header
                else (b_single if (r_idx == num_rows - 1 or has_vertical_bars) else ("single", 4, "E2E8F0")),
                left=b_single if has_vertical_bars else b_none,
                right=b_single if has_vertical_bars else b_none,
                tl2br=b_single if p_cell.diagbox else None,
            )

            p = target_cell.paragraphs[0]
            align_mode = p_cell.align or col_aligns[col_cursor]
            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
                if align_mode == "center"
                else (WD_ALIGN_PARAGRAPH.RIGHT if align_mode == "right" else WD_ALIGN_PARAGRAPH.LEFT)
            )
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if p_cell.diagbox:
                left_lbl, right_lbl = p_cell.diagbox
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                append_rich_latex_text_to_paragraph(p, right_lbl, default_bold=True, font_size_pt=10.5)
                p_bot = target_cell.add_paragraph()
                p_bot.alignment = WD_ALIGN_PARAGRAPH.LEFT
                append_rich_latex_text_to_paragraph(p_bot, left_lbl, default_bold=True, font_size_pt=10.5)
            elif p_cell.text:
                append_rich_latex_text_to_paragraph(
                    p,
                    p_cell.text,
                    default_bold=is_header,
                    font_size_pt=11.0,
                )

            col_cursor += p_cell.colspan

        while col_cursor < num_cols and vmerge_remaining.get(col_cursor, 0) > 0:
            cov_cell = row.cells[col_cursor]
            cov_tcPr = cov_cell._tc.get_or_add_tcPr()
            cov_tcPr.append(OxmlElement("w:vMerge"))
            set_cell_borders(
                cov_cell,
                top=b_none,
                bottom=b_single if r_idx == num_rows - 1 else b_none,
                left=b_single if has_vertical_bars else b_none,
                right=b_single if has_vertical_bars else b_none,
            )
            vmerge_remaining[col_cursor] -= 1
            col_cursor += 1

    keep_table_together(table)
    return table


# ==============================================================================
# 4. BẢNG DÀN TRANG SONG SONG ĐỀ BÀI + HÌNH VẼ (\IMMINI / MINIPAGE)
# ==============================================================================
def build_immini_layout_table(doc, left_content_fn, right_drawing_el):
    """Create a borderless 2-column table placing question text on the left (62%) and DrawingML figure on the right (38%)."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borderless(table)
    set_table_width_pct(table, 5000)

    row = table.rows[0]
    prevent_row_split(row)
    left_cell, right_cell = row.cells[0], row.cells[1]

    for cell, pct in ((left_cell, 3100), (right_cell, 1900)):
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(pct))
        tcW.set(qn("w:type"), "pct")
        tcPr.append(tcW)
        set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
        set_cell_borders(
            cell,
            top=("none", 0, "auto"),
            bottom=("none", 0, "auto"),
            left=("none", 0, "auto"),
            right=("none", 0, "auto"),
        )

    left_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    left_content_fn(left_cell)

    rp = right_cell.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp._p.append(right_drawing_el)
    keep_table_together(table)
    return table
