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

import contextvars

#: Chế độ hiển thị đáp án do pipeline đặt cho ngữ cảnh chuyển đổi hiện tại:
#:   "source"  — giữ đúng cách trình bày của nguồn (mặc định, như trước);
#:   "student" — KHÔNG làm lộ đáp án (không gạch chân/in đậm phương án đúng, để trống cột Đ/S).
ANSWER_MODE: contextvars.ContextVar[str] = contextvars.ContextVar("l2w_answer_mode", default="source")
_EVENTS: contextvars.ContextVar[list | None] = contextvars.ContextVar("l2w_table_events", default=None)


def start_table_events() -> list:
    """Bật ghi sự kiện dựng bảng (đáp án có nguồn/không rõ, bảng chưa dựng được…) cho ngữ cảnh hiện tại."""
    ev: list = []
    _EVENTS.set(ev)
    return ev


def _event(kind: str, **data) -> None:
    ev = _EVENTS.get()
    if ev is not None:
        ev.append({"kind": kind, **data})


def _tag_role(table, role: str, **extra) -> None:
    try:
        from .ooxml_tools import set_table_role
        set_table_role(table._tbl, role, **extra)
    except Exception:
        pass


def _set_tcw(cell, value: int, wtype: str = "pct") -> None:
    """Đặt w:tcW duy nhất (python-docx đã tạo sẵn một w:tcW — không được thêm cái thứ hai)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(value))
    tcW.set(qn("w:type"), wtype)
    tcPr.insert(0, tcW) if tcPr.find(qn("w:cnfStyle")) is None else tcPr.find(qn("w:cnfStyle")).addnext(tcW)


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


def _split_top_tkz(s: str, sep: str) -> list[str]:
    """Tách theo `sep` ở mức ngoài cùng: không tách trong {…}, […] hay $…$ (vd. $1,5$; $\frac{1}{2}$)."""
    out: list[str] = []
    cur: list[str] = []
    depth = 0
    in_math = False
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            cur.append(s[i:i + 2])
            i += 2
            continue
        if c == "$":
            in_math = not in_math
        elif c in "{[":
            depth += 1
        elif c in "}]":
            depth = max(0, depth - 1)
        if c == sep and depth == 0 and not in_math:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(c)
        i += 1
    out.append("".join(cur))
    return out


_TKZ_VAR_CODE_RE = re.compile(r"^\s*([+-])?\s*(D|R)?\s*([+-])?\s*$")
_TKZ_UNSUPPORTED_CMDS = ("tkzTabIma", "tkzTabVal", "tkzTabSlope", "tkzTabTan", "tkzTabSetup")


def parse_tkz_tab(tikz_code: str) -> dict:
    """Đọc `\tkzTabInit` + mọi `\tkzTabLine`/`\tkzTabVar` THEO THỨ TỰ (nhiều dòng xét dấu/biến thiên).
    Trả về {"rows", "xs", "seq", "unsupported": [lý do…]}. Không đoán giá trị không có trong nguồn."""
    info: dict = {"rows": [], "xs": [], "seq": [], "unsupported": []}
    m_init = re.search(r"\\tkzTabInit", tikz_code)
    if not m_init:
        info["unsupported"].append("không có \\tkzTabInit")
        return info
    groups = _read_braced_groups(tikz_code, m_init.end(), 2)
    if len(groups) < 2:
        info["unsupported"].append("\\tkzTabInit thiếu đối số")
        return info
    for spec in _split_top_tkz(groups[0], ","):
        if not spec.strip():
            continue
        parts = _split_top_tkz(spec, "/")
        label = parts[0].strip()
        try:
            height = float(parts[1]) if len(parts) > 1 and parts[1].strip() else 1.0
        except ValueError:
            height = 1.0
        info["rows"].append((label, height))
    info["xs"] = [x.strip() for x in _split_top_tkz(groups[1], ",")]
    for name in _TKZ_UNSUPPORTED_CMDS:
        if "\\" + name in tikz_code:
            info["unsupported"].append(f"lệnh \\{name} chưa dựng native")
    if re.search(r"\\(?:draw|node|fill|path)\b", tikz_code):
        info["unsupported"].append("có lệnh TikZ vẽ thêm trong bảng (\\draw/\\node…)")
    for m in re.finditer(r"\\tkzTab(Line|Var)\b", tikz_code):
        g = _read_braced_groups(tikz_code, m.end(), 1)
        if not g:
            info["unsupported"].append(f"\\tkzTab{m.group(1)} thiếu đối số")
            continue
        toks = [t.strip() for t in _split_top_tkz(g[0], ",")]
        if m.group(1) == "Line":
            for t in toks:
                if t in ("", "+", "-", "z", "t", "d", "h") or t.startswith("$") or re.fullmatch(r"[0-9.,]+", t):
                    continue
                info["unsupported"].append(f"ký hiệu dòng dấu chưa hỗ trợ: {t!r}")
            info["seq"].append(("line", toks))
        else:
            pts = []
            for t in toks:
                parts = _split_top_tkz(t, "/")
                code = parts[0].strip()
                vals = [v.strip() for v in parts[1:]]
                mm = _TKZ_VAR_CODE_RE.match(code)
                if not mm:
                    info["unsupported"].append(f"mã biến thiên chưa hỗ trợ: {code!r}")
                    pts.append({"code": code, "unsupported": True})
                    continue
                left, kind, right = mm.group(1), mm.group(2), mm.group(3)
                if kind == "D":
                    # -D+ : trái dưới / phải trên; D- : chỉ bên phải; -D : chỉ bên trái
                    lv = vals[0] if left and vals else None
                    rv = (vals[1] if left and len(vals) > 1 else (vals[0] if not left and vals else None)) if right else None
                    pts.append({"code": code, "double": True, "left": left, "right": right, "lv": lv, "rv": rv})
                elif kind == "R":
                    pts.append({"code": code, "skip": True})
                else:
                    lvl = left or right
                    if not lvl:
                        info["unsupported"].append(f"mã biến thiên rỗng ở điểm {len(pts) + 1}")
                        pts.append({"code": code, "unsupported": True})
                        continue
                    v = vals[0] if vals else ""
                    pts.append({"code": code, "left": lvl, "right": lvl, "lv": v, "rv": v})
            info["seq"].append(("var", pts))
    n_rows_cmd = len(info["seq"])
    if n_rows_cmd != max(0, len(info["rows"]) - 1):
        info["unsupported"].append(
            f"số dòng khai báo ({len(info['rows']) - 1}) khác số lệnh \\tkzTabLine/\\tkzTabVar ({n_rows_cmd})")
    return info


def tkz_tab_support(tikz_code: str) -> dict:
    """Pipeline dùng để quyết định: dựng native hay chụp ảnh TeX (table_as_image)."""
    if r"\tkzTabInit" not in tikz_code:
        return {"native": False, "reasons": ["bảng biến thiên vẽ tay bằng TikZ — không đọc được cấu trúc an toàn"]}
    info = parse_tkz_tab(tikz_code)
    return {"native": not info["unsupported"], "reasons": info["unsupported"]}


def _tkz_level_arrow(start: str | None, end: str | None) -> str | None:
    if start is None or end is None:
        return None
    if start == "+" and end == "-":
        return "down"
    if start == "-" and end == "+":
        return "up"
    return "flat"


def build_tkz_tab_word_table(doc, tikz_code: str, table_id: int = 500):
    """Dựng bảng xét dấu / biến thiên `tkz-tab` thành bảng Word sửa được.

    Hỗ trợ: nhiều \\tkzTabLine/\\tkzTabVar theo thứ tự; ký hiệu dòng dấu  + - z t d h;  mã biến thiên
    + - R  và họ D (-D+, +D-, -D-, +D+, D-, D+, -D, +D) với tiệm cận (vạch kép) và giá trị hai phía.
    Mã khác (H, C…, \\tkzTabIma…) được ghi `unsupported`; pipeline sẽ chụp ảnh TeX thay vì dựng sai."""
    if r"\tkzTabInit" not in tikz_code:
        return _build_handdrawn_tikz_bbt_table(doc, tikz_code, table_id=table_id)
    info = parse_tkz_tab(tikz_code)
    if not info["rows"] or not info["xs"]:
        return _bbt_unresolved_notice(doc, tikz_code, "không đọc được \\tkzTabInit")
    if info["unsupported"]:
        _event("bbt_partial", reasons=info["unsupported"][:10])

    xs = info["xs"]
    n_pts = len(xs)
    n_grid = max(1, 2 * n_pts - 1)
    total_cols = 1 + n_grid
    n_rows = len(info["rows"])
    table = doc.add_table(rows=n_rows, cols=total_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # độ rộng cột theo nội dung thật (ước lượng em), không cố định
    def _w_em(tex: str | None) -> float:
        return _option_width_em(tex or "") if tex else 0.0

    pt_w = [max(1.2, _w_em(x)) for x in xs]
    has_double_col = [False] * n_pts
    for kind, toks in info["seq"]:
        if kind == "var":
            for i, p in enumerate(toks[:n_pts]):
                if p.get("double"):
                    has_double_col[i] = True
                    pt_w[i] = max(pt_w[i], _w_em(p.get("lv")) + _w_em(p.get("rv")) + 1.2)
                else:
                    pt_w[i] = max(pt_w[i], _w_em(p.get("lv")))
        else:
            for c, t in enumerate(toks[: n_grid]):
                if c % 2 == 0 and t not in ("", "+", "-", "z", "t", "d", "h"):
                    pt_w[c // 2] = max(pt_w[c // 2], _w_em(t))
    head_w = max(1.6, max(_w_em(lbl) for lbl, _h in info["rows"]))
    em_dxa = 240  # 12pt
    col_widths = [int(head_w * em_dxa) + 160]
    for c in range(n_grid):
        if c % 2 == 0:
            col_widths.append(max(480, int(pt_w[c // 2] * em_dxa * 0.9) + 120))
        else:
            col_widths.append(780)
    apply_fixed_table_grid(table, col_widths)

    b_solid = ("single", 6, "0F172A")
    b_none = ("none", 0, "auto")
    for r_idx, row in enumerate(table.rows):
        prevent_row_split(row)
        for c_idx, cell in enumerate(row.cells):
            left_b = b_solid if c_idx in (0, 1) else b_none
            right_b = b_solid if c_idx in (0, total_cols - 1) else b_none
            set_cell_borders(cell, top=b_solid, bottom=b_solid, left=left_b, right=right_b)
            set_cell_margins(cell, top=40, bottom=40, left=15, right=15)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def _para(cell, text: str | None, align=WD_ALIGN_PARAGRAPH.CENTER, first: bool = True, size: float = 11.0):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        if text:
            append_rich_latex_text_to_paragraph(p, text, font_size_pt=size)
        return p

    for r_idx, (label, _h) in enumerate(info["rows"]):
        _para(table.cell(r_idx, 0), label)
    for i, xv in enumerate(xs):
        if xv:
            _para(table.cell(0, 1 + 2 * i), xv)

    for k, (kind, toks) in enumerate(info["seq"]):
        r_idx = k + 1
        if r_idx >= n_rows:
            break
        if kind == "line":
            for c_off, tok in enumerate(toks[:n_grid]):
                cell = table.cell(r_idx, 1 + c_off)
                if tok in ("", "t"):
                    continue
                if tok == "d":
                    set_cell_borders(cell, left=b_solid, right=b_solid)
                    _para(cell, r"$\Vert$")
                elif tok == "z":
                    _para(cell, "$0$")
                elif tok == "h":
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="BFBFBF"/>'))
                elif tok in ("+", "-"):
                    _para(cell, f"${tok}$")
                else:
                    _para(cell, tok if "$" in tok else f"${tok}$")
            continue
        # dòng biến thiên
        pts = toks[:n_pts]
        row = table.rows[r_idx]
        tr_pr = row._tr.get_or_add_trPr()
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), str(int(240 * 2.4)))
        h.set(qn("w:hRule"), "atLeast")
        tr_pr.append(h)
        for i, p in enumerate(pts):
            if p.get("unsupported") or p.get("skip"):
                continue
            cell = table.cell(r_idx, 1 + 2 * i)
            if p.get("double"):
                set_cell_borders(cell, left=b_solid, right=b_solid)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                top_l = p["lv"] if p.get("left") == "+" else None
                top_r = p["rv"] if p.get("right") == "+" else None
                bot_l = p["lv"] if p.get("left") == "-" else None
                bot_r = p["rv"] if p.get("right") == "-" else None
                tp = _para(cell, None, WD_ALIGN_PARAGRAPH.LEFT, first=True, size=10.5)
                _two_side_line(tp, top_l, top_r, col_widths[1 + 2 * i])
                mid = _para(cell, None, WD_ALIGN_PARAGRAPH.CENTER, first=False)
                mid.paragraph_format.space_before = Pt(8)
                mid.paragraph_format.space_after = Pt(8)
                bp = _para(cell, None, WD_ALIGN_PARAGRAPH.LEFT, first=False, size=10.5)
                _two_side_line(bp, bot_l, bot_r, col_widths[1 + 2 * i])
            else:
                top = p["left"] == "+"
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP if top else WD_ALIGN_VERTICAL.BOTTOM
                _para(cell, p.get("lv") or "")
        # mũi tên giữa hai điểm liên tiếp (bỏ qua R: nối tới điểm có giá trị gần nhất)
        def right_level(idx: int):
            j = idx
            while j >= 0 and j < len(pts) and pts[j].get("skip"):
                j -= 1
            if j < 0 or j >= len(pts) or pts[j].get("unsupported"):
                return None
            return pts[j].get("right")

        def left_level(idx: int):
            j = idx
            while j < len(pts) and pts[j].get("skip"):
                j += 1
            if j >= len(pts) or pts[j].get("unsupported"):
                return None
            return pts[j].get("left")

        for i in range(len(pts) - 1):
            direction = _tkz_level_arrow(right_level(i), left_level(i + 1))
            if direction is None:
                continue
            ac = table.cell(r_idx, 2 + 2 * i)
            ac.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            ap = ac.paragraphs[0]
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if direction == "flat":
                append_rich_latex_text_to_paragraph(ap, "$\\to$", font_size_pt=11.0)
            else:
                ap._p.append(_create_variation_arrow_drawingml(
                    going_down=(direction == "down"), width_emu=460000, height_emu=280000,
                    doc_id=table_id * 100 + r_idx * 10 + i))
    keep_table_together(table)
    _tag_role(table, "variation", source="tkz-tab", partial="1" if info["unsupported"] else None)
    if hasattr(doc, "add_paragraph"):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(2)
    return table


def _two_side_line(p, left_tex: str | None, right_tex: str | None, cell_w_dxa: int) -> None:
    """Một dòng trong ô tiệm cận: giá trị trái sát trái, giá trị phải sát phải (tab phải)."""
    if right_tex:
        tabs = p._p.get_or_add_pPr()
        tabs_el = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:pos"), str(max(200, cell_w_dxa - 60)))
        tabs_el.append(tab)
        tabs.append(tabs_el)
    if left_tex:
        append_rich_latex_text_to_paragraph(p, left_tex, font_size_pt=10.5)
    if right_tex:
        p.add_run("\t")
        append_rich_latex_text_to_paragraph(p, right_tex, font_size_pt=10.5)


def _bbt_unresolved_notice(doc, tikz_code: str, reason: str):
    """Không dựng được bảng biến thiên một cách chắc chắn → đoạn thông báo hiển thị (KHÔNG bịa bảng)."""
    _event("bbt_unresolved", reason=reason, source=tikz_code[:4000])
    if not hasattr(doc, "add_paragraph"):
        return None
    p = doc.add_paragraph()
    r = p.add_run(f"[L2W: CHƯA DỰNG ĐƯỢC BẢNG BIẾN THIÊN — {reason}. Cần TeX để chụp ảnh bảng gốc; "
                  "mã nguồn được giữ trong báo cáo.]")
    r.bold = True
    return None


def _build_handdrawn_tikz_bbt_table(doc, tikz_code: str, table_id: int = 550):
    """Bảng biến thiên vẽ tay bằng TikZ (\\draw/\\node, không có tkz-tab).

    Bản cũ dựng một bảng CỐ ĐỊNH (x: −∞, 1, +∞; y: 2 ↘ −∞ ‖ +∞ ↘ 2) cho MỌI bảng vẽ tay → bịa nội dung.
    Nay: không đoán cấu trúc; để lại thông báo hiển thị + ghi sự kiện `bbt_unresolved`. Pipeline sẽ chụp
    ảnh TeX cho khối này trước khi tới đây khi có trình biên dịch."""
    return _bbt_unresolved_notice(doc, tikz_code, "bảng vẽ tay bằng lệnh TikZ, không đọc được cấu trúc an toàn")


# ==============================================================================
# 2. BẢNG ĐÁP ÁN TRẮC NGHIỆM (4 PHƯƠNG ÁN & ĐÚNG/SAI THPT 2025)
# ==============================================================================
def _estimate_option_visual_length(opt_tex: str) -> int:
    """(Giữ cho tương thích) số ký tự hiển thị ước lượng."""
    return int(round(_option_width_em(opt_tex) * 2))


def _option_width_em(tex: str) -> float:
    """Ước lượng bề rộng hiển thị (đơn vị em) của nội dung LaTeX: phân số lấy max(tử, mẫu) chứ không cộng
    dồn, căn thêm dấu căn, chỉ số trên/dưới thu nhỏ. Là ước lượng heuristic để chọn số cột."""
    s = tex or ""

    def width(t: str) -> float:
        total = 0.0
        i = 0
        while i < len(t):
            c = t[i]
            if t.startswith("\\", i):
                m = re.match(r"\\([A-Za-z]+)", t[i:])
                if not m:
                    total += 0.5
                    i += 2
                    continue
                name = m.group(1)
                j = i + m.end()
                if name in ("frac", "dfrac", "tfrac", "cfrac", "binom", "dbinom"):
                    a, j = _take_group(t, j)
                    b, j = _take_group(t, j)
                    total += max(width(a), width(b)) * (0.8 if name == "tfrac" else 1.0) + 0.3
                elif name == "sqrt":
                    mm = re.match(r"\s*\[([^\]]*)\]", t[j:])
                    if mm:
                        j += mm.end()
                        total += 0.3
                    a, j = _take_group(t, j)
                    total += width(a) + 0.9
                elif name in ("text", "textbf", "textit", "mathrm", "mathbf", "operatorname", "mbox"):
                    a, j = _take_group(t, j)
                    total += width(a)
                elif name in ("left", "right", "displaystyle", "limits", "quad", "qquad"):
                    total += 0.9 if name in ("quad", "qquad") else 0.0
                elif name in ("int", "sum", "prod", "lim"):
                    total += 1.2
                else:
                    total += 0.8 if len(name) > 1 else 0.5
                i = j
                continue
            if c in "^_":
                a, j = _take_group(t, i + 1)
                total += width(a) * 0.7
                i = j
                continue
            if c in "{}$":
                i += 1
                continue
            total += 0.55 if not c.isspace() else 0.28
            i += 1
        return total

    return width(s)


def _take_group(t: str, j: int) -> tuple[str, int]:
    while j < len(t) and t[j] in " \t\n":
        j += 1
    if j < len(t) and t[j] == "{":
        depth = 0
        k = j
        while k < len(t):
            if t[k] == "\\":
                k += 2
                continue
            if t[k] == "{":
                depth += 1
            elif t[k] == "}":
                depth -= 1
                if depth == 0:
                    return t[j + 1: k], k + 1
            k += 1
        return t[j + 1:], len(t)
    if j < len(t) and t[j] == "\\":
        m = re.match(r"\\[A-Za-z]+|\\.", t[j:])
        return (m.group(0), j + m.end()) if m else ("", j + 1)
    return (t[j], j + 1) if j < len(t) else ("", j)


def choose_mcq_columns(options: list[str], avail_em: float = 40.0) -> int:
    """Chọn 4 / 2 / 1 cột: số cột lớn nhất mà phương án DÀI NHẤT vẫn vừa một ô (trừ nhãn + lề)."""
    widths = [_option_width_em(o) for o in options] or [0.0]
    for cols in (4, 2, 1):
        col_em = avail_em / cols - 2.4  # nhãn "A." + lề ô
        if max(widths) <= col_em:
            return cols
    return 1


def _label_format(label_tex: str) -> tuple[bool, bool]:
    """(in đậm, gạch chân) của nhãn phương án như nguồn viết — chỉ để GIỮ TRÌNH BÀY, không phải bằng chứng đáp án."""
    return ("\\textbf" in label_tex or "\\bf" in label_tex, "\\underline" in label_tex or "\\uline" in label_tex)


def build_multiple_choice_table(
    doc,
    options: list[str],
    correct_index: int | None = None,
    force_cols: int | None = None,
    *,
    label_formats: list[tuple[bool, bool]] | None = None,
    answer_source: str | None = None,
    avail_em: float | None = None,
):
    """Bảng 4 phương án A–D không viền, bố cục 1×4 / 2×2 / 4×1.

    * `correct_index` chỉ được truyền khi NGUỒN đánh dấu đáp án theo quy ước rõ ràng (\\choiceA..D, \\True);
      khi đó ghi `answer_key=explicit`. Không có → `answer_key=unknown`.
    * `label_formats` giữ đúng định dạng nhãn như nguồn (vd. tabularx tự viết \\underline{\\textbf{B.}}) —
      KHÔNG suy ra đáp án từ in đậm/gạch chân.
    * Chế độ `student` (ANSWER_MODE): không gạch chân/in đậm đánh dấu phương án đúng."""
    labels = ["A.", "B.", "C.", "D."]
    mode = ANSWER_MODE.get()
    if force_cols in (1, 2, 4):
        num_cols = force_cols
    else:
        num_cols = choose_mcq_columns(options, avail_em or 40.0)
    n_opts = min(len(options), 4) if options else 0
    num_rows = max(1, (max(n_opts, 1) + num_cols - 1) // num_cols)
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
            _set_tcw(cell, col_pct, "pct")
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            set_cell_borders(
                cell,
                top=("none", 0, "auto"),
                bottom=("none", 0, "auto"),
                left=("none", 0, "auto"),
                right=("none", 0, "auto"),
            )
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            if opt_idx < len(options) and opt_idx < 4:
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15

                lbl_run = p.add_run(labels[opt_idx] + " ")
                lbl_run.bold = True
                lbl_run.font.size = Pt(12)
                if mode != "student":
                    if label_formats and opt_idx < len(label_formats):
                        b, u = label_formats[opt_idx]
                        lbl_run.bold = True if b else lbl_run.bold
                        if u:
                            lbl_run.underline = True
                    elif correct_index == opt_idx:
                        lbl_run.underline = True

                append_rich_latex_text_to_paragraph(p, options[opt_idx].strip(), font_size_pt=12.0)
    if len(options) > 4:
        _event("mcq_extra_options", count=len(options))
        extra = doc.add_paragraph() if hasattr(doc, "add_paragraph") else None
        if extra is not None:
            for k, opt in enumerate(options[4:], start=5):
                append_rich_latex_text_to_paragraph(extra, f"({k}) {opt}", font_size_pt=12.0)

    keep_table_together(table)
    if correct_index is not None:
        key = "explicit"
    else:
        key = "unknown"
    shown = "hidden" if mode == "student" else ("shown" if (correct_index is not None or label_formats) else "none")
    _tag_role(table, "mcq", cols=num_cols, answer_key=key, answer_source=answer_source, answer_display=shown)
    _event("mcq", answer_key=key, answer_source=answer_source, cols=num_cols, answer_display=shown)
    return table


def build_true_false_table(doc, statements: list[tuple[str, str]]):
    """Bảng Đúng/Sai 4 mệnh đề (THPT 2025). Chế độ `student`: để trống cột Đ/S."""
    mode = ANSWER_MODE.get()
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
        _set_tcw(cell, col_pcts[c_i], "pct")
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="F1F5F9"/>'))
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        set_cell_borders(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt)
        r.bold = True
        r.font.size = Pt(11)

    sub_labels = ["a)", "b)", "c)", "d)"]
    explicit = 0
    for idx, (stmt_tex, ans) in enumerate(statements):
        row = table.rows[idx + 1]
        prevent_row_split(row)
        for c_i in range(3):
            cell = row.cells[c_i]
            _set_tcw(cell, col_pcts[c_i], "pct")
            set_cell_margins(cell, top=70, bottom=70, left=120, right=120)
            set_cell_borders(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p0 = row.cells[0].paragraphs[0]
        r_lbl = p0.add_run(f"{sub_labels[idx % 4]} ")
        r_lbl.bold = True
        r_lbl.font.size = Pt(11.5)
        append_rich_latex_text_to_paragraph(p0, stmt_tex, font_size_pt=11.5)

        ans_up = (ans or "").strip().upper()
        is_true = ans_up in ("Đ", "D", "TRUE", "1")
        is_false = ans_up in ("S", "FALSE", "0")
        explicit += int(is_true or is_false)
        if mode == "student":
            continue
        if is_true:
            p1 = row.cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run("Đ")
            r1.bold = True
            r1.font.size = Pt(11.5)
        elif is_false:
            p2 = row.cells[2].paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run("S")
            r2.bold = True
            r2.font.size = Pt(11.5)

    keep_table_together(table)
    key = "explicit" if explicit else "unknown"
    _tag_role(table, "true_false", answer_key=key, answer_display="hidden" if mode == "student" else "shown")
    _event("true_false", answer_key=key, explicit_answers=explicit)
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
            # Bảng phương án tự viết bằng tabularx: GIỮ định dạng nhãn như nguồn (in đậm/gạch chân), KHÔNG coi
            # in đậm/gạch chân là bằng chứng đáp án (bản cũ đánh dấu đáp án đúng chỉ vì ô có \textbf).
            opts: list[str] = []
            fmts: list[tuple[bool, bool]] = []
            label_re = re.compile(
                r"^\s*((?:\\(?:underline|textbf|uline)\s*\{\s*)*[A-D]\.(?:\s*\})*)\s*")
            for rc in raw_cells:
                m_lbl = label_re.match(rc)
                lbl = m_lbl.group(1) if m_lbl else ""
                fmts.append(_label_format(lbl))
                opts.append(rc[m_lbl.end():].strip() if m_lbl else rc.strip())
            return build_multiple_choice_table(doc, opts, correct_index=None, label_formats=fmts,
                                               answer_source="presentation_only")

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
    is_long = "longtable" in tabular_tex[:40] or num_rows > 12

    border_color = "475569" if has_vertical_bars else "94A3B8"
    b_single = ("single", 4, border_color)
    b_header_bottom = ("single", 8, "1E293B")
    b_none = ("none", 0, "auto")

    vmerge_remaining: dict[int, int] = {}

    for r_idx, r_cells in enumerate(parsed_rows):
        row = table.rows[r_idx]
        prevent_row_split(row)
        is_header = r_idx == 0
        if is_header and is_long:
            # bảng dài/longtable: lặp dòng tiêu đề khi sang trang
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

        col_cursor = 0
        overflow: list[str] = []
        for p_cell in r_cells:
            # Ô giữ chỗ (rỗng) mà LaTeX yêu cầu ở cột bị \multirow phủ: ô đó CHÍNH LÀ ô bị phủ → tiêu thụ nó,
            # không đẩy các ô sau lệch sang phải (trước đây làm rơi ô cuối hàng, vd. "Không đạt").
            placeholder = not p_cell.text.strip() and p_cell.diagbox is None
            covered_here = 0
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
                covered_here += 1
                if placeholder and covered_here >= p_cell.colspan:
                    break
            if placeholder and covered_here >= p_cell.colspan:
                continue

            if col_cursor >= num_cols:
                if p_cell.text.strip():
                    overflow.append(p_cell.text)
                continue

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
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="F8FAFC"/>'))

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

        if overflow:
            # hàng có nhiều ô hơn số cột: KHÔNG bỏ nội dung — ghép vào ô cuối và ghi sự kiện để báo cáo
            last = row.cells[num_cols - 1]
            for txt in overflow:
                append_rich_latex_text_to_paragraph(last.paragraphs[0], " " + txt, font_size_pt=11.0)
            _event("table_row_overflow", row=r_idx, cells=len(overflow))

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

    if not is_long:
        keep_table_together(table)
    _tag_role(table, "data", rows=num_rows, cols=num_cols)
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
    # KHÔNG khoá dòng (cantSplit): ô trái có thể dài hơn một trang, khoá dòng sẽ làm Word cắt mất nội dung.
    left_cell, right_cell = row.cells[0], row.cells[1]

    for cell, pct in ((left_cell, 3100), (right_cell, 1900)):
        _set_tcw(cell, pct, "pct")
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
    _tag_role(table, "layout_immini")
    return table
