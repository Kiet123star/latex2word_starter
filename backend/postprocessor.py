"""OOXML Table Post-Processor (`backend/postprocessor.py`).

Kiểm kê bảng Word theo VAI TRÒ:
* ưu tiên metadata `w:tblDescription = "l2w:role=…"` do bộ dựng bảng ghi (mcq, true_false, variation,
  data, layout_immini…);
* chỉ khi bảng không có metadata (vd. engine cũ) mới đoán theo cấu trúc — và ghi rõ là heuristic.

Bản cũ đoán bảng biến thiên chỉ vì ô đầu là "x" (bảng số liệu x|1|2|3 bị nhận nhầm) và nhận bảng
dàn trang immini theo tên media; nay vai trò lấy từ metadata.

Module này KHÔNG sửa nội dung bảng nên không ghi lại tệp (tránh lưu lại không cần thiết).
"""
from __future__ import annotations

from pathlib import Path

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

from .ooxml_tools import get_table_role


def _is_variation_table(table) -> bool:
    """Heuristic (chỉ dùng khi thiếu metadata): bảng có mũi tên biến thiên DrawingML."""
    return "VarArrow_" in table._tbl.xml


def _is_immini_table(table) -> bool:
    """Heuristic (chỉ dùng khi thiếu metadata): 1×2, ô phải chứa hình (nhóm DrawingML hoặc ảnh)."""
    if len(table.rows) == 1 and len(table.rows[0].cells) == 2:
        right_xml = table.rows[0].cells[1]._tc.xml
        return "EditableMathShapeGroup_" in right_xml or "<pic:pic" in right_xml or "wpg:wgp" in right_xml
    return False


def _is_multiple_choice_table(table) -> bool:
    """Heuristic: đúng 4 ô có chữ, lần lượt bắt đầu bằng A., B., C., D."""
    non_empty = [c.text.strip() for r in table.rows for c in r.cells if c.text.strip()]
    if len(non_empty) != 4:
        return False
    expected = ("A.", "B.", "C.", "D.")
    return all(non_empty[i].startswith(expected[i]) for i in range(4))


_ROLE_KEY = {
    "variation": "variation_tables", "mcq": "choice_tables", "true_false": "choice_tables",
    "layout_immini": "immini_tables", "layout_mcq_pdf": "choice_tables", "data": "data_tables",
}


def postprocess_docx_tables(docx_path: Path, is_mode2: bool = True) -> dict[str, int]:
    """Kiểm kê toàn bộ bảng của tệp .docx (đọc, không ghi)."""
    stats = {
        "tables_processed": 0,
        "variation_tables": 0,
        "choice_tables": 0,
        "immini_tables": 0,
        "data_tables": 0,
        "other_tables": 0,
        "cells_formatted": 0,
        "merged_cells_detected": 0,
        "role_from_metadata": 0,
        "role_from_heuristic": 0,
    }
    if not HAS_DOCX or not Path(docx_path).exists():
        return stats

    doc = docx.Document(str(docx_path))
    for tbl_el in doc.element.body.iter(qn("w:tbl")):
        table = Table(tbl_el, doc._body)
        stats["tables_processed"] += 1
        meta = get_table_role(tbl_el)
        if meta:
            stats["role_from_metadata"] += 1
            role = meta["role"]
        else:
            stats["role_from_heuristic"] += 1
            if _is_variation_table(table):
                role = "variation"
            elif _is_immini_table(table):
                role = "layout_immini"
            elif _is_multiple_choice_table(table):
                role = "mcq"
            else:
                role = "data"
        stats[_ROLE_KEY.get(role, "other_tables")] += 1
        for tc in tbl_el.iter(qn("w:tc")):
            stats["cells_formatted"] += 1
            tc_pr = tc.find(qn("w:tcPr"))
            if role == "data" and tc_pr is not None and (
                    tc_pr.find(qn("w:gridSpan")) is not None or tc_pr.find(qn("w:vMerge")) is not None):
                stats["merged_cells_detected"] += 1
    return stats
