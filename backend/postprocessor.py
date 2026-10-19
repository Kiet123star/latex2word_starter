import re
from pathlib import Path

try:
    import docx
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def postprocess_docx_tables(docx_path: Path, is_mode2: bool = True) -> dict[str, int]:
    """Chuẩn hóa bảng trong Word:
    1. Bảng lựa chọn trắc nghiệm (A, B, C, D): gỡ bỏ hoàn toàn đường viền (w:val="none"),
       căn đều 100% độ rộng cột, loại bỏ viền ô để hiển thị thoáng đẹp như PDF gốc.
    2. Bảng dữ liệu thông thường: tạo đường viền thanh lịch và phát hiện ô gộp.
    """
    stats = {
        "tables_processed": 0,
        "choice_tables": 0,
        "data_tables": 0,
        "cells_formatted": 0,
        "merged_cells_detected": 0,
    }
    if not HAS_DOCX or not docx_path.exists():
        return stats

    doc = docx.Document(str(docx_path))
    for table in doc.tables:
        stats["tables_processed"] += 1
        tbl_pr = table._tbl.tblPr
        if tbl_pr is None:
            continue

        # Check if table contains multiple-choice options (e.g. A. ..., B. ...)
        all_text = " ".join(c.text.strip() for row in table.rows for c in row.cells)
        is_choice = bool(
            re.search(r"\b[A-D]\.\s", all_text) or
            re.search(r"^[A-D]\.", all_text)
        )

        # Remove existing borders to prevent Word grid inheritance
        existing_borders = tbl_pr.find(qn("w:tblBorders"))
        if existing_borders is not None:
            tbl_pr.remove(existing_borders)

        new_borders = OxmlElement("w:tblBorders")

        if is_choice:
            stats["choice_tables"] += 1
            # Multiple-choice layout: strictly NO borders
            for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
                border = OxmlElement(f"w:{name}")
                border.set(qn("w:val"), "none")
                border.set(qn("w:sz"), "0")
                border.set(qn("w:space"), "0")
                border.set(qn("w:color"), "auto")
                new_borders.append(border)
            tbl_pr.append(new_borders)

            # Set table width to 100% (5000 pct)
            tbl_w = tbl_pr.find(qn("w:tblW"))
            if tbl_w is None:
                tbl_w = OxmlElement("w:tblW")
                tbl_pr.append(tbl_w)
            tbl_w.set(qn("w:w"), "5000")
            tbl_w.set(qn("w:type"), "pct")

            # Equalize column widths across rows
            num_cols = max(len(row.cells) for row in table.rows) if table.rows else 0
            if num_cols > 0:
                col_pct = str(5000 // num_cols)
                for row in table.rows:
                    for cell in row.cells:
                        tc_pr = cell._tc.get_or_add_tcPr()
                        # Remove cell-level borders if any
                        tc_borders = tc_pr.find(qn("w:tcBorders"))
                        if tc_borders is not None:
                            tc_pr.remove(tc_borders)
                        # Set cell width
                        tc_w = tc_pr.find(qn("w:tcW"))
                        if tc_w is None:
                            tc_w = OxmlElement("w:tcW")
                            tc_pr.append(tc_w)
                        tc_w.set(qn("w:w"), col_pct)
                        tc_w.set(qn("w:type"), "pct")
                        stats["cells_formatted"] += 1
        else:
            stats["data_tables"] += 1
            if is_mode2:
                # Genuine data table: apply clean modern borders
                for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
                    border = OxmlElement(f"w:{name}")
                    border.set(qn("w:val"), "single")
                    border.set(qn("w:sz"), "4")
                    border.set(qn("w:space"), "0")
                    border.set(qn("w:color"), "D0D5DD")
                    new_borders.append(border)
                tbl_pr.append(new_borders)

                for row in table.rows:
                    for cell in row.cells:
                        stats["cells_formatted"] += 1
                        tc_pr = cell._tc.get_or_add_tcPr()
                        if (tc_pr.find(qn("w:gridSpan")) is not None
                                or tc_pr.find(qn("w:vMerge")) is not None):
                            stats["merged_cells_detected"] += 1

    doc.save(str(docx_path))
    return stats
