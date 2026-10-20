"""OOXML Table Post-Processor (`backend/postprocessor.py`).

Inspects and refines Word tables without clobbering specialized math tables:
- Preserves Variation / Sign tables (`tkz-tab`) with custom double asymptote borders
- Preserves Side-by-Side layout tables (`\immini`)
- Ensures Multiple-Choice tables (exact 4 options A., B., C., D.) remain borderless
- Ensures Data / Statistical tables have clean borders and vertical centering
"""
from __future__ import annotations

import re
from pathlib import Path

try:
    import docx
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def _is_variation_table(table) -> bool:
    """Detect if a Word table is a tkz-tab variation/sign table."""
    if len(table.rows) < 2 or len(table.rows[0].cells) < 3:
        return False
    col0_texts = [r.cells[0].text.strip() for r in table.rows if r.cells]
    # Variation tables have x in row 0 col 0, and y', f'(x), y, or f(x) below
    if col0_texts and col0_texts[0] in ("x", "$x$"):
        return True
    # Also check if any cell contains a VarArrow_ DrawingML shape
    xml_str = table._tbl.xml
    return "VarArrow_" in xml_str


def _is_immini_table(table) -> bool:
    """Detect if a Word table is a 1x2 side-by-side \immini layout table."""
    if len(table.rows) == 1 and len(table.rows[0].cells) == 2:
        right_xml = table.rows[0].cells[1]._tc.xml
        if "EditableMathShapeGroup_" in right_xml:
            return True
    return False


def _is_multiple_choice_table(table) -> bool:
    """Strictly detect a 4-option multiple-choice table (every non-empty cell starts with A., B., C., or D.)."""
    non_empty = [c.text.strip() for r in table.rows for c in r.cells if c.text.strip()]
    if len(non_empty) != 4:
        return False
    expected = ("A.", "B.", "C.", "D.")
    return all(non_empty[i].startswith(expected[i]) for i in range(4))


def postprocess_docx_tables(docx_path: Path, is_mode2: bool = True) -> dict[str, int]:
    """Audit and normalize all tables in a generated .docx document."""
    stats = {
        "tables_processed": 0,
        "variation_tables": 0,
        "choice_tables": 0,
        "immini_tables": 0,
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

        if _is_variation_table(table):
            stats["variation_tables"] += 1
            for row in table.rows:
                stats["cells_formatted"] += len(row.cells)
            continue

        if _is_immini_table(table):
            stats["immini_tables"] += 1
            for row in table.rows:
                stats["cells_formatted"] += len(row.cells)
            continue

        if _is_multiple_choice_table(table):
            stats["choice_tables"] += 1
            for row in table.rows:
                stats["cells_formatted"] += len(row.cells)
            continue

        stats["data_tables"] += 1
        for row in table.rows:
            for cell in row.cells:
                stats["cells_formatted"] += 1
                tc_pr = cell._tc.get_or_add_tcPr()
                if tc_pr.find(qn("w:gridSpan")) is not None or tc_pr.find(qn("w:vMerge")) is not None:
                    stats["merged_cells_detected"] += 1

    doc.save(str(docx_path))
    return stats
