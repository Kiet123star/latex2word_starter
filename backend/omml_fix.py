"""`backend/omml_fix.py` — sửa dấu mũ/mũi tên trong công thức Word (OMML).

Trong Word, dấu đặt **trên đầu** một ký hiệu (`\\vec`, `\\overrightarrow`, `\\hat`, `\\bar`…)
phải dùng **ký tự tổ hợp** (combining character), ví dụ U+20D7 cho mũi tên vectơ.
Engine gốc đang dùng ký tự thường — `→` (U+2192) cho `\\vec` — nên Word vẽ một mũi tên
cỡ lớn nằm ngang **đè lên chữ**, trông như gạch ngang. Đó chính là lỗi "mũi tên vectơ"
bạn gặp ở `$\\vec{u} = \\overrightarrow{DC'} + \\overrightarrow{AD'}$`.

Module này **sửa thẳng trên file .docx đã tạo**, nên:

* không phải sửa `backend/omml_math.py` → **nhánh tải file không bị đụng tới**;
* sửa được mọi dấu đặt trên, không riêng mũi tên.

Nếu bạn muốn nhánh tải file cũng hết lỗi này, sửa `ACCENT_MAP` trong
`backend/omml_math.py` theo bảng `COMBINING` bên dưới (đổi đúng mấy ký tự, không đổi gì khác).
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx.oxml.ns import qn

__all__ = ["fix_math_accents", "COMBINING"]

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

#: ký tự engine đang dùng (sai) → ký tự tổ hợp Word cần (đúng)
COMBINING: dict[str, str] = {
    "→": "⃗",   # →  mũi tên phải      → combining right arrow above   (\vec, \overrightarrow)
    "←": "⃖",   # ←  mũi tên trái      → combining left arrow above    (\overleftarrow)
    "↔": "⃡",   # ↔  mũi tên hai chiều → combining left-right arrow    (\overleftrightarrow)
    "^": "̂",   # ^  dấu mũ ASCII      → combining circumflex          (\hat, \widehat)
    "ˆ": "̂",   # ˆ  modifier circumflex
    "~": "̃",   # ~  ngã ASCII         → combining tilde               (\tilde, \widetilde)
    "˜": "̃",   # ˜  small tilde
    "¯": "̄",   # ¯  macron            → combining macron              (\bar)
    "ˉ": "̄",   # ˉ  modifier macron
    "¨": "̈",   # ¨  diaeresis         → combining diaeresis           (\ddot)
    "˙": "̇",   # ˙  dot above         → combining dot above           (\dot)
    "´": "́",   # ´  acute             → combining acute               (\acute)
    "`": "̀",   # `  grave             → combining grave               (\grave)
    "˘": "̆",   # ˘  breve             → combining breve               (\breve)
    "ˇ": "̌",   # ˇ  caron             → combining caron               (\check)
    "˚": "̊",   # ˚  ring above        → combining ring above          (\mathring)
}

#: Dấu đã đúng rồi thì để yên.
_ALREADY_OK = set(COMBINING.values()) | {"̅", "̲", "⏞", "⏟", "⎴", "⎵"}


def fix_math_accents(docx_path: Path) -> dict[str, Any]:
    """Quét .docx, đổi mọi `m:acc/m:chr` sai thành ký tự tổ hợp đúng.

    Không bao giờ ném lỗi — hỏng thì trả về số 0 và để file nguyên như cũ."""
    stats: dict[str, Any] = {"da_sua": 0, "ky_tu": {}}
    try:
        import docx as _docx
        doc = _docx.Document(str(docx_path))
        roots = [doc.element.body]
        for sec in doc.sections:
            for hf in (sec.header, sec.footer, sec.first_page_header,
                       sec.first_page_footer, sec.even_page_header, sec.even_page_footer):
                try:
                    roots.append(hf._element)
                except Exception:
                    pass
        val = f"{{{M_NS}}}val"
        changed = False
        for root in roots:
            for acc in root.iter(qn("m:acc")):
                for chr_el in acc.iter(qn("m:chr")):
                    cur = chr_el.get(val) or ""
                    if cur in _ALREADY_OK or not cur:
                        continue
                    new = COMBINING.get(cur)
                    if new and new != cur:
                        chr_el.set(val, new)
                        stats["da_sua"] += 1
                        stats["ky_tu"][cur] = stats["ky_tu"].get(cur, 0) + 1
                        changed = True
        if changed:
            doc.save(str(docx_path))
    except Exception as exc:
        stats["loi"] = f"{type(exc).__name__}: {exc}"
    return stats
