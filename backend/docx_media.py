"""`backend/docx_media.py` — nhúng ảnh (SVG / EMF / PNG) vào .docx và thay chỗ đánh dấu.

Vì sao cần file riêng: `python-docx` chỉ biết chèn PNG/JPEG/GIF/BMP/TIFF, không biết
**SVG** (Word 2016+ hiển thị vector) và **EMF** (Word mọi đời hiển thị vector).  Module
này tự tạo part ảnh, tự dựng `<w:drawing>` và tự gắn `svgBlip` để:

* Word 2016 / 365 / Word trên Mac  → hiển thị **SVG vector**, phóng to bao nhiêu cũng nét;
* Word 2007–2013, WPS, Google Docs → tự rơi về **PNG độ phân giải cao** nhúng kèm.

Chỉ phục vụ nhánh dán mã; không có hàm nào bị engine gốc gọi tới.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import qn

__all__ = ["MediaInjector", "MARKER_RE", "make_marker"]

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

CONTENT_TYPES = {
    "png": "image/png", "svg": "image/svg+xml", "emf": "image/x-emf",
    "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
}

# Chỗ đánh dấu: chỉ gồm chữ cái + số nên đi qua LaTeX và engine gốc mà không bị đổi.
MARKER_PREFIX = "ZLZTRUTHZ"
MARKER_SUFFIX = "ZEND"
MARKER_RE = re.compile(MARKER_PREFIX + r"(\d{5})" + MARKER_SUFFIX)


def make_marker(index: int) -> str:
    return f"{MARKER_PREFIX}{index:05d}{MARKER_SUFFIX}"


def _el(xml: str):
    return parse_xml(xml)


class MediaInjector:
    """Gắn ảnh vào một `docx.Document` đang mở và thay chỗ đánh dấu bằng ảnh."""

    def __init__(self, document) -> None:
        self.doc = document
        self.part = document.part
        self._rid_cache: dict[bytes, str] = {}
        self._seq = 0
        self.stats: dict[str, Any] = {"images_added": 0, "markers_replaced": 0, "markers_missing": 0}

    # ---------------- part ảnh ----------------
    def _add_image_part(self, data: bytes, ext: str) -> str:
        if not data:
            return ""
        cached = self._rid_cache.get(data)
        if cached:
            return cached
        self._seq += 1
        ext = ext.lower().lstrip(".")
        ct = CONTENT_TYPES.get(ext, "application/octet-stream")
        partname = PackURI(f"/word/media/l2w_truth_{self._seq}.{ext}")
        image_part = Part(partname, ct, data, self.part.package)
        rid = self.part.relate_to(image_part, RT.IMAGE)
        self._rid_cache[data] = rid
        self.stats["images_added"] += 1
        return rid

    # ---------------- dựng <w:drawing> ----------------
    def build_drawing(self, *, main_data: bytes, main_ext: str, png_data: bytes,
                      cx: int, cy: int, name: str = "Hình LaTeX") -> Any:
        """Trả về phần tử `<w:drawing>` chứa ảnh (SVG → kèm PNG dự phòng)."""
        main_ext = (main_ext or "png").lower()
        if main_ext == "svg":
            png_rid = self._add_image_part(png_data or main_data, "png" if png_data else "svg")
            svg_rid = self._add_image_part(main_data, "svg")
            blip = (f'<a:blip xmlns:a="{NS_A}" xmlns:r="{NS_R}" r:embed="{png_rid}">'
                    f'<a:extLst><a:ext uri="{SVG_EXT_URI}">'
                    f'<asvg:svgBlip xmlns:asvg="{NS_ASVG}" r:embed="{svg_rid}"/>'
                    f'</a:ext></a:extLst></a:blip>')
        else:
            rid = self._add_image_part(main_data, main_ext)
            blip = f'<a:blip xmlns:a="{NS_A}" xmlns:r="{NS_R}" r:embed="{rid}"/>'
        self._seq += 1
        doc_pr_id = 9000 + self._seq
        xml = (
            f'<w:drawing xmlns:w="{NS_W}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" '
            f'xmlns:pic="{NS_PIC}" xmlns:r="{NS_R}">'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{doc_pr_id}" name="{_xml_attr(name)}{doc_pr_id}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="{_xml_attr(name)}{doc_pr_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
        )
        return _el(xml)

    # ---------------- thay chỗ đánh dấu ----------------
    def replace_markers(self, resolver) -> dict[str, Any]:
        """Quét toàn bộ tài liệu (kể cả bảng, header, footer) và thay mỗi chỗ đánh dấu
        bằng ảnh tương ứng.  `resolver(index) -> dict | None` cung cấp dữ liệu ảnh."""
        targets = [self.doc.element.body]
        for sec in self.doc.sections:
            for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                       sec.even_page_header, sec.even_page_footer):
                try:
                    targets.append(hf._element)
                except Exception:
                    pass
        for root in targets:
            for para in list(root.iter(qn("w:p"))):
                self._replace_in_math(para, resolver)
                self._replace_in_paragraph(para, resolver)
        return self.stats

    def _replace_in_math(self, para, resolver) -> None:
        """Dấu đánh dấu lọt vào trong công thức OMML → thay NGUYÊN khối công thức bằng ảnh."""
        for omath in list(para.iter(qn("m:oMath"))):
            text = "".join(t.text or "" for t in omath.iter(qn("m:t")))
            m = MARKER_RE.search(text)
            if not m:
                continue
            info = resolver(int(m.group(1)))
            parent = omath.getparent()
            if parent is None:
                continue
            pos = list(parent).index(omath)
            parent.remove(omath)
            if not info:
                self.stats["markers_missing"] += 1
                continue
            run = _el(f'<w:r xmlns:w="{NS_W}"/>')
            if info.get("position_half_pt"):
                _set_position(run, int(info["position_half_pt"]))
            run.append(self.build_drawing(
                main_data=info["data"], main_ext=info["ext"], png_data=info.get("png", b""),
                cx=info["cx"], cy=info["cy"], name=info.get("name", "Công thức LaTeX")))
            parent.insert(pos, run)
            self.stats["markers_replaced"] += 1

    def _replace_in_paragraph(self, para, resolver) -> None:
        runs = [r for r in para.iter(qn("w:r")) if r.findall(qn("w:t"))]
        if not runs:
            return
        full = "".join("".join(t.text or "" for t in r.findall(qn("w:t"))) for r in runs)
        if MARKER_PREFIX not in full:
            return
        matches = list(MARKER_RE.finditer(full))
        if not matches:
            return
        alone = full.strip() == "".join(m.group(0) for m in matches) and len(matches) == 1
        # ánh xạ offset ký tự → (run, phần tử w:t, vị trí trong w:t)
        spans: list[tuple[int, int, Any, Any]] = []
        pos = 0
        for r in runs:
            for t in r.findall(qn("w:t")):
                txt = t.text or ""
                spans.append((pos, pos + len(txt), r, t))
                pos += len(txt)
        for m in reversed(matches):
            idx = int(m.group(1))
            info = resolver(idx)
            if not info:
                self.stats["markers_missing"] += 1
                continue
            drawing = self.build_drawing(
                main_data=info["data"], main_ext=info["ext"], png_data=info.get("png", b""),
                cx=info["cx"], cy=info["cy"], name=info.get("name", "Hình LaTeX"),
            )
            run = _el(f'<w:r xmlns:w="{NS_W}"/>')
            rpr = info.get("rPr")
            if rpr is not None:
                run.append(rpr)
            if info.get("position_half_pt"):
                _set_position(run, int(info["position_half_pt"]))
            run.append(drawing)
            if alone:
                for r in runs:
                    parent = r.getparent()
                    if parent is not None:
                        parent.remove(r)
                para.append(run)
                _center(para)
            else:
                self._splice(para, spans, m.start(), m.end(), run)
            self.stats["markers_replaced"] += 1

    @staticmethod
    def _splice(para, spans, start: int, end: int, new_run) -> None:
        """Cắt đoạn [start,end) khỏi chuỗi chữ và đặt `new_run` vào đúng chỗ đó."""
        anchor = None
        for s, e, r, t in spans:
            if e <= start or s >= end:
                continue
            txt = t.text or ""
            left = txt[: max(0, start - s)] if s < start else ""
            right = txt[max(0, end - s):] if e > end else ""
            t.text = left + right
            if anchor is None:
                anchor = r
        if anchor is None:
            para.append(new_run)
            return
        parent = anchor.getparent()
        parent.insert(list(parent).index(anchor) + 1, new_run)


def _center(para) -> None:
    ppr = para.find(qn("w:pPr"))
    if ppr is None:
        ppr = _el(f'<w:pPr xmlns:w="{NS_W}"/>')
        para.insert(0, ppr)
    for old in ppr.findall(qn("w:jc")):
        ppr.remove(old)
    ppr.append(_el(f'<w:jc xmlns:w="{NS_W}" w:val="center"/>'))


def _set_position(run, half_points: int) -> None:
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = _el(f'<w:rPr xmlns:w="{NS_W}"/>')
        run.insert(0, rpr)
    for old in rpr.findall(qn("w:position")):
        rpr.remove(old)
    rpr.append(_el(f'<w:position xmlns:w="{NS_W}" w:val="{half_points}"/>'))


def _xml_attr(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;").replace("'", "&apos;"))
