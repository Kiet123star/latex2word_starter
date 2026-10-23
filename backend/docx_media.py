"""`backend/docx_media.py` — nhúng ảnh (SVG / EMF / PNG / JPEG) vào .docx và thay chỗ đánh dấu.

Vì sao cần file riêng: `python-docx` chỉ biết chèn PNG/JPEG/GIF/BMP/TIFF, không biết
**SVG** (Word 2016+ hiển thị vector) và **EMF**.  Module này tự tạo part ảnh, tự dựng
`<w:drawing>` và gắn `svgBlip` để:

* Word 2016 / 365 / Word trên Mac  → hiển thị **SVG vector**;
* Word 2007–2013, WPS, Google Docs → dùng **PNG** nhúng kèm (cùng một đối tượng logic).

Thay đổi so với bản trước (sửa lỗi đóng gói):
* `docPr id` được cấp tiếp sau id lớn nhất đang có trong tài liệu (trước đây cố định 9000+n
  nên có thể trùng khi chạy nhiều lần / khi engine đã dùng id lớn);
* tên part ảnh lấy từ `package.next_partname` → không trùng khi MediaInjector chạy nhiều lần;
* chỗ đánh dấu nằm trong công thức mà resolver không cấp được ảnh → GIỮ NGUYÊN công thức
  (trước đây công thức bị xoá mất);
* resolver có thể trả `{"text": ...}` để thay chỗ đánh dấu bằng thông báo hiển thị cho
  khối không dựng được (không bao giờ xoá âm thầm);
* có thể dùng regex chỗ đánh dấu khác (ảnh `\\includegraphics`, hình chụp TeX của pipeline).
"""
from __future__ import annotations

import re
from typing import Any, Callable

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import qn

__all__ = ["MediaInjector", "MARKER_RE", "make_marker", "IMG_MARKER_RE", "make_img_marker", "xml_attr",
           "next_docpr_id", "sanitize_xml_text"]

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
# Chỗ đánh dấu riêng của pipeline tải tệp (ảnh \includegraphics, hình chụp TeX) — không trùng
# với chỗ đánh dấu của lớp phủ universal_fallback.
IMG_MARKER_PREFIX = "ZLZIMGZ"
IMG_MARKER_RE = re.compile(IMG_MARKER_PREFIX + r"(\d{5})" + MARKER_SUFFIX)

_ILLEGAL_XML_CHARS_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ud800-\udfff￾￿]")


def sanitize_xml_text(s: str) -> str:
    return _ILLEGAL_XML_CHARS_RE.sub("", s or "")


def make_marker(index: int) -> str:
    return f"{MARKER_PREFIX}{index:05d}{MARKER_SUFFIX}"


def make_img_marker(index: int) -> str:
    return f"{IMG_MARKER_PREFIX}{index:05d}{MARKER_SUFFIX}"


def _el(xml: str):
    return parse_xml(xml)


def next_docpr_id(document) -> int:
    """1 + id số lớn nhất đang có trong document.xml (docPr, cNvPr, bookmark…)."""
    ids = [0]
    for v in document.element.xpath("//@id"):
        try:
            ids.append(int(v))
        except (TypeError, ValueError):
            pass
    return max(ids) + 1


def _own_runs(para) -> list:
    """Các w:r thuộc chính đoạn này (không lấy run trong textbox/đoạn lồng bên trong)."""
    out = []
    for r in para.iter(qn("w:r")):
        anc = r.getparent()
        while anc is not None and anc.tag != qn("w:p"):
            anc = anc.getparent()
        if anc is para:
            out.append(r)
    return out


class MediaInjector:
    """Gắn ảnh vào một `docx.Document` đang mở và thay chỗ đánh dấu bằng ảnh."""

    def __init__(self, document) -> None:
        self.doc = document
        self.part = document.part
        self._rid_cache: dict[tuple[bytes, str], str] = {}
        self._seq = 0
        self._next_id = next_docpr_id(document)
        self.stats: dict[str, Any] = {"images_added": 0, "markers_replaced": 0, "markers_missing": 0,
                                      "markers_text_fallback": 0, "logical_images": 0}
        self.found: set[int] = set()

    # ---------------- part ảnh ----------------
    def _add_image_part(self, data: bytes, ext: str) -> str:
        if not data:
            return ""
        ext = ext.lower().lstrip(".")
        if ext == "jpeg":
            ext = "jpg"
        cached = self._rid_cache.get((data, ext))
        if cached:
            return cached
        ct = CONTENT_TYPES.get(ext, "application/octet-stream")
        partname = self.part.package.next_partname(f"/word/media/l2w_media%d.{ext}")
        image_part = Part(PackURI(str(partname)), ct, data, self.part.package)
        rid = self.part.relate_to(image_part, RT.IMAGE)
        self._rid_cache[(data, ext)] = rid
        self.stats["images_added"] += 1
        return rid

    # ---------------- dựng <w:drawing> ----------------
    def build_drawing(self, *, main_data: bytes, main_ext: str, png_data: bytes,
                      cx: int, cy: int, name: str = "Hình LaTeX", descr: str = "") -> Any:
        """Trả về phần tử `<w:drawing>` chứa ảnh (SVG → kèm PNG dự phòng; một đối tượng logic)."""
        main_ext = (main_ext or "png").lower()
        if main_ext == "svg" and not png_data:
            rid = self._add_image_part(main_data, "svg")
            blip = f'<a:blip xmlns:a="{NS_A}" xmlns:r="{NS_R}" r:embed="{rid}"/>'
        elif main_ext == "svg":
            png_rid = self._add_image_part(png_data, "png")
            svg_rid = self._add_image_part(main_data, "svg")
            blip = (f'<a:blip xmlns:a="{NS_A}" xmlns:r="{NS_R}" r:embed="{png_rid}">'
                    f'<a:extLst><a:ext uri="{SVG_EXT_URI}">'
                    f'<asvg:svgBlip xmlns:asvg="{NS_ASVG}" r:embed="{svg_rid}"/>'
                    f'</a:ext></a:extLst></a:blip>')
        else:
            rid = self._add_image_part(main_data, main_ext)
            blip = f'<a:blip xmlns:a="{NS_A}" xmlns:r="{NS_R}" r:embed="{rid}"/>'
        doc_pr_id = self._next_id
        self._next_id += 1
        self._seq += 1
        self.stats["logical_images"] += 1
        nm = xml_attr(name)
        ds = xml_attr(descr)
        xml = (
            f'<w:drawing xmlns:w="{NS_W}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" '
            f'xmlns:pic="{NS_PIC}" xmlns:r="{NS_R}">'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{int(cx)}" cy="{int(cy)}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{doc_pr_id}" name="{nm}" descr="{ds}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="0" name="{nm}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
        )
        return _el(xml)

    def _text_run(self, text: str, rpr=None):
        run = _el(f'<w:r xmlns:w="{NS_W}"/>')
        if rpr is not None:
            run.append(rpr)
        t = _el(f'<w:t xmlns:w="{NS_W}" xml:space="preserve"/>')
        t.text = sanitize_xml_text(text)
        run.append(t)
        return run

    # ---------------- thay chỗ đánh dấu ----------------
    def replace_markers(self, resolver: Callable[[int], dict[str, Any] | None],
                        marker_re: re.Pattern = MARKER_RE) -> dict[str, Any]:
        """Quét toàn bộ tài liệu (kể cả bảng, header, footer) và thay mỗi chỗ đánh dấu bằng ảnh
        tương ứng. `resolver(index)` trả dict ảnh, `{"text": ...}` (thông báo hiển thị) hoặc None
        (giữ nguyên chỗ đánh dấu để bước sau phát hiện — không xoá âm thầm)."""
        targets = [self.doc.element.body]
        seen_parts = set()
        for sec in self.doc.sections:
            for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                       sec.even_page_header, sec.even_page_footer):
                try:
                    if hf.is_linked_to_previous:
                        continue
                    el = hf._element
                    if id(el) in seen_parts:
                        continue
                    seen_parts.add(id(el))
                    targets.append(el)
                except Exception:
                    pass
        for root in targets:
            for para in list(root.iter(qn("w:p"))):
                self._replace_in_math(para, resolver, marker_re)
                self._replace_in_paragraph(para, resolver, marker_re)
        return self.stats

    def _replace_in_math(self, para, resolver, marker_re) -> None:
        """Chỗ đánh dấu lọt vào công thức OMML → thay NGUYÊN khối công thức bằng ảnh (nếu có ảnh)."""
        for omath in list(para.iter(qn("m:oMath"))):
            text = "".join(t.text or "" for t in omath.iter(qn("m:t")))
            m = marker_re.search(text)
            if not m:
                continue
            idx = int(m.group(1))
            self.found.add(idx)
            info = resolver(idx)
            parent = omath.getparent()
            if parent is None or not info:
                self.stats["markers_missing"] += 1
                continue                      # giữ nguyên công thức — không xoá
            pos = list(parent).index(omath)
            if "text" in info and "data" not in info:
                new_el = self._text_run(info["text"])
                self.stats["markers_text_fallback"] += 1
            else:
                new_el = _el(f'<w:r xmlns:w="{NS_W}"/>')
                if info.get("position_half_pt"):
                    _set_position(new_el, int(info["position_half_pt"]))
                new_el.append(self.build_drawing(
                    main_data=info["data"], main_ext=info["ext"], png_data=info.get("png", b""),
                    cx=info["cx"], cy=info["cy"], name=info.get("name", "Công thức LaTeX"),
                    descr=info.get("descr", "")))
                self.stats["markers_replaced"] += 1
            if parent.tag == qn("m:oMathPara"):
                # oMathPara chỉ chứa oMath → thay cả oMathPara bằng run trong đoạn
                gp = parent.getparent()
                gpos = list(gp).index(parent)
                gp.remove(parent)
                gp.insert(gpos, new_el)
            else:
                parent.remove(omath)
                parent.insert(pos, new_el)

    def _replace_in_paragraph(self, para, resolver, marker_re) -> None:
        runs = [r for r in _own_runs(para) if r.findall(qn("w:t"))]
        if not runs:
            return
        full = "".join("".join(t.text or "" for t in r.findall(qn("w:t"))) for r in runs)
        matches = list(marker_re.finditer(full))
        if not matches:
            return
        alone = full.strip() == "".join(m.group(0) for m in matches) and len(matches) == 1
        spans: list[tuple[int, int, Any, Any]] = []
        pos = 0
        for r in runs:
            for t in r.findall(qn("w:t")):
                txt = t.text or ""
                spans.append((pos, pos + len(txt), r, t))
                pos += len(txt)
        for m in reversed(matches):
            idx = int(m.group(1))
            self.found.add(idx)
            info = resolver(idx)
            if not info:
                self.stats["markers_missing"] += 1
                continue
            if "text" in info and "data" not in info:
                rpr = None
                new_run = self._text_run(info["text"], rpr)
                self.stats["markers_text_fallback"] += 1
                self._splice(para, spans, m.start(), m.end(), new_run)
                continue
            drawing = self.build_drawing(
                main_data=info["data"], main_ext=info["ext"], png_data=info.get("png", b""),
                cx=info["cx"], cy=info["cy"], name=info.get("name", "Hình LaTeX"), descr=info.get("descr", ""),
            )
            run = _el(f'<w:r xmlns:w="{NS_W}"/>')
            rpr = info.get("rPr")
            if rpr is not None:
                run.append(rpr)
            if info.get("position_half_pt"):
                _set_position(run, int(info["position_half_pt"]))
            run.append(drawing)
            if alone and info.get("block", True):
                for r in runs:
                    parent = r.getparent()
                    if parent is not None:
                        parent.remove(r)
                para.append(run)
                if info.get("center", True):
                    _center(para)
            else:
                self._splice(para, spans, m.start(), m.end(), run)
            self.stats["markers_replaced"] += 1

    @staticmethod
    def _splice(para, spans, start: int, end: int, new_run) -> None:
        """Cắt đoạn [start,end) khỏi chuỗi chữ và đặt `new_run` vào đúng chỗ đó."""
        anchor = None
        anchor_left_empty = False
        for s, e, r, t in spans:
            if e <= start or s >= end:
                continue
            txt = t.text or ""
            left = txt[: max(0, start - s)] if s < start else ""
            right = txt[max(0, end - s):] if e > end else ""
            if anchor is None:
                anchor = r
                anchor_left_empty = left == ""
                if right and left:
                    # chỗ đánh dấu nằm giữa một run: tách run, ảnh nằm đúng giữa hai phần chữ
                    t.text = left
                    _preserve(t)
                    clone = _clone_run_with_text(r, right)
                    parent = r.getparent()
                    idx = list(parent).index(r)
                    parent.insert(idx + 1, new_run)
                    parent.insert(idx + 2, clone)
                    return
            t.text = left + right
            _preserve(t)
        if anchor is None:
            para.append(new_run)
            return
        parent = anchor.getparent()
        idx = list(parent).index(anchor)
        # phần chữ trước chỗ đánh dấu rỗng → ảnh đứng TRƯỚC run neo; ngược lại đứng sau
        parent.insert(idx if anchor_left_empty else idx + 1, new_run)


def _clone_run_with_text(run, text: str):
    import copy
    clone = copy.deepcopy(run)
    ts = clone.findall(qn("w:t"))
    for t in ts[1:]:
        clone.remove(t)
    if ts:
        ts[0].text = text
        _preserve(ts[0])
    return clone


def _preserve(t) -> None:
    if t.text and (t.text[:1].isspace() or t.text[-1:].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _center(para) -> None:
    ppr = para.find(qn("w:pPr"))
    if ppr is None:
        ppr = _el(f'<w:pPr xmlns:w="{NS_W}"/>')
        para.insert(0, ppr)
    for old in ppr.findall(qn("w:jc")):
        ppr.remove(old)
    jc = _el(f'<w:jc xmlns:w="{NS_W}" w:val="center"/>')
    # w:jc phải đứng trước textDirection/textAlignment/.../rPr/sectPr/pPrChange
    after = {qn(t) for t in ("w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                             "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")}
    for i, child in enumerate(list(ppr)):
        if child.tag in after:
            ppr.insert(i, jc)
            return
    ppr.append(jc)


def _set_position(run, half_points: int) -> None:
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = _el(f'<w:rPr xmlns:w="{NS_W}"/>')
        run.insert(0, rpr)
    for old in rpr.findall(qn("w:position")):
        rpr.remove(old)
    pos = _el(f'<w:position xmlns:w="{NS_W}" w:val="{half_points}"/>')
    after = {qn(t) for t in ("w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
                             "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                             "w:specVanish", "w:oMath")}
    for i, child in enumerate(list(rpr)):
        if child.tag in after:
            rpr.insert(i, pos)
            return
    rpr.append(pos)


def xml_attr(s: str) -> str:
    s = sanitize_xml_text(s)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;").replace("'", "&apos;"))


_xml_attr = xml_attr  # tên cũ
