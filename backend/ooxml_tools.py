"""`backend/ooxml_tools.py` — kiểm kê và kiểm tra gói DOCX CUỐI CÙNG (sau mọi hậu xử lý).

* `docx_inventory()`   : đếm lại công thức, bảng theo vai trò, hình DrawingML, ảnh LOGIC (SVG+PNG
                         của cùng một hình = 1), chỗ đánh dấu lộ ra, thông báo L2W, lệnh TeX rò rỉ…
* `validate_package()` : [Content_Types], quan hệ (rels) và đích, r:embed/r:id, docPr id trùng,
                         XML hợp lệ của mọi part, python-docx mở lại được.
* `normalize_xml_order()`: sắp lại thứ tự con của pPr/rPr/tcPr/tblPr/sectPr theo lược đồ.
* `set_table_role()/get_table_role()`: vai trò bảng ghi bằng `w:tblDescription` (metadata thật trong
  tblPr, Word giữ lại) thay vì đoán theo tên media.
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

from lxml import etree

__all__ = ["docx_inventory", "validate_package", "normalize_xml_order", "set_table_role", "get_table_role",
           "docx_text_stream", "ROLE_PREFIX", "LEAK_PATTERNS", "W", "M"]

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "m": M, "a": A, "wp": WP, "pic": PIC, "wps": WPS, "wpg": WPG, "r": R, "asvg": ASVG}

ROLE_PREFIX = "l2w:role="

# Chỗ đánh dấu nội bộ không được lộ ra Word cuối cùng.
LEAK_PATTERNS = {
    "truth_marker": re.compile(r"ZLZTRUTHZ\d{5}ZEND"),
    "image_marker": re.compile(r"ZLZIMGZ\d{5}ZEND"),
    "layout_sentinel": re.compile(r"ZLZLAYZ\d{5}Z(?:BEG|MID|FIN)"),
    "engine_block_token": re.compile(r"@@L2W_BLOCK_\d{4}@@"),
    "verbatim_marker": re.compile(r"ZLZVERBZ\d{5}ZEND"),
}
L2W_NOTICE_RE = re.compile(r"\[L2W:[^\]]*\]")
TEX_RESIDUE_RE = re.compile(r"\\[A-Za-z]{2,}")


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


# ───────────────────────────── vai trò bảng ─────────────────────────────
_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription", "tblPrChange"]


def set_table_role(tbl, role: str, **extra: Any) -> None:
    """Ghi vai trò bảng (`mcq`, `true_false`, `variation`, `data`, `layout_immini`, `layout_mcq_pdf`…)."""
    tblPr = tbl.find(_q("w", "tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, _q("w", "tblPr"))
        tbl.remove(tblPr)
        tbl.insert(0, tblPr)
    for old in tblPr.findall(_q("w", "tblDescription")):
        tblPr.remove(old)
    val = ROLE_PREFIX + role + "".join(f";{k}={v}" for k, v in extra.items() if v is not None)
    el = etree.Element(_q("w", "tblDescription"))
    el.set(_q("w", "val"), val)
    # chèn đúng thứ tự lược đồ
    idx = _TBLPR_ORDER.index("tblDescription")
    succ = {_q("w", t) for t in _TBLPR_ORDER[idx + 1:]}
    for i, child in enumerate(list(tblPr)):
        if child.tag in succ:
            tblPr.insert(i, el)
            return
    tblPr.append(el)


def get_table_role(tbl) -> dict[str, str] | None:
    tblPr = tbl.find(_q("w", "tblPr"))
    if tblPr is None:
        return None
    d = tblPr.find(_q("w", "tblDescription"))
    if d is None:
        return None
    val = d.get(_q("w", "val")) or ""
    if not val.startswith(ROLE_PREFIX):
        return None
    parts = val[len(ROLE_PREFIX):].split(";")
    out = {"role": parts[0]}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            out[k] = v
    return out


# ───────────────────────────── kiểm kê ─────────────────────────────
def _text_of(el) -> str:
    return "".join(t.text or "" for t in el.iter(_q("w", "t"), _q("m", "t")))


def _heuristic_table_role(tbl) -> str:
    rows = tbl.findall(_q("w", "tr"))
    cells_text = [_text_of(tc).strip() for tr in rows for tc in tr.findall(_q("w", "tc"))]
    non_empty = [c for c in cells_text if c]
    if len(non_empty) == 4 and all(non_empty[i].startswith(l) for i, l in enumerate(("A.", "B.", "C.", "D."))):
        return "mcq"
    xml = etree.tostring(tbl, encoding="unicode")
    if "VarArrow_" in xml:
        return "variation"
    if len(rows) == 1 and len(rows[0].findall(_q("w", "tc"))) == 2 and (
            "EditableMathShapeGroup_" in xml or "<pic:pic" in xml or "wpg:wgp" in xml):
        return "layout_immini"
    return "data"


def _table_is_empty(tbl) -> bool:
    if tbl.findall(".//" + _q("w", "drawing")) or tbl.findall(".//" + _q("m", "oMath")):
        return False
    return not _text_of(tbl).strip()


def docx_text_stream(root) -> list[dict[str, Any]]:
    """Dòng chảy nội dung theo thứ tự tài liệu: mỗi đoạn → {"text", "math": [chuỗi m:t của từng oMath]}."""
    out = []
    for p in root.iter(_q("w", "p")):
        parts = []
        maths = []
        for el in p.iter(_q("w", "t"), _q("m", "oMath")):
            if el.tag == _q("w", "t"):
                # bỏ w:t nằm trong oMath (không có) hoặc trong đoạn lồng (textbox) — đoạn lồng tự được duyệt
                anc = el.getparent()
                while anc is not None and anc.tag != _q("w", "p"):
                    anc = anc.getparent()
                if anc is p:
                    parts.append(el.text or "")
            else:
                anc = el.getparent()
                while anc is not None and anc.tag != _q("w", "p"):
                    anc = anc.getparent()
                if anc is p:
                    maths.append("".join(t.text or "" for t in el.iter(_q("m", "t"))))
        out.append({"text": "".join(parts), "math": maths})
    return out


def docx_inventory(docx_path: Path) -> dict[str, Any]:
    inv: dict[str, Any] = {"ok": False}
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        doc = etree.fromstring(z.read("word/document.xml"))
        extra_roots = []
        for n in names:
            if re.match(r"word/(header|footer|footnotes|endnotes)\d*\.xml$", n):
                try:
                    extra_roots.append((n, etree.fromstring(z.read(n))))
                except etree.XMLSyntaxError:
                    pass
    body = doc.find(_q("w", "body"))
    roots = [("word/document.xml", doc)] + extra_roots
    omath_all = sum(len(r.findall(".//" + _q("m", "oMath"))) for _, r in roots)
    tables = doc.findall(".//" + _q("w", "tbl"))
    roles: Counter = Counter()
    role_src: Counter = Counter()
    empty_tables = 0
    table_list = []
    for i, tbl in enumerate(tables):
        meta = get_table_role(tbl)
        role = meta["role"] if meta else _heuristic_table_role(tbl)
        roles[role] += 1
        role_src["metadata" if meta else "heuristic"] += 1
        empty = _table_is_empty(tbl)
        empty_tables += int(empty)
        rows = tbl.findall(_q("w", "tr"))
        table_list.append({"index": i, "role": role, "role_source": "metadata" if meta else "heuristic",
                           "rows": len(rows), "max_cols": max((len(r.findall(_q("w", "tc"))) for r in rows), default=0),
                           "empty": empty, **({k: v for k, v in meta.items() if k != "role"} if meta else {})})
    drawings = doc.findall(".//" + _q("w", "drawing"))
    pics = 0
    svg_backed = 0
    groups = 0
    image_names: list[str] = []
    for d in drawings:
        gd = d.find(".//" + _q("a", "graphicData"))
        uri = gd.get("uri") if gd is not None else ""
        if uri and uri.endswith("/picture"):
            pics += 1
            if d.find(".//" + _q("asvg", "svgBlip")) is not None:
                svg_backed += 1
            dp = d.find(".//" + _q("wp", "docPr"))
            image_names.append((dp.get("name") or "") if dp is not None else "")
        elif uri and "wordprocessingGroup" in uri:
            groups += 1
    all_text_nodes = [t.text or "" for _, r in roots for t in r.iter(_q("w", "t"), _q("m", "t"))]
    joined = "\n".join(all_text_nodes)
    leaks = {k: len(p.findall(joined)) for k, p in LEAK_PATTERNS.items()}
    w_text = "\n".join(t.text or "" for t in doc.iter(_q("w", "t")))
    residue = TEX_RESIDUE_RE.findall(w_text)
    inv.update({
        "ok": True,
        "paragraphs": len(body.findall(".//" + _q("w", "p"))) if body is not None else 0,
        "text_chars": len(w_text),
        "equation_nodes": omath_all,
        "omml_paragraphs": len(doc.findall(".//" + _q("m", "oMathPara"))),
        "word_tables": len(tables),
        "tables_by_role": dict(roles),
        "table_role_source": dict(role_src),
        "empty_tables": empty_tables,
        "tables": table_list[:200],
        "drawingml_groups": groups,
        "drawingml_shapes": len(doc.findall(".//" + _q("wps", "wsp"))),
        "logical_images": pics,
        "svg_backed_images": svg_backed,
        "image_names": image_names[:500],
        "media_parts": len([n for n in names if n.startswith("word/media/") and not n.endswith("/")]),
        "marker_leaks": leaks,
        "marker_leaks_total": sum(leaks.values()),
        "l2w_notices": L2W_NOTICE_RE.findall(joined)[:100],
        "tex_residue_count": len(residue),
        "tex_residue_samples": sorted(set(residue))[:30],
    })
    return inv


# ───────────────────────────── kiểm tra gói ─────────────────────────────
def _rels_path_for(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def validate_package(docx_path: Path, *, try_python_docx: bool = True) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []
    try:
        z = zipfile.ZipFile(docx_path)
    except zipfile.BadZipFile as exc:
        return {"ok": False, "errors": [f"Không phải gói ZIP hợp lệ: {exc}"], "warnings": []}
    with z:
        bad = z.testzip()
        if bad:
            errors.append(f"Phần hỏng CRC trong gói: {bad}")
        names = set(z.namelist())
        if "[Content_Types].xml" not in names:
            errors.append("Thiếu [Content_Types].xml")
            return {"ok": False, "errors": errors, "warnings": warnings}
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        defaults = {d.get("Extension", "").lower(): d.get("ContentType") for d in ct.findall(f"{{{CT}}}Default")}
        overrides = {o.get("PartName", ""): o.get("ContentType") for o in ct.findall(f"{{{CT}}}Override")}
        for n in sorted(names):
            if n.endswith("/") or n == "[Content_Types].xml":
                continue
            ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
            if "/" + n not in overrides and ext not in defaults:
                errors.append(f"Part '{n}' không có content type")
        for pn in overrides:
            if pn.lstrip("/") not in names:
                errors.append(f"[Content_Types] khai báo part không tồn tại: {pn}")
        parsed: dict[str, Any] = {}
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    parsed[n] = etree.fromstring(z.read(n))
                except etree.XMLSyntaxError as exc:
                    errors.append(f"XML hỏng trong {n}: {exc}")
        # quan hệ
        rel_ids: dict[str, dict[str, str]] = {}
        for n, root in parsed.items():
            if not n.endswith(".rels"):
                continue
            src_dir = posixpath.dirname(posixpath.dirname(n))
            src_part = posixpath.join(src_dir, posixpath.basename(n)[:-5])
            ids: dict[str, str] = {}
            for rel in root.findall(f"{{{PR}}}Relationship"):
                rid = rel.get("Id", "")
                if rid in ids:
                    errors.append(f"{n}: Id quan hệ trùng {rid}")
                ids[rid] = rel.get("Type", "")
                if rel.get("TargetMode") == "External":
                    continue
                tgt = rel.get("Target", "")
                full = posixpath.normpath(tgt.lstrip("/") if tgt.startswith("/") else posixpath.join(src_dir, tgt))
                if full not in names:
                    errors.append(f"{n}: quan hệ {rid} trỏ tới part không tồn tại '{full}'")
            rel_ids[src_part] = ids
        # tham chiếu r:* trong các part XML
        for n, root in parsed.items():
            if n.endswith(".rels") or not n.startswith("word/"):
                continue
            ids = rel_ids.get(n, {})
            for el in root.iter():
                for attr, val in el.attrib.items():
                    if attr.startswith(f"{{{R}}}") and attr.split("}")[1] in ("embed", "id", "link", "pict"):
                        if val and val not in ids:
                            errors.append(f"{n}: tham chiếu {attr.split('}')[1]}='{val}' không có trong quan hệ")
        # docPr id trùng
        docpr_ids: Counter = Counter()
        for n, root in parsed.items():
            if n.startswith("word/") and not n.endswith(".rels"):
                for dp in root.iter(_q("wp", "docPr")):
                    docpr_ids[dp.get("id")] += 1
        dups = [k for k, v in docpr_ids.items() if v > 1]
        if dups:
            errors.append(f"docPr id trùng: {sorted(dups)[:20]}")
        # media không được tham chiếu (cảnh báo)
        referenced = set()
        for n, root in parsed.items():
            if n.endswith(".rels"):
                src_dir = posixpath.dirname(posixpath.dirname(n))
                for rel in root.findall(f"{{{PR}}}Relationship"):
                    if rel.get("TargetMode") != "External":
                        referenced.add(posixpath.normpath(posixpath.join(src_dir, rel.get("Target", ""))))
        orphan = [n for n in names if n.startswith("word/media/") and n not in referenced]
        if orphan:
            warnings.append(f"{len(orphan)} tệp media không được quan hệ nào tham chiếu")
    if try_python_docx and not errors:
        try:
            import docx
            docx.Document(str(docx_path))
        except Exception as exc:
            errors.append(f"python-docx không mở lại được: {type(exc).__name__}: {exc}")
    return {"ok": not errors, "errors": errors[:100], "warnings": warnings}


# ───────────────────────────── thứ tự phần tử ─────────────────────────────
_SEQS = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
            "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
            "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w",
            "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
            "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
             "tcPrChange"],
    "tblPr": _TBLPR_ORDER,
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
             "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
               "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
               "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    "tcBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"],
    "tblBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV"],
    "tcMar": ["top", "start", "left", "bottom", "end", "right"],
}


def normalize_xml_order(root) -> int:
    """Sắp lại con của các phần tử thuộc tính theo thứ tự lược đồ. Chỉ đụng tới phần tử mà mọi con đều
    có trong danh sách đã biết (phần tử lạ → để nguyên, không đoán)."""
    fixed = 0
    seqs = {_q("w", k): [_q("w", t) for t in v] for k, v in _SEQS.items()}
    for el in list(root.iter(*seqs.keys())):
        seq = seqs[el.tag]
        kids = [c for c in el if isinstance(c.tag, str)]
        if not kids or any(c.tag not in seq for c in kids):
            continue
        # phần tử thuộc tính trùng tên là không hợp lệ → giữ cái CUỐI (người ghi sau muốn ghi đè)
        last: dict[str, Any] = {}
        for c in kids:
            last[c.tag] = c
        wanted = sorted(last.values(), key=lambda c: seq.index(c.tag))
        if wanted != kids:
            for c in kids:
                el.remove(c)
            for c in wanted:
                el.append(c)
            fixed += 1
    return fixed
