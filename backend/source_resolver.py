"""`backend/source_resolver.py` — nạp dự án TeX đúng ngữ nghĩa trước mọi bước tiền xử lý.

Làm gì:
* đọc tệp với encoding xác định được (UTF-8/UTF-8 BOM/UTF-16 BOM hoặc inputenc khai báo);
  KHÔNG dùng errors='replace' để âm thầm thay byte lỗi;
* che chú thích, verbatim/lstlisting/minted/comment, `\\verb` và nhánh `\\iffalse` khi dò lệnh;
* chọn tệp chính (đường dẫn tương đối chính xác nếu người dùng chỉ định; nếu không thì xét
  `\\documentclass`, `\\begin{document}`, magic comment `% !TEX root` và đồ thị include);
* mở `\\input`, `\\include` (tôn trọng `\\includeonly`), `\\subfile`, `\\import`, `\\subimport`
  theo thư mục gốc của tài liệu chính như TeX; phát hiện vòng lặp, quá sâu, tệp thiếu,
  đường dẫn ra ngoài gói và lệnh động (đối số chứa macro) chưa hỗ trợ;
* giữ bản đồ dòng (dòng phẳng → tệp/dòng gốc) để báo cáo vị trí nguồn của từng khối.

Module chỉ đọc tệp; không ghi đè gì trong thư mục dự án.
"""
from __future__ import annotations

import codecs
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any

__all__ = [
    "SourceEncodingError", "MainSelectionError", "decode_tex_bytes", "mask_tex", "scan_project",
    "select_main", "flatten_project", "FlattenResult", "split_preamble", "find_matching_brace",
]

MAX_INCLUDE_DEPTH = 32
TEX_EXTS = (".tex", ".ltx", ".latex")


class SourceEncodingError(ValueError):
    pass


class MainSelectionError(ValueError):
    """Không chọn được tệp chính một cách chắc chắn. `candidates` để giao diện hiển thị cho người dùng chọn."""

    def __init__(self, message: str, candidates: list[dict[str, Any]] | None = None, code: str = "main_tex") -> None:
        super().__init__(message)
        self.candidates = candidates or []
        self.code = code


# ───────────────────────────── encoding ─────────────────────────────
_INPUTENC_RE = re.compile(rb"\\usepackage\s*\[([^\]]*)\]\s*\{inputenc\}")
_INPUTENC_CODECS = {
    "utf8": "utf-8", "utf8x": "utf-8", "latin1": "latin-1", "latin9": "iso8859-15", "latin2": "iso8859-2",
    "cp1252": "cp1252", "ansinew": "cp1252", "cp1250": "cp1250", "cp1258": "cp1258", "applemac": "mac_roman",
    "koi8-r": "koi8_r", "cp1251": "cp1251",
}


def decode_tex_bytes(data: bytes, name: str = "") -> tuple[str, str, list[str]]:
    """Trả về (văn bản, encoding, cảnh báo). Không đoán bừa: lỗi giải mã → SourceEncodingError."""
    warnings: list[str] = []
    if data.startswith(codecs.BOM_UTF8):
        try:
            return data[3:].decode("utf-8"), "utf-8-sig", warnings
        except UnicodeDecodeError as exc:
            raise SourceEncodingError(f"{name}: có BOM UTF-8 nhưng chứa byte không hợp lệ ở vị trí {exc.start}.") from exc
    if data.startswith(codecs.BOM_UTF16_LE) or data.startswith(codecs.BOM_UTF16_BE):
        try:
            return data.decode("utf-16"), "utf-16", warnings
        except UnicodeDecodeError as exc:
            raise SourceEncodingError(f"{name}: UTF-16 hỏng ở vị trí {exc.start}.") from exc
    try:
        return data.decode("utf-8"), "utf-8", warnings
    except UnicodeDecodeError as utf_exc:
        m = _INPUTENC_RE.search(data)
        declared = None
        if m:
            opts = [o.strip().lower() for o in m.group(1).decode("ascii", "ignore").split(",")]
            for o in opts:
                if o in _INPUTENC_CODECS:
                    declared = _INPUTENC_CODECS[o]
        if declared and declared != "utf-8":
            try:
                text = data.decode(declared)
                warnings.append(f"{name}: đọc theo inputenc khai báo ({declared}).")
                return text, declared, warnings
            except UnicodeDecodeError as exc:
                raise SourceEncodingError(
                    f"{name}: không giải mã được theo inputenc khai báo ({declared}) ở vị trí {exc.start}.") from exc
        if re.search(rb"\\usepackage\s*\[[^\]]*(tcvn|viscii|vps)[^\]]*\]\s*\{vietnam\}", data):
            raise SourceEncodingError(
                f"{name}: tệp dùng bảng mã tiếng Việt cũ (TCVN3/VISCII/VPS) — chưa hỗ trợ giải mã an toàn. "
                "Hãy lưu lại tệp ở UTF-8.")
        raise SourceEncodingError(
            f"{name}: không phải UTF-8 hợp lệ (byte lỗi ở vị trí {utf_exc.start}) và không có khai báo inputenc "
            "đủ tin cậy. Không thay thế byte lỗi một cách âm thầm; hãy lưu lại tệp ở UTF-8.")


# ───────────────────────────── masking ─────────────────────────────
_VERB_ENV_NAMES = ("verbatim", "verbatim*", "Verbatim", "Verbatim*", "BVerbatim", "LVerbatim", "lstlisting",
                   "minted", "comment", "filecontents", "filecontents*")
_BEGIN_VERB_RE = re.compile(r"\\begin\s*\{(" + "|".join(re.escape(n) for n in _VERB_ENV_NAMES) + r")\}")
_VERB_CMD_RE = re.compile(r"\\(?:verb|lstinline|mintinline\{[^}]*\})\*?(?:\[[^\]]*\])?([^A-Za-z\s*{])")
_IF_TOKEN_RE = re.compile(r"\\(if[A-Za-z@]*|fi|else|newif)(?![A-Za-z@])")


def mask_tex(text: str, *, keep_iffalse: bool = False) -> str:
    """Trả về chuỗi cùng độ dài; chú thích, verbatim, `\\verb`, nhánh `\\iffalse` được thay bằng khoảng
    trắng (giữ nguyên ký tự xuống dòng) để mọi phép dò lệnh/vị trí không đụng vào chúng."""
    out = list(text)
    n = len(text)
    i = 0

    def blank(a: int, b: int) -> None:
        for k in range(max(0, a), min(n, b)):
            if out[k] != "\n":
                out[k] = " "

    while i < n:
        c = text[i]
        if c == "\\":
            m = _BEGIN_VERB_RE.match(text, i)
            if m:
                end_tok = "\\end{" + m.group(1) + "}"
                j = text.find(end_tok, m.end())
                j = n if j < 0 else j + len(end_tok)
                blank(i, j)
                i = j
                continue
            m = _VERB_CMD_RE.match(text, i)
            if m:
                d = m.group(1)
                j = text.find(d, m.end())
                nl = text.find("\n", m.end())
                if j < 0 or (0 <= nl < j):
                    j = nl if nl >= 0 else n
                else:
                    j += 1
                blank(i, j)
                i = j
                continue
            if not keep_iffalse and text.startswith("\\iffalse", i) and not (i + 8 < n and text[i + 8].isalpha()):
                j = _skip_iffalse(text, i + 8)
                blank(i, j)
                i = j
                continue
            i += 2
            continue
        if c == "%":
            j = text.find("\n", i)
            j = n if j < 0 else j
            blank(i, j)
            i = j
            continue
        i += 1
    return "".join(out)


def _skip_iffalse(text: str, pos: int) -> int:
    """Từ sau `\\iffalse`, trả vị trí kết thúc nhánh sai (sau `\\fi` khớp, hoặc ngay sau `\\else`)."""
    depth = 1
    prev_newif = False
    for m in _IF_TOKEN_RE.finditer(text, pos):
        # bỏ qua token nằm trong chú thích
        line_start = text.rfind("\n", 0, m.start()) + 1
        if _in_comment(text, line_start, m.start()):
            continue
        tok = m.group(1)
        if tok == "newif":
            prev_newif = True
            continue
        if tok.startswith("if"):
            if prev_newif:
                prev_newif = False
                continue
            depth += 1
        elif tok == "fi":
            depth -= 1
            if depth == 0:
                return m.end()
        elif tok == "else" and depth == 1:
            return m.end()
        prev_newif = False
    return len(text)


def _in_comment(text: str, line_start: int, pos: int) -> bool:
    k = line_start
    while k < pos:
        if text[k] == "\\":
            k += 2
            continue
        if text[k] == "%":
            return True
        k += 1
    return False


def find_matching_brace(text: str, open_idx: int) -> int:
    """`text[open_idx] == '{'` → chỉ số của `}` khớp (bỏ qua ký tự thoát), hoặc -1."""
    if open_idx >= len(text) or text[open_idx] != "{":
        return -1
    depth = 0
    i = open_idx
    while i < len(text):
        c = text[i]
        if c == "\\":
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def split_preamble(text: str) -> tuple[str, str, str]:
    """(preamble, body, tail) theo `\\begin{document}` / `\\end{document}` (trên bản đã che)."""
    masked = mask_tex(text)
    mb = re.search(r"\\begin\s*\{document\}", masked)
    if not mb:
        return "", text, ""
    me = re.search(r"\\end\s*\{document\}", masked[mb.end():])
    body_end = mb.end() + me.start() if me else len(text)
    return text[: mb.start()], text[mb.end(): body_end], text[body_end:]


# ───────────────────────────── include parsing ─────────────────────────────
_INCLUDE_CMD_RE = re.compile(r"\\(input|include|subfile|import|subimport|subinputfrom|inputfrom|includefrom|subincludefrom|includeonly)(?![A-Za-z@])\*?")
_DOCCLASS_RE = re.compile(r"\\documentclass\s*(?:\[([^\]]*)\])?\s*\{([^}]*)\}")
_BEGIN_DOC_RE = re.compile(r"\\begin\s*\{document\}")
_MAGIC_ROOT_RE = re.compile(r"^\s*%\s*!\s*TEX\s+root\s*=\s*(.+?)\s*$", re.I | re.M)
_GRAPHICSPATH_RE = re.compile(r"\\graphicspath\s*\{")


@dataclass
class IncludeRef:
    command: str
    args: list[str]
    start: int
    end: int
    dynamic: bool = False


def _read_group(text: str, pos: int) -> tuple[str | None, int]:
    k = pos
    while k < len(text) and text[k] in " \t\r\n":
        k += 1
    if k < len(text) and text[k] == "{":
        e = find_matching_brace(text, k)
        if e < 0:
            return None, pos
        return text[k + 1: e], e + 1
    return None, pos


def find_includes(text: str, masked: str | None = None) -> list[IncludeRef]:
    masked = mask_tex(text) if masked is None else masked
    refs: list[IncludeRef] = []
    for m in _INCLUDE_CMD_RE.finditer(masked):
        cmd = m.group(1)
        pos = m.end()
        nargs = 2 if cmd in ("import", "subimport", "subinputfrom", "inputfrom", "includefrom", "subincludefrom") else 1
        args: list[str] = []
        p = pos
        ok = True
        for _ in range(nargs):
            a, p2 = _read_group(masked, p)
            if a is None:
                ok = False
                break
            args.append(a)
            p = p2
        if not ok:
            if cmd == "input":
                # dạng TeX gốc: \input tên-tệp (kết thúc ở khoảng trắng)
                mm = re.match(r"[ \t]+([^\s{}\\%]+)", masked[pos:])
                if mm:
                    refs.append(IncludeRef(cmd, [mm.group(1)], m.start(), pos + mm.end()))
                    continue
            refs.append(IncludeRef(cmd, [], m.start(), m.end(), dynamic=True))
            continue
        dyn = any("\\" in a or "#" in a for a in args)
        refs.append(IncludeRef(cmd, [a.strip() for a in args], m.start(), p, dynamic=dyn))
    return refs


def _norm_rel(p: str) -> str:
    return unicodedata.normalize("NFC", str(PurePosixPath(p.replace("\\", "/"))))


def _safe_join(root: Path, base_rel: str, name: str) -> Path | None:
    """Ghép đường dẫn trong phạm vi `root`; trả None nếu tuyệt đối hoặc thoát ra ngoài."""
    name = name.strip().strip('"')
    if not name or name.startswith("/") or re.match(r"^[A-Za-z]:", name) or name.startswith("~"):
        return None
    cand = (root / base_rel / name).resolve()
    rr = root.resolve()
    if cand != rr and rr not in cand.parents:
        return None
    return cand


def _with_tex_ext(p: Path) -> list[Path]:
    if p.suffix.lower() in TEX_EXTS:
        return [p]
    return [p.with_name(p.name + ".tex"), p]


def _ci_lookup(p: Path) -> Path | None:
    """Tìm tệp không phân biệt hoa/thường (gói tạo trên Windows chạy trên Linux)."""
    parent = p.parent
    if not parent.is_dir():
        return None
    low = p.name.casefold()
    for child in parent.iterdir():
        if child.name.casefold() == low and child.is_file():
            return child
    return None


# ───────────────────────────── scan & select ─────────────────────────────
def _relpath(root: Path, p: Path) -> str:
    return _norm_rel(str(p.resolve().relative_to(root.resolve())))


def scan_project(root: Path) -> dict[str, Any]:
    root = root.resolve()
    files: dict[str, dict[str, Any]] = {}
    for p in sorted(root.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in TEX_EXTS:
            continue
        rel = _relpath(root, p)
        info: dict[str, Any] = {"rel": rel}
        try:
            text, enc, _w = decode_tex_bytes(p.read_bytes(), rel)
        except SourceEncodingError as exc:
            info.update({"error": str(exc), "documentclass": None, "begin_document": False, "includes": []})
            files[rel] = info
            continue
        masked = mask_tex(text)
        dm = _DOCCLASS_RE.search(masked)
        info["encoding"] = enc
        info["documentclass"] = dm.group(2).strip() if dm else None
        info["subfiles_child"] = bool(dm and dm.group(2).strip() == "subfiles")
        info["begin_document"] = bool(_BEGIN_DOC_RE.search(masked))
        mr = _MAGIC_ROOT_RE.search(text[:4000])
        info["magic_root"] = mr.group(1).strip() if mr else None
        incs = []
        for ref in find_includes(text, masked):
            if ref.command == "includeonly" or ref.dynamic or not ref.args:
                continue
            name = ref.args[-1]
            base = ref.args[0] if ref.command in ("import", "subimport") else ""
            incs.append(_norm_rel(str(PurePosixPath(base) / name)))
        info["includes"] = incs
        files[rel] = info
    # đồ thị include (so khớp theo đường dẫn tương đối không đuôi)
    def key(r: str) -> str:
        for e in TEX_EXTS:
            if r.lower().endswith(e):
                return r[: -len(e)]
        return r
    by_key = {key(r).casefold(): r for r in files}
    included_by: dict[str, set[str]] = {r: set() for r in files}
    for r, info in files.items():
        rdir = str(PurePosixPath(r).parent)
        for inc in info.get("includes", []):
            for cand in (inc, _norm_rel(str(PurePosixPath(rdir) / inc))):
                tgt = by_key.get(key(cand).casefold())
                if tgt and tgt != r:
                    included_by[tgt].add(r)
    for r, info in files.items():
        info["included_by"] = sorted(included_by[r])
    return {"root": str(root), "files": files}


def _candidate_list(scan: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for r, info in scan["files"].items():
        if info.get("documentclass") and info.get("begin_document") and not info.get("subfiles_child"):
            out.append({"path": r, "documentclass": info["documentclass"], "included_by": info.get("included_by", []),
                        "magic_root": info.get("magic_root")})
    return out


def select_main(root: Path, requested: str | None = None) -> dict[str, Any]:
    """Chọn tệp chính. Trả về {"main": rel, "how": ..., "candidates": [...]} hoặc ném MainSelectionError."""
    root = root.resolve()
    scan = scan_project(root)
    files = scan["files"]
    cands = _candidate_list(scan)
    all_tex = sorted(files)
    if requested and requested.strip():
        req = _norm_rel(requested.strip().lstrip("./"))
        if req in files:
            return {"main": req, "how": "user_exact_path", "candidates": cands, "scan": scan}
        ci = [r for r in all_tex if r.casefold() == req.casefold()]
        if len(ci) == 1:
            return {"main": ci[0], "how": "user_path_case_insensitive", "candidates": cands, "scan": scan}
        if "/" not in req:
            same = [r for r in all_tex if PurePosixPath(r).name.casefold() == req.casefold()]
            if len(same) == 1:
                return {"main": same[0], "how": "user_basename_unique", "candidates": cands, "scan": scan}
            if len(same) > 1:
                raise MainSelectionError(
                    f"Tên '{requested}' khớp nhiều tệp: {', '.join(same)}. Hãy ghi đường dẫn tương đối đầy đủ.",
                    [{"path": s} for s in same], code="main_tex_ambiguous")
        raise MainSelectionError(
            f"Không có tệp '{requested}' trong gói. Các tệp .tex có: {', '.join(all_tex) or '(không có)'}.",
            [{"path": s} for s in all_tex], code="main_tex_not_found")
    if not all_tex:
        raise MainSelectionError("Gói không có tệp .tex nào.", [], code="no_tex")
    # magic comment trỏ tới tệp chính
    roots_from_magic = set()
    for r, info in files.items():
        mr = info.get("magic_root")
        if mr:
            tgt = _norm_rel(str(PurePosixPath(r).parent / mr))
            tgt = str(PurePosixPath(tgt))
            # chuẩn hoá ../
            parts: list[str] = []
            for part in tgt.split("/"):
                if part == "..":
                    if parts:
                        parts.pop()
                elif part not in ("", "."):
                    parts.append(part)
            t = "/".join(parts)
            if t in files:
                roots_from_magic.add(t)
    top = [c for c in cands if not c["included_by"]]
    if len(roots_from_magic) == 1:
        main = next(iter(roots_from_magic))
        return {"main": main, "how": "magic_root_comment", "candidates": cands, "scan": scan}
    if len(top) == 1:
        return {"main": top[0]["path"], "how": "single_root_document", "candidates": cands, "scan": scan}
    if len(cands) == 1:
        return {"main": cands[0]["path"], "how": "single_document", "candidates": cands, "scan": scan}
    if not cands and len(all_tex) == 1:
        return {"main": all_tex[0], "how": "only_tex_file_without_documentclass", "candidates": [], "scan": scan}
    choices = top or cands or [{"path": r} for r in all_tex]
    raise MainSelectionError(
        "Gói có nhiều tài liệu chính hợp lệ: " + ", ".join(c["path"] for c in choices)
        + ". Hãy chọn tệp chính (main_tex) bằng đường dẫn tương đối.", choices, code="main_tex_ambiguous")


# ───────────────────────────── flatten ─────────────────────────────
@dataclass
class FlattenResult:
    main: str
    text: str
    line_map: list[tuple[str, int]] = field(default_factory=list)
    files: list[dict[str, Any]] = field(default_factory=list)
    edges: list[dict[str, Any]] = field(default_factory=list)
    graphicspath: list[str] = field(default_factory=list)
    base_dirs: dict[str, list[str]] = field(default_factory=dict)
    issues: list[dict[str, Any]] = field(default_factory=list)

    def where(self, flat_offset: int) -> dict[str, Any]:
        line = self.text.count("\n", 0, flat_offset)
        if 0 <= line < len(self.line_map):
            f, ln = self.line_map[line]
            return {"file": f, "line": ln}
        return {"file": self.main, "line": None}

    def summary(self) -> dict[str, Any]:
        return {
            "main": self.main,
            "files": self.files,
            "includes": self.edges,
            "graphicspath": self.graphicspath,
            "issues": self.issues,
        }


def flatten_project(root: Path, main_rel: str, *, max_depth: int = MAX_INCLUDE_DEPTH) -> FlattenResult:
    root = root.resolve()
    main_path = (root / main_rel).resolve()
    res = FlattenResult(main=main_rel, text="")
    pieces: list[tuple[str, str, int]] = []      # (văn bản, tệp, dòng bắt đầu)
    includeonly: set[str] | None = None
    main_dir_rel = str(PurePosixPath(main_rel).parent)
    if main_dir_rel == ".":
        main_dir_rel = ""

    def issue(sev: str, code: str, msg: str, file: str, line: int | None) -> None:
        res.issues.append({"severity": sev, "code": code, "message": msg, "file": file, "line": line})

    def emit(text: str, file: str, line: int) -> None:
        if text:
            pieces.append((text, file, line))

    def resolve_child(name: str, bases: list[str], cmd: str) -> tuple[Path | None, str]:
        tried: list[str] = []
        for i_b, base in enumerate(bases):
            target = _safe_join(root, base, name)
            if target is None:
                return None, "outside"
            for cand in _with_tex_ext(target):
                tried.append(str(cand))
                if cand.is_file():
                    return cand, "ok" if i_b == 0 else "ok_fallback_base"
                ci = _ci_lookup(cand)
                if ci is not None:
                    return ci, "ok_case_insensitive"
        return None, "missing"

    def expand(path: Path, bases: list[str], stack: tuple[str, ...], depth: int, body_only: bool) -> None:
        nonlocal includeonly
        rel = _relpath(root, path)
        try:
            text, enc, warns = decode_tex_bytes(path.read_bytes(), rel)
        except SourceEncodingError as exc:
            issue("error", "encoding", str(exc), rel, None)
            emit(f"\\par\\textbf{{[L2W: KHÔNG ĐỌC ĐƯỢC TỆP {rel} — lỗi encoding]}}\\par\n", rel, 1)
            return
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        for w in warns:
            issue("warning", "encoding_declared", w, rel, None)
        res.files.append({"path": rel, "encoding": enc, "lines": text.count("\n") + 1, "depth": depth})
        own_dir = str(PurePosixPath(rel).parent)
        own_dir = "" if own_dir == "." else own_dir
        res.base_dirs.setdefault(rel, list(dict.fromkeys(bases + [own_dir])))
        start_off = 0
        if body_only:
            masked0 = mask_tex(text)
            mb = _BEGIN_DOC_RE.search(masked0)
            if mb:
                me = re.search(r"\\end\s*\{document\}", masked0[mb.end():])
                start_off = mb.end()
                text_end = mb.end() + me.start() if me else len(text)
                text = text[:text_end]
        masked = mask_tex(text)
        if includeonly is None and depth == 0:
            for ref in find_includes(text, masked):
                if ref.command == "includeonly" and ref.args:
                    includeonly = {_norm_rel(a.strip()) for a in ref.args[0].split(",") if a.strip()}
        for gm in _GRAPHICSPATH_RE.finditer(masked):
            ob = gm.end() - 1
            cb = find_matching_brace(text, ob)
            if cb > 0:
                for d in re.findall(r"\{([^{}]*)\}", text[ob + 1: cb]):
                    d = d.strip()
                    if d and d not in res.graphicspath:
                        res.graphicspath.append(d)
        pos = start_off
        for ref in find_includes(text, masked):
            if ref.start < start_off:
                continue
            line_no = text.count("\n", 0, ref.start) + 1
            if ref.command == "includeonly":
                emit(text[pos: ref.start], rel, text.count("\n", 0, pos) + 1)
                pos = ref.end
                continue
            emit(text[pos: ref.start], rel, text.count("\n", 0, pos) + 1)
            raw_cmd = text[ref.start: ref.end]
            edge = {"from": rel, "line": line_no, "command": ref.command, "target": "/".join(ref.args),
                    "status": "ok", "resolved": None}
            res.edges.append(edge)
            if ref.dynamic or not ref.args:
                edge["status"] = "dynamic_unsupported"
                issue("warning", "include_dynamic",
                      f"Lệnh {raw_cmd.strip()} có đối số động/không đọc được → giữ nguyên, chưa nạp.", rel, line_no)
                emit(raw_cmd, rel, line_no)
                pos = ref.end
                continue
            name = ref.args[-1]
            child_bases = bases
            body_only_child = False
            if ref.command in ("import", "inputfrom", "includefrom"):
                child_bases = [_norm_rel(ref.args[0]).rstrip("/")]
            elif ref.command in ("subimport", "subinputfrom", "subincludefrom"):
                child_bases = [_norm_rel(str(PurePosixPath(bases[0] or ".") / ref.args[0])).rstrip("/")]
            elif ref.command == "subfile":
                body_only_child = True
            if ref.command == "include" and includeonly is not None and _norm_rel(name) not in includeonly:
                edge["status"] = "excluded_by_includeonly"
                pos = ref.end
                continue
            search = list(dict.fromkeys(child_bases + bases + [own_dir]))
            child, how = resolve_child(name, search, ref.command)
            if child is None:
                edge["status"] = "outside_project" if how == "outside" else "missing"
                msg = (f"Đường dẫn '{name}' trỏ ra ngoài gói/tuyệt đối → không nạp." if how == "outside"
                       else f"Không tìm thấy tệp '{name}' ({ref.command}).")
                issue("error", "include_" + edge["status"], msg, rel, line_no)
                emit(f"\\par\\textbf{{[L2W: THIẾU TỆP {name}]}}\\par\n", rel, line_no)
                pos = ref.end
                continue
            crel = _relpath(root, child)
            edge["resolved"] = crel
            if how != "ok":
                edge["status"] = how
                issue("warning", "include_" + how, f"'{name}' → {crel} ({how}).", rel, line_no)
            if crel in stack or crel == rel:
                edge["status"] = "cycle"
                issue("error", "include_cycle", "Vòng lặp include: " + " → ".join(stack + (rel, crel)), rel, line_no)
                emit(f"\\par\\textbf{{[L2W: VÒNG LẶP INCLUDE {crel}]}}\\par\n", rel, line_no)
                pos = ref.end
                continue
            if depth + 1 > max_depth:
                edge["status"] = "too_deep"
                issue("error", "include_too_deep", f"Include lồng quá {max_depth} cấp tại {crel}.", rel, line_no)
                emit(f"\\par\\textbf{{[L2W: INCLUDE QUÁ SÂU {crel}]}}\\par\n", rel, line_no)
                pos = ref.end
                continue
            sub_bases = list(child_bases)
            if body_only_child:
                cdir = str(PurePosixPath(crel).parent)
                sub_bases = [("" if cdir == "." else cdir)] + [b for b in bases if b != cdir]
            emit("\n" + ("\\clearpage\n" if ref.command == "include" else ""), rel, line_no)
            expand(child, sub_bases, stack + (rel,), depth + 1, body_only_child)
            emit("\n" + ("\\clearpage\n" if ref.command == "include" else ""), rel, line_no)
            pos = ref.end
        emit(text[pos:], rel, text.count("\n", 0, pos) + 1)

    expand(main_path, [main_dir_rel], (), 0, False)
    flat_parts: list[str] = []
    line_map: list[tuple[str, int]] = []
    cur_line_empty = True          # dòng phẳng hiện tại chưa có ký tự nào
    for text, f, ln in pieces:
        segs = text.split("\n")
        if not line_map:
            line_map.append((f, ln))
        elif cur_line_empty:
            line_map[-1] = (f, ln)  # dòng trống do mảnh trước kết thúc bằng xuống dòng → thuộc mảnh này
        for k in range(1, len(segs)):
            line_map.append((f, ln + k))
        cur_line_empty = segs[-1] == "" if len(segs) > 1 else (cur_line_empty and segs[0] == "")
        flat_parts.append(text)
    res.text = "".join(flat_parts)
    if not line_map:
        line_map.append((main_rel, 1))
    res.line_map = line_map
    return res
