"""`backend/job_workspace.py` — thư mục riêng cho từng job, tên tệp an toàn, nhận diện loại
tệp theo NỘI DUNG và giải nén ZIP có kiểm soát.

Nguyên tắc:
* bytes người dùng tải lên được lưu nguyên trong `input/` và KHÔNG bao giờ bị ghi đè;
  mọi xử lý diễn ra trên bản sao trong `work/`; kết quả nằm trong `out/`;
* mỗi job có thư mục riêng (id ngẫu nhiên), không đọc tệp tạm/ảnh/kết quả của job khác;
* không nhận diện loại tệp chỉ bằng phần mở rộng.
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import stat
import unicodedata
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any

__all__ = [
    "JobWorkspace", "sanitize_upload_name", "sniff_kind", "safe_extract_zip", "ZipSafetyError",
    "ExtractResult", "JOB_ID_RE", "sha256_file", "InputKindError",
]

JOB_ID_RE = re.compile(r"^[0-9a-f]{8,32}$")
_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
_FORBIDDEN_CHARS = re.compile(r'[<>:"|?*\x00-\x1f\x7f]')

# Giới hạn mặc định (có thể đổi qua biến môi trường).
MAX_UPLOAD_BYTES = int(os.environ.get("L2W_MAX_UPLOAD_MB", "100")) * 1024 * 1024
MAX_ZIP_FILES = int(os.environ.get("L2W_MAX_ZIP_FILES", "2000"))
MAX_ZIP_TOTAL_BYTES = int(os.environ.get("L2W_MAX_ZIP_TOTAL_MB", "400")) * 1024 * 1024
MAX_ZIP_RATIO = float(os.environ.get("L2W_MAX_ZIP_RATIO", "200"))
MAX_ZIP_DEPTH = 24


class ZipSafetyError(ValueError):
    """Gói ZIP vi phạm quy tắc an toàn — job phải dừng, không giải nén một phần rồi chạy tiếp."""


class InputKindError(ValueError):
    """Không xác định được loại tệp đầu vào (không phải PDF/ZIP/TeX hợp lệ)."""


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _clean_component(comp: str) -> str:
    comp = unicodedata.normalize("NFC", comp)
    comp = _FORBIDDEN_CHARS.sub("_", comp)
    comp = comp.strip().rstrip(". ")
    stem = comp.split(".")[0].upper()
    if stem in _WIN_RESERVED:
        comp = "_" + comp
    return comp[:180]


def sanitize_upload_name(name: str | None, default: str = "upload") -> str:
    """Tên tệp upload an toàn trên cả Windows và Linux: bỏ thư mục (/ và \\), ký tự cấm,
    tên dành riêng của Windows (CON, NUL…), dấu chấm/khoảng trắng cuối.  Giữ chữ tiếng Việt (NFC)."""
    raw = (name or "").replace("\\", "/").split("/")[-1]
    clean = _clean_component(raw)
    if not clean or clean in (".", ".."):
        clean = default
    return clean


def sniff_kind(data: bytes, name: str = "") -> dict[str, Any]:
    """Nhận diện loại tệp theo chữ ký nội dung; phần mở rộng chỉ để đối chiếu/cảnh báo."""
    ext = Path(name).suffix.lower().lstrip(".")
    head = data[:2048]
    kind = "unknown"
    reason = ""
    if b"%PDF-" in head[:1024]:
        kind, reason = "pdf", "chữ ký %PDF-"
    elif head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06"):
        kind, reason = "zip", "chữ ký ZIP (PK)"
    else:
        txt = None
        if head.startswith(b"\xef\xbb\xbf") or head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff"):
            txt = "bom"
        elif b"\x00" not in head:
            txt = "text"
        if txt:
            sample = data[:65536]
            try:
                if sample.startswith(b"\xff\xfe") or sample.startswith(b"\xfe\xff"):
                    s = sample.decode("utf-16", errors="strict")
                else:
                    s = sample.decode("utf-8-sig", errors="ignore")
            except UnicodeDecodeError:
                s = ""
            if re.search(r"\\[A-Za-z@]+", s):
                kind, reason = "tex", "văn bản có lệnh TeX (\\...)"
            elif s.strip():
                kind, reason = "text", "văn bản không có lệnh TeX"
    warnings: list[str] = []
    expected = {"pdf": "pdf", "zip": "zip", "tex": "tex", "ltx": "tex", "latex": "tex"}.get(ext)
    if expected and kind not in (expected, "unknown") and not (expected == "tex" and kind == "text"):
        warnings.append(f"Phần mở rộng '.{ext}' không khớp nội dung ({kind}: {reason}); xử lý theo nội dung.")
    if kind == "text" and expected == "tex":
        kind = "tex"  # .tex không có lệnh nào vẫn là nguồn TeX (vd. chỉ có chữ)
        reason = "phần mở rộng .tex + nội dung văn bản"
    return {"kind": kind, "reason": reason, "extension": ext, "warnings": warnings}


# ─────────────────────────────── ZIP ───────────────────────────────
@dataclass
class ExtractResult:
    root: Path
    files: list[str] = field(default_factory=list)          # đường dẫn tương đối (POSIX) đã giải nén
    skipped: list[dict[str, str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    total_bytes: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {"files": self.files, "skipped": self.skipped, "warnings": self.warnings,
                "total_bytes": self.total_bytes}


def _decode_zip_name(info: zipfile.ZipInfo) -> str:
    """Tên trong ZIP không bật cờ UTF-8 được Python giải mã cp437; nhiều công cụ Windows thực ra
    ghi UTF-8 → thử giải mã lại, chỉ nhận khi là UTF-8 hợp lệ."""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return unicodedata.normalize("NFC", name)


def safe_extract_zip(zip_path: Path, dest: Path, *, max_files: int = MAX_ZIP_FILES,
                     max_total: int = MAX_ZIP_TOTAL_BYTES, max_ratio: float = MAX_ZIP_RATIO) -> ExtractResult:
    """Giải nén có kiểm tra: đường dẫn tuyệt đối/ổ đĩa, `..` thoát gói, symlink/thiết bị, tệp mã hóa,
    tên trùng khi không phân biệt hoa thường, số tệp, tổng dung lượng và tỉ lệ nén (zip bomb).
    Vi phạm → `ZipSafetyError` và KHÔNG để lại tệp đã giải nén dở."""
    dest = dest.resolve()
    dest.mkdir(parents=True, exist_ok=True)
    res = ExtractResult(root=dest)
    try:
        zf = zipfile.ZipFile(zip_path, "r")
    except zipfile.BadZipFile as exc:
        raise ZipSafetyError(f"Tệp .zip không hợp lệ hoặc bị hỏng: {exc}") from exc
    with zf:
        infos = zf.infolist()
        plan: list[tuple[zipfile.ZipInfo, PurePosixPath]] = []
        seen_ci: dict[str, str] = {}
        declared_total = 0
        for info in infos:
            raw = _decode_zip_name(info)
            norm = raw.replace("\\", "/")
            if info.is_dir() or norm.endswith("/"):
                continue
            parts = [p for p in norm.split("/") if p not in ("", ".")]
            if norm.startswith("/") or re.match(r"^[A-Za-z]:", norm) or norm.startswith("//"):
                raise ZipSafetyError(f"Đường dẫn tuyệt đối trong gói: {raw!r}")
            if any(p == ".." for p in parts):
                raise ZipSafetyError(f"Đường dẫn thoát khỏi gói (..): {raw!r}")
            if len(parts) > MAX_ZIP_DEPTH:
                raise ZipSafetyError(f"Thư mục lồng quá sâu: {raw!r}")
            mode = (info.external_attr >> 16) & 0xFFFF
            if mode and (stat.S_ISLNK(mode) or stat.S_ISCHR(mode) or stat.S_ISBLK(mode)
                         or stat.S_ISFIFO(mode) or stat.S_ISSOCK(mode)):
                raise ZipSafetyError(f"Liên kết/tệp đặc biệt không được phép: {raw!r}")
            if info.flag_bits & 0x1:
                raise ZipSafetyError(f"Tệp trong gói bị mã hóa (cần mật khẩu): {raw!r}")
            if parts and (parts[0] == "__MACOSX" or parts[-1] in (".DS_Store", "Thumbs.db", "desktop.ini")):
                res.skipped.append({"path": raw, "reason": "tệp hệ thống (macOS/Windows)"})
                continue
            clean_parts = [_clean_component(p) for p in parts]
            if any(not p for p in clean_parts):
                raise ZipSafetyError(f"Tên tệp không hợp lệ: {raw!r}")
            if clean_parts != [unicodedata.normalize("NFC", p) for p in parts]:
                res.warnings.append(f"Đổi tên tệp không hợp lệ trên Windows: {raw!r} → {'/'.join(clean_parts)!r}")
            rel = PurePosixPath(*clean_parts)
            key = str(rel).casefold()
            if key in seen_ci:
                raise ZipSafetyError(f"Hai tệp trùng tên khi không phân biệt hoa/thường: {seen_ci[key]!r} và {raw!r}")
            seen_ci[key] = raw
            if info.file_size and info.compress_size and info.file_size / max(1, info.compress_size) > max_ratio \
                    and info.file_size > 1024 * 1024:
                raise ZipSafetyError(f"Tỉ lệ nén bất thường (nghi zip bomb): {raw!r}")
            declared_total += info.file_size
            plan.append((info, rel))
        if len(plan) > max_files:
            raise ZipSafetyError(f"Gói có {len(plan)} tệp, vượt giới hạn {max_files}.")
        if declared_total > max_total:
            raise ZipSafetyError(f"Tổng dung lượng giải nén {declared_total} byte vượt giới hạn {max_total}.")
        written: list[Path] = []
        try:
            for info, rel in plan:
                target = (dest / Path(*rel.parts)).resolve()
                if dest not in target.parents:
                    raise ZipSafetyError(f"Đích giải nén nằm ngoài thư mục job: {rel}")
                target.parent.mkdir(parents=True, exist_ok=True)
                remaining = max_total - res.total_bytes
                with zf.open(info, "r") as src, open(target, "wb") as out:
                    written.append(target)
                    copied = 0
                    while True:
                        chunk = src.read(1 << 16)
                        if not chunk:
                            break
                        copied += len(chunk)
                        if copied > info.file_size + 1024 or copied > remaining:
                            raise ZipSafetyError(f"Dữ liệu thực vượt kích thước khai báo: {rel}")
                        out.write(chunk)
                res.total_bytes += copied
                res.files.append(str(rel))
        except Exception:
            for p in written:
                try:
                    p.unlink()
                except OSError:
                    pass
            raise
    return res


# ─────────────────────────────── JOB ───────────────────────────────
class JobWorkspace:
    """`<jobs_root>/<job_id>/{input,work,out}` — mỗi job một thư mục, id ngẫu nhiên."""

    def __init__(self, jobs_root: Path, job_id: str | None = None) -> None:
        self.job_id = job_id or uuid.uuid4().hex[:12]
        if not JOB_ID_RE.match(self.job_id):
            raise ValueError(f"job_id không hợp lệ: {self.job_id!r}")
        self.root = (Path(jobs_root).resolve() / self.job_id)
        self.input_dir = self.root / "input"
        self.work_dir = self.root / "work"
        self.out_dir = self.root / "out"
        for d in (self.input_dir, self.work_dir, self.out_dir):
            d.mkdir(parents=True, exist_ok=True)
        self.inputs: list[dict[str, Any]] = []

    def store_input(self, name: str | None, data: bytes) -> Path:
        if len(data) > MAX_UPLOAD_BYTES:
            raise ValueError(f"Tệp tải lên {len(data)} byte vượt giới hạn {MAX_UPLOAD_BYTES} byte.")
        safe = sanitize_upload_name(name)
        path = self.input_dir / safe
        path.write_bytes(data)
        try:
            os.chmod(path, stat.S_IRUSR | stat.S_IRGRP)  # chỉ đọc: không bước nào được ghi đè bản gốc
        except OSError:
            pass
        self.inputs.append({"name": safe, "original_name": name, "path": str(path),
                            "sha256": hashlib.sha256(data).hexdigest(), "size": len(data)})
        return path

    def verify_inputs_unchanged(self) -> bool:
        for rec in self.inputs:
            p = Path(rec["path"])
            if not p.exists() or sha256_file(p) != rec["sha256"]:
                return False
        return True

    def copy_to_work(self, src: Path, name: str | None = None) -> Path:
        dst = self.work_dir / (name or src.name)
        shutil.copy2(src, dst)
        try:
            os.chmod(dst, stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP)
        except OSError:
            pass
        return dst

    @classmethod
    def resolve_existing(cls, jobs_root: Path, job_id: str) -> Path | None:
        """Đường dẫn job cho API tải xuống — chỉ nhận id đúng định dạng, không cho thoát thư mục."""
        if not JOB_ID_RE.match(job_id or ""):
            return None
        root = (Path(jobs_root).resolve() / job_id)
        return root if root.is_dir() else None
