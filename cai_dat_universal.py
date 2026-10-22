"""Cài đặt / gỡ LỚP BẢO HIỂM PHỦ 100% vào nhánh "Dán mã LaTeX".

    python cai_dat_universal.py              # kiểm tra & gắn móc nối vào backend/smart_paste.py
    python cai_dat_universal.py --go         # gỡ móc nối, trả file về đúng như trước
    python cai_dat_universal.py --kiem-tra   # chỉ xem trạng thái, không sửa gì

Chỉ đụng vào **một** file duy nhất là `backend/smart_paste.py` (nhánh dán mã).
Engine gốc `backend/core_converter.py` và nhánh tải file .pdf/.tex/.zip **không bị sửa
một dòng nào**.  Script luôn tạo bản sao lưu `.bak` trước khi ghi.
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
TARGET = ROOT / "backend" / "smart_paste.py"
NEW_MODULES = ["universal_fallback.py", "tex_truth.py", "docx_media.py", "structure_registry.py"]

HELPERS = '''# ─────────── LỚP BẢO HIỂM PHỦ 100% (backend/universal_fallback.py) ───────────
# Nhập mềm: thiếu module hoặc module lỗi thì mọi thứ chạy y như trước khi có lớp này.
try:
    from . import universal_fallback as _universal
except Exception as _uf_exc:  # pragma: no cover
    _universal = None
    _log(f"không nạp được universal_fallback ({_uf_exc}) → chạy như cũ")


def _universal_protect(src: str, base_dir: Path, workdir: Path | None = None):
    """Thay khối lạ bằng chỗ đánh dấu + khởi động dựng ảnh ở luồng nền."""
    if _universal is None:
        return src, None
    try:
        return _universal.protect(src, base_dir=base_dir, workdir=workdir)
    except Exception as exc:
        _log(f"universal.protect lỗi → bỏ qua: {type(exc).__name__}: {exc}")
        return src, None


def _universal_restore(output: Path, ctx) -> dict[str, Any]:
    """Đổi chỗ đánh dấu trong .docx thành ảnh vector do TeX dựng."""
    if _universal is None or ctx is None:
        return {}
    try:
        return _universal.restore(output, ctx)
    except Exception as exc:
        _log(f"universal.restore lỗi → bỏ qua: {type(exc).__name__}: {exc}")
        return {}


def _universal_abort(ctx) -> None:
    """Thu dọn khi pipeline rơi về engine gốc."""
    try:
        if ctx is not None and getattr(ctx, "renderer", None) is not None:
            ctx.renderer.close()
    except Exception:
        pass


'''

# Mỗi móc nối = (tên, mẫu tìm, mẫu thay khi CÀI, mẫu thay khi GỠ).
# Dùng biểu thức chính quy nên vẫn khớp khi bản smart_paste.py của bạn hơi khác bản chuẩn.
# Mỗi móc nối gồm hai cặp (mẫu tìm, mẫu thay): một để CÀI, một để GỠ.
# Dùng biểu thức chính quy nên vẫn khớp khi bản smart_paste.py của bạn hơi khác bản chuẩn.
EDITS: list[dict[str, object]] = [
    {
        "ten": "móc nối 1/4 — ba hàm bọc an toàn",
        "cai": (r"(?m)^def convert_pasted_latex_unified\(",
                HELPERS.replace("\\", "\\\\") + "def convert_pasted_latex_unified("),
        "go": (re.escape(HELPERS) + r"def convert_pasted_latex_unified\(",
               "def convert_pasted_latex_unified("),
    },
    {
        "ten": "móc nối 2/4 — bảo vệ khối lạ trước khi tiền xử lý",
        "cai": (r"(?m)^(    pre = None\n)(    try:\n)(        pre = preprocess_latex\()src(, base_dir=tex\.parent)",
                "\\1    # Bước 0 — khối nào engine chưa dựng được thì giao thẳng cho TeX dựng.\n"
                '    uf_src, uf_ctx = _universal_protect(src, tex.parent, workdir=tex.parent / "_l2w_truth")\n'
                "\\2\\3uf_src\\4"),
        "go": (r"(?m)^    # Bước 0 —[^\n]*\n    uf_src, uf_ctx = _universal_protect\([^\n]*\n"
               r"(    try:\n        pre = preprocess_latex\()uf_src(, base_dir=tex\.parent)",
               "\\1src\\2"),
    },
    {
        "ten": "móc nối 3/4 — đổi chỗ đánh dấu thành ảnh vector",
        "cai": (r"(?m)^(        layout_stats = apply_layout_markers\(output, pre\)\n)"
                r"(    except Exception as exc:[^\n]*\n)",
                "\\1        universal_stats = _universal_restore(output, uf_ctx)\n"
                "\\2        _universal_abort(uf_ctx)\n"),
        "go": (r"(?m)^(        layout_stats = apply_layout_markers\(output, pre\)\n)"
               r"        universal_stats = _universal_restore\(output, uf_ctx\)\n"
               r"(    except Exception as exc:[^\n]*\n)        _universal_abort\(uf_ctx\)\n",
               "\\1\\2"),
    },
    {
        "ten": "móc nối 4/4 — ghi báo cáo độ phủ vào report.json",
        "cai": (r'("layout": layout_stats, "raw_commands_left_in_word": leftovers)\}',
                '\\1,\n'
                '                             "universal_fallback": {**(uf_ctx.summary() if uf_ctx is not None else\n'
                '                                                       {"bat": False, "ly_do": "chua nap duoc module"}),\n'
                '                                                    "ket_qua": universal_stats}}'),
        "go": (r'("layout": layout_stats, "raw_commands_left_in_word": leftovers),\n'
               r'\s*"universal_fallback": \{\*\*\(uf_ctx\.summary\(\)[^\n]*\n[^\n]*\n[^\n]*"ket_qua": universal_stats\}\}',
               "\\1}"),
    },
]


def _status(text: str) -> str:
    return "ĐÃ CÀI" if "universal_fallback" in text else "CHƯA CÀI"


def install() -> int:
    if not TARGET.exists():
        print(f"✗ Không thấy {TARGET} — chạy script này ở thư mục gốc dự án.")
        return 1
    missing = [m for m in NEW_MODULES if not (ROOT / "backend" / m).exists()]
    if missing:
        print("✗ Thiếu module mới trong backend/: " + ", ".join(missing))
        return 1
    text = TARGET.read_text(encoding="utf-8")
    if _status(text) == "ĐÃ CÀI":
        print("• Đã cài từ trước — không sửa gì thêm.")
        return 0
    for e in EDITS:
        pat, repl = e["cai"]           # type: ignore[misc]
        new_text, n = re.subn(pat, repl, text, count=1)
        if not n:
            print(f"✗ Không tìm thấy điểm neo cho {e['ten']}.")
            print("  → Bản smart_paste.py của bạn khác bản chuẩn. Hãy dùng file smart_paste.py kèm theo,")
            print("    hoặc chèn tay 4 móc nối theo hướng dẫn trong PHU_100.md.")
            return 1
        text = new_text
    shutil.copy2(TARGET, TARGET.with_suffix(".py.bak"))
    TARGET.write_text(text, encoding="utf-8")
    print(f"✓ Đã cài 4 móc nối vào {TARGET.name} (bản cũ lưu ở {TARGET.name}.bak)")
    return 0


def uninstall() -> int:
    bak = TARGET.with_suffix(".py.bak")
    if bak.exists():
        shutil.copy2(bak, TARGET)
        print(f"✓ Đã trả {TARGET.name} về bản sao lưu.")
        return 0
    text = TARGET.read_text(encoding="utf-8")
    for e in reversed(EDITS):
        pat, repl = e["go"]            # type: ignore[misc]
        text, n = re.subn(pat, repl, text, count=1)
        if not n:
            print(f"  ⚠ không gỡ được {e['ten']} — hãy xoá tay hoặc dùng bản .bak")
    TARGET.write_text(text, encoding="utf-8")
    print("✓ Đã gỡ móc nối (không có .bak nên gỡ bằng cách thay ngược).")
    return 0


def check() -> int:
    print("=" * 64)
    print(" TRẠNG THÁI LỚP BẢO HIỂM PHỦ 100% ".center(64, "="))
    print("=" * 64)
    print(f"Thư mục dự án : {ROOT}")
    for m in NEW_MODULES:
        f = ROOT / "backend" / m
        print(f"  {'✓' if f.exists() else '✗'} backend/{m}" + (f"  ({f.stat().st_size:,} byte)" if f.exists() else ""))
    if TARGET.exists():
        print(f"  {'✓' if _status(TARGET.read_text(encoding='utf-8')) == 'ĐÃ CÀI' else '✗'} "
              f"backend/smart_paste.py — {_status(TARGET.read_text(encoding='utf-8'))}")
    core = ROOT / "backend" / "core_converter.py"
    if core.exists():
        print(f"  • backend/core_converter.py — KHÔNG BỊ SỬA ({core.stat().st_size:,} byte)")
    print("-" * 64)
    sys.path.insert(0, str(ROOT))
    try:
        from backend.tex_truth import describe_toolchain
        from backend.structure_registry import registry_size
        tc = describe_toolchain()
        print("Bộ công cụ trên máy này:")
        for k in ("tex_engine", "dvisvgm", "pdftocairo", "inkscape", "soffice"):
            print(f"  {'✓' if tc.get(k) else '✗'} {k:12s} {tc.get(k) or '(chưa có)'}")
        print(f"  → định dạng ảnh sẽ dùng: {tc['best_vector_format'].upper()}"
              f"   | sẵn sàng: {'CÓ' if tc['ready'] else 'CHƯA'}")
        if not tc["tex_engine"]:
            print("  ⚠ Chưa có LaTeX. Cài MiKTeX (miktex.org) hoặc TeX Live rồi chạy lại —")
            print("    thiếu nó thì hình lạ vẫn chuyển như cũ, không giống bản gốc 100%.")
        rs = registry_size()
        print(f"Sổ đăng ký: {rs['environments']} môi trường + {rs['commands']} lệnh "
              f"= {rs['environments'] + rs['commands']} cấu trúc đã khai báo.")
    except Exception as exc:
        print(f"✗ Không nạp được module mới: {type(exc).__name__}: {exc}")
        return 1
    print("=" * 64)
    return 0


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else ""
    if arg in ("--go", "--gỡ", "--uninstall"):
        raise SystemExit(uninstall())
    if arg in ("--kiem-tra", "--check"):
        raise SystemExit(check())
    rc = install()
    print()
    check()
    raise SystemExit(rc)
