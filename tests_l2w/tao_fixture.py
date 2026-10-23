"""Bộ tạo fixture NHỎ, TÁI LẬP ĐƯỢC cho `kiem_tra_pdf_va_file.py`.

Mọi tệp được sinh từ mã (không cần tải gì): mã TeX (chuỗi), PDF (reportlab + pikepdf + Pillow), PDF biên dịch
từ TeX (xelatex, nếu máy có), gói ZIP (zipfile). Không tệp nào được ghi ra ngoài thư mục đích truyền vào.

Quy ước: hàm `pdf_*` / `zip_*` / `tex_*` nhận thư mục đích, trả về Path tệp đã tạo (hoặc ném FixtureSkip kèm lý do
khi thiếu điều kiện như phông Unicode / TeX).
"""
from __future__ import annotations

import io
import os
import random
import re
import shutil
import stat
import subprocess
import unicodedata
import zipfile
from pathlib import Path


class FixtureSkip(RuntimeError):
    """Không tạo được fixture vì môi trường thiếu điều kiện (lý do nằm trong thông điệp)."""


# ───────────────────────────── phông chữ ─────────────────────────────
_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    "/usr/share/fonts/dejavu/DejaVuSerif.ttf",
    "/usr/share/fonts/TTF/DejaVuSerif.ttf",
    "/usr/local/share/fonts/DejaVuSerif.ttf",
    "/Library/Fonts/Times New Roman.ttf",
    "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    r"C:\Windows\Fonts\times.ttf",
    r"C:\Windows\Fonts\arial.ttf",
]
_SANS_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/TTF/DejaVuSans.ttf",
    "/Library/Fonts/Arial.ttf",
    r"C:\Windows\Fonts\arial.ttf",
]


def unicode_font(sans: bool = False) -> Path:
    env = os.environ.get("L2W_TEST_FONT")
    for p in ([env] if env else []) + (_SANS_CANDIDATES if sans else _FONT_CANDIDATES):
        if p and Path(p).is_file():
            return Path(p)
    raise FixtureSkip("không tìm thấy phông TrueType có chữ tiếng Việt (đặt L2W_TEST_FONT=<đường dẫn .ttf>)")


def _register(font_path: Path, name: str = "L2WSerif") -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    if name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(name, str(font_path)))
    return name


# ───────────────────────────── văn bản mẫu ─────────────────────────────
VN_LINES = [
    "Trường Trung học phổ thông Nguyễn Huệ",
    "Đề kiểm tra môn Toán lớp mười hai",
    "Học sinh đọc kỹ đề bài trước khi làm.",
    "Phương trình bậc hai có nghiệm thực khi biệt thức không âm.",
    "Người ra đề: thầy Lê Văn Thưởng, tổ Toán – Tin.",
    "Thời gian làm bài chín mươi phút, không kể thời gian phát đề.",
]

EN_SCAN_LINES = [
    "The quick brown fox jumps over the lazy dog.",
    "Scanned documents need optical character recognition.",
    "This page contains only a raster image of text.",
    "Numbers like 2024 and 365 should be recognised too.",
    "Every line here is printed in a large clear font.",
]


# ───────────────────────────── PDF bằng reportlab ─────────────────────────────
def _canvas(path: Path, size=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path), pagesize=size or A4)
    c.setCreator("l2w-tests")
    c.setProducer("l2w-tests (reportlab)")
    return c


def pdf_vietnamese(out: Path, *, nfd: bool = False, name: str | None = None) -> Path:
    """Một trang chữ tiếng Việt, phông TrueType nhúng tập con (subset). nfd=True: chữ ở dạng tổ hợp (NFD)."""
    font = _register(unicode_font())
    path = out / (name or ("vn_nfd.pdf" if nfd else "vn_dung_san.pdf"))
    c = _canvas(path)
    y = 780
    c.setFont(font, 13)
    for line in VN_LINES:
        c.drawString(60, y, unicodedata.normalize("NFD", line) if nfd else line)
        y -= 26
    c.showPage()
    c.save()
    return path


def _font_dicts(pdf):
    for page in pdf.pages:
        res = page.obj.get("/Resources")
        fonts = res.get("/Font") if res is not None else None
        if fonts is None:
            continue
        for k in list(fonts.keys()):
            yield fonts[k]


def pdf_broken_tounicode(out: Path, *, mode: str) -> Path:
    """Từ PDF tiếng Việt chuẩn: mode="removed" — xoá /ToUnicode (glyph không ánh xạ được);
    mode="shifted" — ánh xạ SAI có chủ đích (mỗi mã đích +1), chữ trích ra là chữ khác nhưng vẫn "hợp lệ"."""
    import pikepdf
    src = pdf_vietnamese(out, name=f"_src_{mode}.pdf")
    path = out / f"vn_tounicode_{mode}.pdf"
    with pikepdf.open(str(src)) as pdf:
        n = 0
        for f in _font_dicts(pdf):
            if "/ToUnicode" not in f:
                continue
            if mode == "removed":
                del f["/ToUnicode"]
                n += 1
            else:
                data = f["/ToUnicode"].read_bytes().decode("latin-1")

                def shift(m: re.Match) -> str:
                    v = int(m.group(2), 16) + 1
                    return f"<{m.group(1)}> <{v:0{len(m.group(2))}X}>"

                new = re.sub(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", shift, data)
                f["/ToUnicode"] = pikepdf.Stream(pdf, new.encode("latin-1"))
                n += 1
        if n == 0:
            raise FixtureSkip("PDF nguồn không có /ToUnicode để làm hỏng")
        pdf.save(str(path))
    src.unlink()
    return path


def pdf_two_columns(out: Path) -> Path:
    font = _register(unicode_font())
    path = out / "hai_cot.pdf"
    c = _canvas(path)
    c.setFont(font, 11)
    for i in range(1, 9):
        c.drawString(50, 790 - i * 22, f"Cột trái câu số {i} kết thúc.")
        c.drawString(320, 790 - i * 22, f"Cột phải dòng thứ {i} hết.")
    c.showPage()
    c.save()
    return path


def pdf_rotated_and_cropbox(out: Path) -> Path:
    """Trang 1: /Rotate 90, nội dung được xoay ngược để HIỂN THỊ thẳng. Trang 2: CropBox nhỏ hơn MediaBox, có chữ
    nằm ngoài CropBox (không hiển thị → không được đưa vào Word)."""
    import pikepdf
    from reportlab.lib.pagesizes import A4
    font = _register(unicode_font())
    tmp = out / "_rot_src.pdf"
    c = _canvas(tmp)
    W, H = A4
    c.setFont(font, 14)
    c.drawString(60, H - 80, "Trang xoay chín mươi độ vẫn đọc thẳng.")
    c.drawString(60, H - 110, "Dòng thứ hai của trang xoay.")
    c.setLineWidth(1.2)
    c.rect(60, H - 300, 160, 120)
    c.drawString(64, H - 175, "A")
    c.drawString(206, H - 175, "B")
    c.showPage()
    c.setFont(font, 14)
    c.drawString(80, 650, "Chữ trong vùng cắt hiển thị.")
    c.drawString(80, 150, "NGOAIVUNGCAT không hiển thị.")
    c.showPage()
    c.save()
    path = out / "xoay_va_cropbox.pdf"
    with pikepdf.open(str(tmp)) as pdf:
        p1 = pdf.pages[0]
        p1.contents_coalesce()
        body = p1.obj["/Contents"].read_bytes()
        pre = f"q 0 1 -1 0 {H:.4f} 0 cm\n".encode()
        p1.obj["/Contents"] = pikepdf.Stream(pdf, pre + body + b"\nQ\n")
        p1.obj["/MediaBox"] = pikepdf.Array([0, 0, H, W])
        p1.obj["/Rotate"] = 90
        p2 = pdf.pages[1]
        p2.obj["/CropBox"] = pikepdf.Array([40, 400, 555, 842])
        pdf.save(str(path))
    tmp.unlink()
    return path


def pdf_form_xobject(out: Path) -> Path:
    font = _register(unicode_font())
    path = out / "form_xobject.pdf"
    c = _canvas(path)
    c.beginForm("khoi_chu")
    c.setFont(font, 13)
    c.drawString(60, 700, "Chữ nằm trong Form XObject.")
    c.drawString(60, 676, "Dòng thứ hai trong XObject.")
    c.endForm()
    c.doForm("khoi_chu")
    c.setFont(font, 13)
    c.drawString(60, 740, "Chữ thường ngoài XObject.")
    c.showPage()
    c.save()
    return path


def _text_image(lines: list[str], size=(1654, 2339), font_px: int = 46):
    from PIL import Image, ImageDraw, ImageFont
    img = Image.new("L", size, 255)
    d = ImageDraw.Draw(img)
    try:
        f = ImageFont.truetype(str(unicode_font(sans=True)), font_px)
    except FixtureSkip:
        f = ImageFont.load_default()
    y = 160
    for ln in lines:
        d.text((120, y), ln, fill=0, font=f)
        y += int(font_px * 1.9)
    return img


def pdf_scan(out: Path, *, hidden_text: bool = False, lines: list[str] | None = None, name: str | None = None) -> Path:
    """Trang chỉ có ảnh raster của chữ (giả lập scan 200 dpi). hidden_text=True: thêm lớp chữ ẩn (Tr 3) trùng nội dung
    — như PDF đã OCR sẵn."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    lines = lines or EN_SCAN_LINES
    path = out / (name or ("scan_lop_an.pdf" if hidden_text else "scan.pdf"))
    img = _text_image(lines)
    c = _canvas(path)
    W, H = A4
    c.drawImage(ImageReader(img), 0, 0, W, H)
    if hidden_text:
        t = c.beginText()
        t.setTextRenderMode(3)
        t.setFont("Helvetica", 16)
        y = H - 160 * H / img.size[1] - 14
        for ln in lines:
            t.setTextOrigin(120 * W / img.size[0], y)
            t.textLine(ln)
            y -= 46 * 1.9 * H / img.size[1]
        c.drawText(t)
    c.showPage()
    c.save()
    return path


def pdf_mixed(out: Path) -> Path:
    """Trang hỗn hợp: chữ thật + một ảnh raster (ảnh chụp giả) chiếm ~30% trang."""
    from PIL import Image
    from reportlab.lib.utils import ImageReader
    font = _register(unicode_font())
    rnd = random.Random(7)
    img = Image.new("RGB", (600, 400))
    px = img.load()
    for x in range(600):
        for y in range(400):
            px[x, y] = ((x * 3 + rnd.randint(0, 40)) % 256, (y * 2 + rnd.randint(0, 40)) % 256, 120)
    path = out / "hon_hop.pdf"
    c = _canvas(path)
    c.setFont(font, 13)
    c.drawString(60, 790, "Trang hỗn hợp có chữ và ảnh chụp.")
    c.drawImage(ImageReader(img), 60, 420, 480, 320)
    c.drawString(60, 380, "Chú thích dưới ảnh chụp minh hoạ.")
    c.drawString(60, 350, "Đoạn văn tiếp theo sau ảnh.")
    c.showPage()
    c.save()
    return path


def pdf_long_table(out: Path, rows: int = 70) -> Path:
    """Bảng dữ liệu dài tự sang trang (dòng tiêu đề lặp lại) — bảng trải 2 trang."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    font = _register(unicode_font())
    path = out / "bang_hai_trang.pdf"
    data = [["STT", "Họ và tên", "Điểm"]] + [[str(i), f"Học sinh số {i}", f"{(i * 7) % 10}.{i % 10}"]
                                              for i in range(1, rows + 1)]
    t = Table(data, repeatRows=1, colWidths=[50, 250, 80])
    t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), font, 11), ("GRID", (0, 0), (-1, -1), 0.6, colors.black)]))
    doc = SimpleDocTemplate(str(path), pagesize=A4, creator="l2w-tests")
    doc.build([t])
    return path


def pdf_mcq(out: Path) -> Path:
    font = _register(unicode_font())
    path = out / "trac_nghiem.pdf"
    c = _canvas(path)
    c.setFont(font, 12)
    c.drawString(50, 790, "Câu 1. Giá trị của biểu thức hai cộng ba bằng")
    for x, s in zip((50, 180, 310, 440), ("A. 4.", "B. 5.", "C. 6.", "D. 7.")):
        c.drawString(x, 766, s)
    c.drawString(50, 730, "Câu 2. Thủ đô của Việt Nam là")
    c.drawString(50, 706, "A. Hà Nội.")
    c.drawString(300, 706, "B. Huế.")
    c.drawString(50, 684, "C. Đà Nẵng.")
    c.drawString(300, 684, "D. Cần Thơ.")
    c.showPage()
    c.save()
    return path


def pdf_figures(out: Path) -> Path:
    """Hình đơn giản (tam giác + nhãn đỉnh sát mép) và hình cong/tô/nét đứt (khó dựng native)."""
    font = _register(unicode_font())
    path = out / "hinh_ve.pdf"
    c = _canvas(path)
    c.setFont(font, 12)
    c.drawString(50, 800, "Hình 1. Tam giác ABC.")
    c.setLineWidth(1.2)
    p = c.beginPath()
    p.moveTo(80, 600)
    p.lineTo(260, 600)
    p.lineTo(140, 740)
    p.close()
    c.drawPath(p, stroke=1, fill=0)
    c.drawString(66, 590, "A")
    c.drawString(263, 590, "B")
    c.drawString(136, 744, "C")
    c.drawString(50, 540, "Hình 2. Đường cong tô màu và nét đứt.")
    c.setFillColorRGB(0.75, 0.85, 1.0)
    p2 = c.beginPath()
    p2.moveTo(80, 330)
    p2.curveTo(140, 520, 260, 300, 330, 480)
    p2.lineTo(330, 330)
    p2.close()
    c.drawPath(p2, stroke=1, fill=1)
    c.setDash(4, 3)
    c.line(80, 330, 330, 480)
    c.setDash()
    c.setFillColorRGB(0, 0, 0)
    c.drawString(334, 484, "M")
    c.drawString(50, 300, "Đoạn văn sau hình vẽ.")
    c.showPage()
    c.save()
    return path


def pdf_two_pages(out: Path) -> Path:
    font = _register(unicode_font())
    path = out / "hai_trang.pdf"
    c = _canvas(path)
    c.setFont(font, 13)
    for i in range(1, 7):
        c.drawString(60, 800 - 26 * i, f"Trang một dòng {i}: nội dung kiểm tra số {i}.")
    c.showPage()
    c.setFont(font, 13)
    for i in range(1, 7):
        c.drawString(60, 800 - 26 * i, f"Trang hai dòng {i}: phần cuối tài liệu {i}.")
    c.showPage()
    c.save()
    return path


# ───────────────────────────── PDF biên dịch từ TeX ─────────────────────────────
def compile_tex_pdf(tex_path: Path, out: Path, engine: str = "xelatex", timeout: int = 180) -> Path:
    exe = shutil.which(engine)
    if not exe:
        raise FixtureSkip(f"máy không có {engine}")
    build = out / f"_build_{tex_path.stem}"
    build.mkdir(parents=True, exist_ok=True)
    text = tex_path.read_text(encoding="utf-8")
    m = re.search(r"%L2W-FONT-BEGIN.*?%L2W-FONT-END[^\n]*\n", text, flags=re.S)
    if m and engine in ("xelatex", "lualatex"):
        # Dùng ĐÚNG tệp phông TrueType mà các fixture PDF khác dùng (luôn có thật trên máy) → không phụ thuộc tên
        # phông cài trong TeX (MiKTeX/TeX Live khác nhau).
        try:
            font = unicode_font()
            opts = "Path=%s/" % font.parent.as_posix()
            for cand in (font.stem + "bd", font.stem + "-Bold", font.stem + "b"):
                if (font.parent / (cand + font.suffix)).is_file():
                    opts += ",BoldFont=" + cand + font.suffix
                    break
            repl = "\\setmainfont{%s}[%s]\n" % (font.name, opts)
            text = text[:m.start()] + repl + text[m.end():]
        except FixtureSkip:
            pass
    (build / tex_path.name).write_text(text, encoding="utf-8")
    env = dict(os.environ, openout_any="p", openin_any="p")
    for _ in range(2):
        r = subprocess.run([exe, "-no-shell-escape", "-interaction=nonstopmode", "-halt-on-error", tex_path.name],
                           cwd=build, capture_output=True, timeout=timeout, env=env)
    pdf = build / (tex_path.stem + ".pdf")
    if r.returncode != 0 or not pdf.exists():
        out_txt = r.stdout.decode("utf-8", "replace")
        errs = [ln for ln in out_txt.splitlines() if ln.startswith("!")][:3]
        tail = out_txt[-400:]
        raise FixtureSkip(f"{engine} không biên dịch được {tex_path.name}: " + (" | ".join(errs) or "") + f" …{tail}")
    dst = out / pdf.name
    shutil.copy2(pdf, dst)
    shutil.rmtree(build, ignore_errors=True)
    return dst


# ───────────────────────────── nguồn TeX ─────────────────────────────
PREAMBLE = r"""\documentclass[12pt]{article}
\usepackage{amsmath,amssymb}
\usepackage{tikz}
"""

TEX_G1 = PREAMBLE + r"""\begin{document}
\section*{Đề kiểm tra nhanh}
Cho hàm số $f(x)=\dfrac{2x+1}{x-1}$ với $x\neq 1$. Tính $\lim\limits_{x\to+\infty} f(x)$.

\begin{ex}
Giá trị của $\sqrt{16}+\vec{u}\cdot\vec{v}$ khi $\vec u\perp\vec v$ là
\choice{$2$}{\True $4$}{$6$}{$8$}
\loigiai{Vì $\vec u\perp\vec v$ nên $\vec u\cdot\vec v=0$, do đó kết quả là $4$.}
\end{ex}

\begin{tabular}{|c|c|c|}
\hline
$x$ & $1$ & $2$ \\ \hline
$y$ & $3$ & $5$ \\ \hline
\end{tabular}

Kết thúc đề với dòng chữ có dấu: Trường học, người thầy, ước mơ.
\end{document}
"""

TEX_G3_MATH = PREAMBLE + r"""\begin{document}
Hệ phương trình $\begin{cases} x+y=3 \\ x-y=1 \end{cases}$ và $\heva{&x\ge 0\\&y\le 2}$, $\hoac{&x=1\\&x=2}$.

Ma trận $A=\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}$, định thức $\begin{vmatrix} a & b \\ c & d \end{vmatrix}=ad-bc$.

Phân số lồng $\dfrac{1}{1+\dfrac{1}{x}}$, căn bậc ba $\sqrt[3]{x^2+1}$, tích phân $\displaystyle\int_0^1 x^2\,dx=\frac{1}{3}$.

Tổng $\sum\limits_{k=1}^{n} \binom{n}{k} = 2^n - 1$ và $C_n^k = \dfrac{n!}{k!(n-k)!}$, giới hạn $\lim\limits_{n\to\infty}\left(1+\frac{1}{n}\right)^n = e$.

Vectơ $\vec{a}$, $\overrightarrow{AB}$, $\overleftarrow{CD}$, $\overleftrightarrow{EF}$, $\hat{x}$, $\bar{y}$, $\tilde{z}$, $\dot{u}$, $\ddot{v}$, $\widehat{ABC}$.

Mũi tên $a \rightarrow b$, $x \Rightarrow y$, $p \rightleftharpoons q$, trần $\left\lceil x \right\rceil$, sàn $\lfloor y \rfloor$, $x \equiv 1 \pmod 3$.

\begin{align}
I &= \int_0^1 x e^{x}\,\mathrm{d}x \label{eq:I}\\
  &= \Big[x e^x\Big]_0^1 - \int_0^1 e^x\,dx = 1. \nonumber
\end{align}
\end{document}
"""


def tex_mcq(kind: str) -> str:
    opts = {
        "short": ("$1$", "$2$", "\\True $3$", "$4$"),
        "medium": ("Hàm số đồng biến trên $\\mathbb{R}$", "\\True Hàm số nghịch biến trên $(0;1)$",
                   "Hàm số có hai cực trị", "Hàm số không có cực trị"),
        "long": ("Đồ thị hàm số luôn cắt trục hoành tại đúng một điểm duy nhất với mọi giá trị của tham số $m$",
                 "Đồ thị hàm số không cắt trục hoành với mọi giá trị thực của tham số $m$ đã cho trong đề bài",
                 "\\True Đồ thị hàm số cắt trục hoành tại ba điểm phân biệt khi và chỉ khi $m$ thuộc khoảng $(-2;2)$",
                 "Đồ thị hàm số tiếp xúc với trục hoành tại gốc tọa độ khi tham số $m$ bằng không"),
    }[kind]
    return PREAMBLE + "\\begin{document}\n\\begin{ex}\nChọn khẳng định đúng.\n\\choice{" + "}{".join(opts) + \
        "}\n\\loigiai{Lời giải chi tiết ở đây.}\n\\end{ex}\n\\end{document}\n"


TEX_TF = PREAMBLE + r"""\begin{document}
\begin{ex}
Xét tính đúng sai của các mệnh đề sau.
\choiceTF{\True Hàm số $y=x^3$ đồng biến trên $\mathbb{R}$}{Phương trình $x^2+1=0$ có nghiệm thực}{\True $\sqrt{4}=2$}{$\pi$ là số hữu tỉ}
\end{ex}
\end{document}
"""

TEX_BOLD_ONLY = PREAMBLE + r"""\begin{document}
Câu 1. Chọn đáp án đúng.

\begin{tabularx}{\textwidth}{XXXX}
A. $1$ & \textbf{B.} $2$ & C. $3$ & D. $4$
\end{tabularx}
\end{document}
"""

TEX_BBT = PREAMBLE + r"""\usepackage{tkz-tab}
\begin{document}
Bảng biến thiên nhiều dòng:

\begin{tikzpicture}
\tkzTabInit[lgt=1.5,espcl=2.2]{$x$ /1, $f'(x)$ /1, $f(x)$ /2.2, $g(x)$ /2}{$-\infty$, $-1$, $1$, $+\infty$}
\tkzTabLine{ , +, 0, -, 0, + , }
\tkzTabVar{-/ $-\infty$, +/ $3$, -/ $-1$, +/ $+\infty$}
\tkzTabVar{+/ $2$, -/ $0$, +/ $5$, -/ $1$}
\end{tikzpicture}

Bảng có tiệm cận hai phía:

\begin{tikzpicture}
\tkzTabInit{$x$ /1, $y'$ /1, $y$ /2}{$-\infty$, $1$, $+\infty$}
\tkzTabLine{ , -, d, -, }
\tkzTabVar{+/ $2$, -D+/ $-\infty$ / $+\infty$, -/ $2$}
\end{tikzpicture}
\end{document}
"""

TEX_MERGED = PREAMBLE + r"""\usepackage{multirow}
\begin{document}
\begin{tabular}{|c|c|c|}
\hline
\multirow{2}{*}{Nhóm} & \multicolumn{2}{c|}{Kết quả} \\ \cline{2-3}
 & Đạt & Không đạt \\ \hline
Lớp A & 12 & 3 \\ \hline
\end{tabular}
\end{document}
"""


def tex_longtable(rows: int = 40) -> str:
    body = "\n".join(f"{i} & Học sinh {i} & {(i * 7) % 10} \\\\ \\hline" for i in range(1, rows + 1))
    return PREAMBLE + "\\usepackage{longtable}\n\\begin{document}\n\\begin{longtable}{|c|l|c|}\n\\hline\n" \
        "STT & Họ tên & Điểm \\\\ \\hline\n\\endhead\n" + body + "\n\\end{longtable}\n\\end{document}\n"


TEX_FIG_SIMPLE = PREAMBLE + r"""\begin{document}
Hình đơn giản:

\begin{tikzpicture}
\draw[thick] (0,0) -- (4,0) -- (1,3) -- cycle;
\node[below left] at (0,0) {$A$};
\node[below right] at (4,0) {$B$};
\node[above] at (1,3) {$C$};
\end{tikzpicture}

Đoạn sau hình.
\end{document}
"""

TEX_FIG_HARD = PREAMBLE + r"""\begin{document}
Hình khó dựng native (cong, tô, nét đứt, 3D):

\begin{tikzpicture}
\shade[left color=blue!20,right color=red!30] (0,0) circle (1.2);
\draw[dashed,thick,smooth] plot coordinates {(-2,0) (-1,1) (0,0.3) (1,1.4) (2,0)};
\fill[opacity=0.4,green] (0,0) -- (1,2) -- (2,0) -- cycle;
\draw[->] (0,0,0) -- (0,0,2) node[left] {$z$};
\node at (2.4,1.2) {$M$};
\end{tikzpicture}

Đoạn sau hình khó.
\end{document}
"""

TEX_IMMINI = PREAMBLE + r"""\begin{document}
\immini{Cho tam giác $ABC$ vuông tại $A$ như hình bên. Tính độ dài cạnh $BC$ biết $AB=3$, $AC=4$.}{
\begin{tikzpicture}
\draw[thick] (0,0) -- (3,0) -- (0,2) -- cycle;
\node[below left] at (0,0) {$A$};
\node[below right] at (3,0) {$B$};
\node[above left] at (0,2) {$C$};
\end{tikzpicture}}

Đoạn sau bố cục immini.
\end{document}
"""

TEX_TIKZ_ERROR = PREAMBLE + r"""\begin{document}
Đoạn TRƯỚC khối hỏng phải còn nguyên.

\begin{tikzpicture}
\shade[ball color=red] (0,0) circle (1);
\draw (0,0) -- (1,1) node {$\undefinedmacroxyz$};
\end{tikzpicture}

Đoạn SAU khối hỏng phải còn nguyên.
\end{document}
"""

TEX_VN_PRECOMPOSED = PREAMBLE + "\\begin{document}\n" + "\n\n".join(VN_LINES) + "\n\\end{document}\n"


# ───────────────────────────── gói ZIP ─────────────────────────────
def _zip(path: Path, files: dict[str, bytes | str]) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
    return path


def zip_same_source(out: Path) -> Path:
    return _zip(out / "cung_nguon.zip", {"de/main.tex": TEX_G1})


def zip_nested(out: Path) -> Path:
    main = PREAMBLE + "\\begin{document}\nMở đầu tài liệu.\n\\input{chuong/c1}\n\\include{chuong/c2}\nKết thúc tài liệu.\n\\end{document}\n"
    return _zip(out / "long_nhau.zip", {
        "goc/main.tex": main,
        "goc/chuong/c1.tex": "Chương một nội dung.\n\\input{chuong/phan/p1}\n",
        "goc/chuong/phan/p1.tex": "Phần lồng sâu với công thức $a^2+b^2=c^2$.\n",
        "goc/chuong/c2.tex": "Chương hai được include.\n",
    })


def zip_multi_main(out: Path) -> Path:
    a = PREAMBLE + "\\begin{document}\nTài liệu thứ nhất.\n\\end{document}\n"
    b = PREAMBLE + "\\begin{document}\nTài liệu thứ hai.\n\\end{document}\n"
    return _zip(out / "nhieu_main.zip", {"de1.tex": a, "de2.tex": b})


def zip_missing_asset(out: Path) -> Path:
    main = PREAMBLE + "\\usepackage{graphicx}\n\\begin{document}\nTrước ảnh.\n\n\\includegraphics[width=5cm]{hinh/khong_co.png}\n\nSau ảnh.\n\\end{document}\n"
    return _zip(out / "thieu_anh.zip", {"main.tex": main})


def zip_with_image(out: Path) -> Path:
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (120, 80), (200, 30, 30)).save(buf, "PNG")
    main = PREAMBLE + "\\usepackage{graphicx}\n\\graphicspath{{hinh/}}\n\\begin{document}\nẢnh có thật:\n\n" \
        "\\includegraphics[width=4cm]{do}\n\nHết.\n\\end{document}\n"
    return _zip(out / "co_anh.zip", {"main.tex": main, "hinh/do.png": buf.getvalue()})


def zip_cycle(out: Path) -> Path:
    main = PREAMBLE + "\\begin{document}\nGốc.\n\\input{a}\n\\end{document}\n"
    return _zip(out / "vong_lap.zip", {"main.tex": main, "a.tex": "Tệp A.\n\\input{b}\n", "b.tex": "Tệp B.\n\\input{a}\n"})


def zip_bad_package(out: Path) -> Path:
    main = "\\documentclass{article}\n\\usepackage{goi_khong_ton_tai_xyz}\n\\usepackage{amsmath}\n\\begin{document}\n" \
           "Tài liệu dùng gói không tồn tại nhưng nội dung vẫn phải giữ: $x^2$.\n\\end{document}\n"
    return _zip(out / "goi_khong_co.zip", {"main.tex": main})


def zip_slip(out: Path) -> Path:
    return _zip(out / "zip_slip.zip", {"main.tex": TEX_G1, "../../thoat_ra_ngoai.txt": "x"})


def zip_absolute(out: Path) -> Path:
    return _zip(out / "zip_tuyet_doi.zip", {"main.tex": TEX_G1, "/tmp/l2w_tuyet_doi.txt": "x"})


def zip_symlink(out: Path) -> Path:
    path = out / "zip_symlink.zip"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("main.tex", TEX_G1)
        info = zipfile.ZipInfo("lien_ket")
        info.create_system = 3
        info.external_attr = (stat.S_IFLNK | 0o777) << 16
        z.writestr(info, "/etc/passwd")
    return path


def zip_bomb(out: Path, mb: int = 60) -> Path:
    path = out / "zip_ty_le_cao.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("main.tex", TEX_G1)
        z.writestr("rong.bin", b"\0" * (mb * 1024 * 1024))
    return path
