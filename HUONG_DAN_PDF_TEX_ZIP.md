# Hướng dẫn: chuyển PDF / .tex / .zip → Word (pipeline chung)

Bản nâng cấp này đưa **mọi đường vào** (dán mã, tải `.tex`, tải `.zip`, tải `.pdf`) qua **một pipeline chung**
(`backend/pipeline.py`) dùng cho cả CLI (`latex2word.py`), Streamlit (`app.py`) và REST API (`backend/server.py` +
`frontend/`). Kết quả luôn kèm báo cáo `*.report.json` ghi rõ khối nào dựng **native** (chữ, OMML, bảng Word,
DrawingML), khối nào **giữ ảnh vùng gốc**, khối nào **chưa xử lý được** — và vì sao.

> "passed" chỉ có nghĩa là tệp đạt các tiêu chí tự động đã nêu trong báo cáo; không phải "đúng 100%". SVG/PNG
> không phải OMML hay hình Word sửa được — báo cáo tách riêng hai loại.

## 1. Cài đặt

```bash
python -m pip install -r requirements.txt
```

Công cụ hệ thống (pipeline tự dò, **không tự tải**; thiếu thì báo năng lực và chuyển sang cách dự phòng):

| Công cụ | Dùng để | Thiếu thì |
|---|---|---|
| TeX Live / MiKTeX (`xelatex`, `pdflatex`; gói `vntex` nếu nguồn dùng `\usepackage[utf8]{vietnam}`) | chụp ảnh hình TikZ khó; ảnh đối chiếu để kiểm chứng DrawingML; PDF tham chiếu nhóm E | hình giữ DrawingML "chưa kiểm chứng" (status `review`) hoặc `unresolved` + thông báo hiện trong Word |
| LibreOffice **kèm module Math** (`libreoffice-math` trên Linux) hoặc Microsoft Word | render Word → PDF cho nhóm E; thử DrawingML | nhóm E = `NOT_MEASURED`; hình dùng ảnh TeX |
| Poppler `pdftocairo` | bản SVG của vùng cắt | chỉ có PNG |
| `tesseract` + `vie.traineddata` | OCR trang scan tiếng Việt | trang scan giữ **ảnh trang** + cảnh báo |
| Ghostscript `gs` | ảnh `.eps` trong `\includegraphics` | ảnh đó `unresolved` |

Kiểm tra năng lực máy: `GET /api/status` hoặc khung "Năng lực máy" trong Streamlit.

## 2. Chạy

```bash
# CLI
python latex2word.py de_thi.tex -o de_thi.docx --qa-pdf
python latex2word.py du_an.zip --main chuong/main.tex -o ket_qua.docx
python latex2word.py tai_lieu.pdf -o tai_lieu.docx --ocr-lang vie+eng
python latex2word.py de_thi.tex -o de_hs.docx --answer-mode student --no-solutions
#   mã thoát: 0 = có .docx; 2 = --strict mà chưa 'passed'; 1 = lỗi đầu vào/chuyển đổi

# Streamlit
streamlit run app.py

# REST API + giao diện web (http://localhost:8000)
python run_server.py
```

Tham số mới (CLI / form API / Streamlit):

| Tham số | Giá trị | Ý nghĩa |
|---|---|---|
| `main_tex` / `--main` | đường dẫn **tương đối** trong gói | sai → lỗi 409 kèm danh sách ứng viên (không tự chọn tệp đầu tiên) |
| `answer_mode` | `source` \| `student` | `student`: không gạch chân/đánh dấu đáp án, để trống cột Đ/S |
| `show_solutions` | true/false | ẩn `\loigiai`, môi trường lời giải |
| `figure_policy` | `auto` \| `native` \| `snapshot` | `auto`: DrawingML chỉ khi khớp ảnh TeX (F1 hình học ≥ 0.75 + đủ nhãn), ngược lại ảnh TeX |
| `ocr`, `ocr_lang` | `auto`/`off`, vd. `vie+eng` | không có gói `vie` thì mặc định KHÔNG OCR (giữ ảnh trang) |
| `strict` | true/false | API trả 422, CLI mã thoát 2 khi chưa `passed` (tệp vẫn tải được) |

## 3. Trạng thái trong báo cáo

* `conversion_completed`: có tạo được gói `.docx` hợp lệ (kiểm content types, quan hệ, `r:embed`, `docPr id`
  duy nhất, mở lại được bằng python-docx).
* `quality_status`: `passed` | `review_required` | `failed`. `SKIP`/`NOT_MEASURED` **không** được tính là đạt.
* Nhóm kiểm định: A nội dung chữ, B toán, C hình, D bảng, E hiển thị (render Word ↔ PDF nguồn). Với PDF, chữ tham
  chiếu lấy **độc lập** từ ký tự hiển thị của chính PDF (trừ vùng đã giữ ảnh); với TeX, từ nguồn sau các biến đổi
  ngữ nghĩa (chế độ đáp án, `\newif`, macro người dùng) — **trước** bước chuẩn hoá cấu trúc của pipeline.
* `blocks.items[]` (PDF): `block_id`, `type`, `page`, `bbox` (pt, gốc trên-trái trang **hiển thị**, y hướng xuống —
  đã tính CropBox và /Rotate), `method` ∈ {text_native, math_omml, table_native, drawingml, vector_image,
  raster_image, page_image, unresolved}, `status`, `warnings`, `fallback_reason`, `assets`, `confidence`
  (luôn ghi rõ là heuristic).

## 4. Kiểm thử

```bash
python kiem_tra_pdf_va_file.py                 # 56 ca, 12 nhóm; fixture sinh lại từ mã
python kiem_tra_pdf_va_file.py --qa            # thêm render Word→PDF (cần LibreOffice/Word)
python kiem_tra_pdf_va_file.py --nhom G7,G8    # một số nhóm
python kiem_tra_pdf_va_file.py --chay-test-cu  # chạy kèm kiem_tra_smart_paste.py, kiem_tra_universal.py
```

Kết quả: `ket_qua_kiem_tra/BAO_CAO_KIEM_TRA.md` + `ket_qua_kiem_tra.json`. Trạng thái: PASS / FAIL / SKIP (thiếu điều
kiện, không tính là đạt) / LIMITATION (hành vi đã biết chưa lý tưởng; ca kiểm việc báo cáo trung thực).

## 5. Giới hạn còn lại (đã biết)

* PDF: công thức 2D chỉ nhận dạng được tập con (phân số, căn, chỉ số, dấu mũ/vectơ, tích phân/tổng có cận) từ
  **lớp chữ** của PDF; cấu trúc khác giữ **ảnh vùng**. Không có OCR toán.
* PDF: hình vẽ luôn giữ ảnh vùng (SVG + PNG), chưa dựng lại thành DrawingML; bảng biến thiên trong PDF giữ ảnh.
* PDF: bảng trải nhiều trang thành một bảng Word cho mỗi trang (chưa nối).
* PDF: ánh xạ ToUnicode SAI nhưng "trông hợp lý" không phát hiện được bằng chữ; heuristic chỉ bắt các trường hợp
  chữ trích ra không giống văn bản tự nhiên. Phông bảng mã cũ (VNI/TCVN3) chưa giải mã → giữ ảnh dòng.
* PDF scan: chất lượng OCR phụ thuộc tesseract; khối OCR luôn ở trạng thái `review`.
* TeX: macro người dùng nở được: `\newcommand`/`\renewcommand`/`\providecommand`/`\DeclareRobustCommand`,
  `\def` (kể cả tham số phân cách đơn giản), `\let`, `\DeclareMathOperator`, `\DeclarePairedDelimiter`, xparse
  `s m o O{}`; thân có điều kiện TeX/`\csname` không nở (giữ nguyên + báo cáo).
* Nhóm E so chữ hiển thị và bố cục (IoU mực là heuristic). Nếu LibreOffice thiếu module Math, PDF render không có
  công thức → báo cáo ghi `qa_renderer_drops_omml` và loại chữ trong OMML khỏi phép so.
* Kiểm chứng DrawingML bằng LibreOffice: trình này không vẽ hộp chữ trong nhóm hình → nhãn chỉ kiểm bằng **số lượng**
  trong XML (vị trí nhãn chưa kiểm bằng hình).
