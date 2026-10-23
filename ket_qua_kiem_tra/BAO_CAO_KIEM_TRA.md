# Báo cáo kiểm thử PDF / .tex / .zip → Word

- Lệnh: `python.exe kiem_tra_pdf_va_file.py --qa`
- Thời điểm: 2026-10-24 00:27:49 — 109.1 s
- Môi trường: Python 3.14.7, Windows-11-10.0.26200-SP0
- Công cụ: xelatex=có, pdflatex=có, lualatex=có, pdftocairo=có, soffice=KHÔNG, tesseract=KHÔNG, gs=KHÔNG
- tesseract: []
- Tiền xử lý TeX: engine · smart_paste nạp được: True

**Tổng:** PASS 54, SKIP 2

Ý nghĩa: PASS = mọi tiêu chí ghi trong ca đúng trên máy này với fixture này (không phải "đúng 100%"); SKIP = thiếu điều kiện, KHÔNG tính là đạt; LIMITATION = hành vi chưa đạt mức lý tưởng, ca kiểm việc báo cáo trung thực.


## G1

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G1.1 | Cùng một nguồn qua dán mã / tệp .tex / gói .zip cho cùng kết quả | **PASS** | 6.2 s | 8 tiêu chí đúng |
| G1.2 | QA hiển thị (render Word→PDF + so với PDF tham chiếu) cho .tex và .pdf | **PASS** | 21.0 s | 4 tiêu chí đúng |

## G2

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G2.1 | TeX: chữ dựng sẵn (NFC) và chữ tổ hợp (NFD) cho cùng văn bản Word | **PASS** | 0.5 s | 8 tiêu chí đúng |
| G2.2 | PDF phông TrueType nhúng tập con (subset), chữ tiếng Việt dựng sẵn | **PASS** | 0.3 s | 7 tiêu chí đúng |
| G2.3 | PDF chữ tổ hợp (NFD: chữ cơ sở + dấu rời) → chữ dựng sẵn trong Word | **PASS** | 0.3 s | 8 tiêu chí đúng |
| G2.4 | PDF thiếu ToUnicode (glyph không ánh xạ) → giữ ảnh glyph + cảnh báo, không bịa chữ | **PASS** | 0.9 s | 4 tiêu chí đúng |
| G2.5 | PDF ToUnicode SAI có chủ đích (dịch mã) — có phát hiện được không | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G2.6 | Tên tệp tiếng Việt / dấu cách / tên cấm của Windows / đường dẫn lạ | **PASS** | 1.0 s | 4 tiêu chí đúng |

## G3

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G3.1 | TeX: toán 2D phức tạp → OMML native đúng cấu trúc | **PASS** | 0.4 s | 17 tiêu chí đúng |
| G3.2 | PDF biên dịch từ TeX (đề có công thức) → công thức OMML hoặc ảnh vùng, không mất nội dung | **SKIP** | 5.6 s | xelatex không biên dịch được de_thi_mau.tex: …; this (fontspec)                may be but usually is not a fontspec bug. Either (fontspec)                there is a typo in the font name/file, the font is (fontspec)   |

## G4

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G4.1 | Trắc nghiệm phương án ngắn → 1×4 (4 cột) | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G4.2 | Trắc nghiệm phương án vừa → 2×2 | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G4.3 | Trắc nghiệm phương án dài → 4×1 | **PASS** | 0.4 s | 2 tiêu chí đúng |
| G4.4 | Đáp án đánh dấu TƯỜNG MINH (\True) → answer_key=explicit; Đúng/Sai (\choiceTF) | **PASS** | 0.7 s | 4 tiêu chí đúng |
| G4.5 | Chỉ in đậm nhãn (không có dấu tường minh) → answer_key=unknown, không suy đáp án | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G4.6 | Chế độ học sinh: không lộ đáp án (gạch chân, cột Đ/S, lời giải) | **PASS** | 0.8 s | 4 tiêu chí đúng |
| G4.7 | PDF trắc nghiệm (A./B./C./D. cùng dòng và 2×2) → bảng phương án, answer_key=unknown | **PASS** | 0.3 s | 6 tiêu chí đúng |

## G5

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G5.1 | Bảng biến thiên nhiều dòng (2 dòng biến thiên) → bảng Word native | **PASS** | 0.4 s | 4 tiêu chí đúng |
| G5.2 | Tiệm cận đứng hai phía (-D+, d) → một ô có giá trị hai bên vạch kép | **PASS** | 0.4 s | 3 tiêu chí đúng |
| G5.3 | Bảng gộp ô (\multirow + \multicolumn) giữ đủ ô | **PASS** | 0.3 s | 10 tiêu chí đúng |
| G5.4 | longtable dài → lặp dòng tiêu đề khi sang trang (tblHeader) | **PASS** | 0.5 s | 3 tiêu chí đúng |
| G5.5 | PDF bảng trải nhiều trang → bảng native từng trang, không mất hàng | **PASS** | 1.2 s | 3 tiêu chí đúng |

## G6

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G6.1 | Hình TikZ đơn giản → DrawingML (đã kiểm chứng) hoặc ảnh vector; không mất hình | **PASS** | 4.6 s | 4 tiêu chí đúng |
| G6.2 | Hình cong/tô/nét đứt/3D → giữ ảnh vùng (snapshot TeX), không dựng DrawingML sai | **PASS** | 2.6 s | 3 tiêu chí đúng |
| G6.3 | \immini (chữ + hình bên cạnh) → bảng bố cục; SVG+PNG tính MỘT ảnh logic | **PASS** | 2.7 s | 4 tiêu chí đúng |
| G6.4 | PDF: hình đơn giản + hình cong/tô; nhãn đỉnh sát mép nằm trong vùng cắt | **PASS** | 1.5 s | 5 tiêu chí đúng |

## G7

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G7.1 | PDF 2 cột (đường cơ sở thẳng hàng) → đọc hết cột trái rồi cột phải | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G7.2 | Trang xoay (/Rotate 90) + trang có CropBox nhỏ | **PASS** | 0.9 s | 7 tiêu chí đúng |
| G7.3 | Chữ nằm trong Form XObject | **PASS** | 0.3 s | 2 tiêu chí đúng |
| G7.4 | Trang scan, OCR mặc định (cần gói 'vie') → giữ ảnh trang + cảnh báo, không bịa chữ | **PASS** | 0.4 s | 4 tiêu chí đúng |
| G7.5 | Trang scan tiếng Anh với ocr_lang=eng → chữ OCR (review) hoặc ảnh trang | **SKIP** | 0.0 s | máy không có tesseract |
| G7.6 | Trang hỗn hợp (chữ + ảnh chụp) | **PASS** | 0.9 s | 5 tiêu chí đúng |
| G7.7 | Scan có sẵn lớp chữ OCR ẩn → không nhân đôi chữ, báo lớp OCR | **PASS** | 0.5 s | 3 tiêu chí đúng |

## G8

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G8.1 | ZIP: \input/\include lồng nhiều cấp, main trong thư mục con | **PASS** | 0.3 s | 7 tiêu chí đúng |
| G8.2 | ZIP: main_tex SAI → lỗi 409 kèm danh sách ứng viên (không tự chọn) | **PASS** | 0.0 s | 3 tiêu chí đúng |
| G8.3 | ZIP: nhiều tệp chính → 409 main_tex_ambiguous; chọn rõ thì đúng tệp | **PASS** | 0.4 s | 2 tiêu chí đúng |
| G8.4 | ZIP: ảnh thiếu → thông báo hiện trong Word, giữ chữ trước/sau, không 'passed' | **PASS** | 0.3 s | 5 tiêu chí đúng |
| G8.5 | ZIP: ảnh có thật qua \graphicspath → ảnh trong Word | **PASS** | 0.4 s | 2 tiêu chí đúng |
| G8.6 | ZIP: vòng lặp \input → phát hiện, không treo, nội dung không lặp vô hạn | **PASS** | 0.3 s | 3 tiêu chí đúng |
| G8.7 | ZIP: gói LaTeX không tồn tại → vẫn chuyển nội dung (không cần biên dịch) | **PASS** | 0.4 s | 3 tiêu chí đúng |
| G8.8 | ZIP độc hại: thoát thư mục (../), đường dẫn tuyệt đối, symlink, tỉ lệ nén cao → 400 | **PASS** | 0.3 s | 5 tiêu chí đúng |
| G8.9 | Nhận diện theo NỘI DUNG: tệp TeX đặt đuôi .zip; ZIP nhiều PDF | **PASS** | 0.7 s | 3 tiêu chí đúng |

## G9

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G9.1 | Không có TeX (PATH rỗng): hình cần chụp → không mất, báo rõ, không 'passed' | **PASS** | 4.0 s | 5 tiêu chí đúng |
| G9.2 | Không có LibreOffice mà bật QA → nhóm E NOT_MEASURED (không tính là đạt) | **PASS** | 15.9 s | 3 tiêu chí đúng |
| G9.3 | Không có tesseract + trang scan → ảnh trang, báo ocr_unavailable | **PASS** | 4.2 s | 3 tiêu chí đúng |
| G9.4 | Khối TikZ cố ý hỏng → khối báo lỗi, chữ trước/sau còn nguyên, không 'passed' | **PASS** | 3.4 s | 5 tiêu chí đúng |

## G10

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G10.1 | 100 công thức nguồn, Word chỉ có 1 → L2 FAIL | **PASS** | 0.1 s | 2 tiêu chí đúng |
| G10.2 | PDF 2 trang, Word chỉ có trang 1 → L1 không PASS | **PASS** | 0.1 s | 1 tiêu chí đúng |
| G10.3 | Nguồn có hình, Word không có → L3 FAIL | **PASS** | 0.1 s | 1 tiêu chí đúng |
| G10.4 | Bảng nguồn có dữ liệu, Word có bảng rỗng → L4 không PASS | **PASS** | 0.1 s | 1 tiêu chí đúng |
| G10.5 | Chỉ có ảnh Word, không có nguồn đối chiếu → NOT_MEASURED, similarity=None, không 'passed' | **PASS** | 0.0 s | 5 tiêu chí đúng |

## G12

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G12.1 | Hai job song song, CÙNG tên tệp, khác nội dung → không lẫn kết quả | **PASS** | 0.4 s | 8 tiêu chí đúng |
| G12.2 | Hai PDF khác nhau cùng tên de.pdf → mỗi job đúng nội dung của mình | **PASS** | 0.6 s | 3 tiêu chí đúng |

## G11

| Ca | Nội dung | Kết quả | Thời gian | Ghi chú |
|---|---|---|---|---|
| G11.1 | Mọi .docx tạo ra trong lượt chạy: quan hệ/media/ID hợp lệ | **PASS** | 4.1 s | 2 tiêu chí đúng |
| G11.2 | Mở + lưu lại bằng python-docx: giữ công thức, bảng, ảnh, nhóm hình | **PASS** | 2.4 s | 1 tiêu chí đúng |
| G11.3 | Ca mẫu có sẵn trong repo vẫn chuyển được (không lỗi nội bộ, gói hợp lệ) | **PASS** | 13.6 s | 10 tiêu chí đúng |

Chi tiết từng tiêu chí: `ket_qua_kiem_tra.json`.
