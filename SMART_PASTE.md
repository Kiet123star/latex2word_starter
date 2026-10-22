# Lớp tiền xử lý thông minh cho nhánh "Dán mã LaTeX"

Chỉ áp dụng cho **dán mã** (`/api/convert/paste` và tab "Cách 1" trong `app.py`).
Nhánh **tải file** (.pdf / .tex / .zip, `decompile_pdf_to_latex`) vẫn gọi đúng engine gốc, không đổi gì.

```
mã dán ──► backend/latex_preprocessor.py   (mở rộng macro + dịch cấu trúc lạ về lệnh engine đã hỗ trợ)
       ──► core_converter.convert_latex_file_unified   (engine gốc, KHÔNG sửa)
       ──► backend/smart_paste.py          (áp căn lề, màu, viền bảng chuẩn LaTeX, footnote, cột, header/footer…)
```

- Lỗi ở bất kỳ bước tiền/hậu xử lý nào → **tự động quay về engine gốc** (kết quả không bao giờ tệ hơn trước).
- Tắt tạm thời: đặt biến môi trường `L2W_SMART_PASTE=0`.
- Báo cáo nằm trong `report.json` → khoá `smart_paste` (số macro đã mở rộng, lệnh/môi trường lạ đã xử lý, cảnh báo).
- Giữ lại file trung gian để gỡ lỗi: `L2W_KEEP_SMART_TEX=1` (tạo `document.l2w_smart.tex`).

## File
| File | Vai trò |
|---|---|
| `backend/latex_preprocessor.py` | Phase A: macro người dùng (`\newcommand`, `\def` có tham số phân cách, `\let`, xparse, `\newenvironment`, `\newif`, `\ifthenelse`, `\DeclareMathOperator`, `\DeclarePairedDelimiter`, `\newcolumntype`…); khung dựng khối, công thức đánh số |
| `backend/latex_preprocessor_parts.py` | Phase B: đề thi (ex_test, exam), `\cau`, `\bonpa`, `\choice*`, `\choiceTF` 4/8 tham số, `\loigiai`, `\shortans`, `\immini`, danh sách, bảng, định lý, chuẩn hoá toán, `\ce`, siunitx, fallback lệnh lạ |
| `backend/smart_paste.py` | Hậu xử lý .docx + hàm `convert_pasted_latex_unified()` (cùng chữ ký với `convert_latex_file_unified`) |

## Nguyên tắc fallback
Lệnh không ai định nghĩa (`\abc{...}`) → bỏ tên lệnh, giữ nội dung đối số (bỏ đối số kiểu kích thước/màu);
lệnh có tên giống trắc nghiệm + đúng 4 đối số → bảng A/B/C/D; môi trường lạ → giữ nội dung (đối số chữ thành dòng tiêu đề).
Không bao giờ in nguyên văn `\tenlenh` ra Word.

## Kiểm thử
`examples/mau_kiem_thu_de_thi.tex` và `examples/mau_kiem_thu_cau_truc_la.tex` — dán vào tab "Cách 1" để xem kết quả.

## Tự kiểm tra
Chạy `python kiem_tra_smart_paste.py` ở thư mục gốc dự án: in ra file backend đang nạp, hook đã gắn chưa,
và chuyển đổi thử đoạn `\cau` / `\bonpa` bằng engine thật → tạo `kiem_tra_ket_qua.docx`, kết luận ĐẠT / CHƯA ĐẠT.
Trên giao diện web, thẻ "Tiền Xử Lý Thông Minh (dán mã)" hiện `OK` hoặc lý do lỗi; terminal chạy server in dòng `[smart_paste] ...`.
