# LaTeX → Word Converter v2.0 (Unified Core Engine: Native DrawingML Shapes + Full Math Tables)

Hệ thống chuyển đổi tài liệu & đề thi Toán học từ **LaTeX (`.tex`) sang Microsoft Word (`.docx`)** với **một nhân xử lý hợp nhất duy nhất (`backend/core_converter.py`)** dùng chung cho cả 3 giao diện:
- **CLI**: `python latex2word.py input.tex -o output.docx`
- **FastAPI + Web UI**: `python run_server.py` (`http://localhost:8000`)
- **Streamlit UI**: `streamlit run app.py`

## Điểm nâng cấp cốt lõi (v2.0)

| Thành phần Toán học | Công nghệ xử lý trong Nhân Hợp Nhất | Khả năng chỉnh sửa trong Microsoft Word |
| :--- | :--- | :--- |
| **Hình học TikZ & Không gian 3D** (`\draw`, `\fill`, `\coordinate`, `circle`, `parabola bend`, `dashed`, `->`) | `backend/drawingml_engine.py` biên dịch trực tiếp sang nhóm hình vẽ gốc **WordprocessingGroup (`<wpg:wgp>`)** và **WordprocessingShape (`<wps:wsp>`)** | **Chỉnh sửa 100% từng nét vẽ**, đổi màu, kéo đỉnh, di chuyển điểm và sửa nhãn công thức `<m:oMath>` ngay trong Word |
| **Đồ thị hàm số PGFPlots & TikZ Plot** (`\begin{axis}`, `\addplot`, `\draw[domain=...] plot`) | Nội suy đường cong **Cubic Bezier (`<a:cubicBezTo>`)** trong `<a:custGeom>`, tự động cắt miền `\clip` tại tiệm cận đứng và dựng hệ trục $Ox, Oy$ vector | **Chỉnh sửa trực tiếp đường cong** bằng công cụ *Edit Points* của Word, kéo thả nhãn trục $x, y, O$ và đường tiệm cận |
| **Bảng biến thiên & Bảng xét dấu (`tkz-tab` & TikZ BBT)** | `backend/math_tables.py` chuyển `\tkzTabInit`, `\tkzTabLine`, `\tkzTabVar` và BBT vẽ bằng TikZ thành **Bảng Word gốc (`<w:tbl>`)** kèm mũi tên biến thiên DrawingML | **Bấm vào sửa trực tiếp từng con số** ($-\infty, +\infty, 0, +, -$, cực trị), đường kẻ đôi tiệm cận đứng và kéo chỉnh từng mũi tên biến thiên |
| **Bảng đáp án Trắc nghiệm 4 phương án & Đúng/Sai THPT 2025** (`\choiceA..D`, `\choice`, `\boncot`, `\haicot`, `\motcot`, `\choiceTF`) | Tự động đo độ dài công thức để chia **1×4 (4 cột)**, **2×2 (2 cột)** hoặc **4×1 (1 cột)** không viền; hỗ trợ bảng 3 cột `Mệnh đề \| Đúng \| Sai` | Giữ nguyên gạch chân/in đậm đáp án đúng, căn đều cột tự động |
| **Bảng dữ liệu, Thống kê ghép nhóm & Bố cục song song (`tabular`, `\multirow`, `\multicolumn`, `\diagbox`, `\immini`)** | Hỗ trợ gộp ô ngang (`<w:gridSpan>`), gộp ô dọc (`<w:vMerge>`), đường chéo chia ô tiêu đề (`<w:tl2br>`), và đóng khung ẩn 2 cột cho `\immini{đề bài}{hình vẽ}` | Bảng Word gốc căn giữa dọc/ngang chuẩn định dạng in ấn |

## Cấu trúc thư mục dự án

```text
latex2word_starter/
├── latex2word.py               # Giao diện dòng lệnh (CLI) gọi vào backend.core_converter
├── run_server.py               # Khởi chạy FastAPI Server + Web UI (http://localhost:8000)
├── app.py                      # Giao diện Streamlit đồng bộ cùng nhân xử lý
├── requirements.txt            # Danh sách thư viện Python
├── backend/
│   ├── __init__.py
│   ├── core_converter.py       # NHÂN XỬ LÝ HỢP NHẤT (Unified Core Pipeline)
│   ├── drawingml_engine.py     # Bộ biên dịch TikZ & PGFPlots -> Word Shapes (DrawingML <wpg:wgp>)
│   ├── math_tables.py          # Bộ dựng Bảng biến thiên, Trắc nghiệm, Bảng gộp ô & \immini
│   ├── omml_math.py            # Bộ chuyển đổi LaTeX Math -> Office Math (<m:oMath>)
│   ├── postprocessor.py        # Hậu xử lý & kiểm toán bảng OOXML
│   ├── validators.py           # Hệ thống kiểm thử tự động 5 tầng (5 Validation Layers)
│   └── server.py               # FastAPI REST API endpoints
└── frontend/
    ├── index.html              # Giao diện Web Dark/Light Glassmorphism
    ├── app.js                  # Logic kết nối API & hiển thị thống kê DrawingML
    └── styles.css              # Thiết kế giao diện
```

## Hướng dẫn cài đặt & chạy nhanh

```bash
pip install -r requirements.txt

# 1. Chạy qua CLI:
python latex2word.py example.tex -o example.docx --qa-pdf

# 2. Chạy Web Server (FastAPI + Frontend):
python run_server.py

# 3. Hoặc chạy giao diện Streamlit:
py -m streamlit run app.py
```
