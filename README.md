# Slide & PDF Smart Compressor

Ứng dụng Web thông minh giúp **nén dung lượng** và **giải nén / trích xuất tài nguyên (`.ZIP`)** cho các tệp thuyết trình (**PDF**, **PowerPoint `.pptx` / `.ppt`**) có nhiều hình ảnh lớn (Canva, Slide thiết kế, báo cáo học thuật) về dưới ngưỡng **Dung lượng mục tiêu (Target MB - VD: < 5 MB, < 10 MB)** mà vẫn bảo toàn độ sắc nét tối đa.

## Tính năng nổi bật

### 1. Nén Thông Minh theo Dung lượng mục tiêu (`Auto-Target MB`)
- **Tự động cân chỉnh thông số tối ưu**: Nhập giới hạn dung lượng mong muốn (ví dụ `10 MB`), thuật toán tự động lấy mẫu và tìm kiếm cặp `(scale, jpeg_quality)` tốt nhất kèm **chốt chặn chống phình dung lượng (non-inflation guard)** và cảnh báo rõ ràng nếu tệp đã chạm ngưỡng nén an toàn (`target_met`).
- **2 Chế độ nén PDF chuyên biệt**:
  1. **Gộp phẳng trang (`flatten`)**: Giải quyết triệt để các tệp PDF xuất từ Canva có hàng trăm lớp ảnh PNG trong suốt chồng lớp (giảm tới 90–96% dung lượng), giữ nguyên chính xác kích thước vật lý `(width_pt, height_pt)` từng trang với cơ chế đệm byte tiết kiệm RAM.
  2. **Giữ nguyên chữ Vector (`vector`)**: Chỉ tối ưu các ảnh nhúng (`XObject`) khi thực sự giảm được dung lượng, giữ nguyên 100% lớp chữ vector tiếng Việt để bôi đen, sao chép và tìm kiếm.
- **Xử lý thông minh PowerPoint (`.pptx` vs `.ppt` cũ)**:
  1. **Nén trực tiếp `.pptx` → `.pptx` nhẹ**: Tối ưu hóa đa định dạng ảnh trong `ppt/media/*` (`.jpg`, `.png`, `.webp`, `.bmp`, `.tiff`, `.gif`), tự động chuyển ảnh PNG/BMP/TIFF không trong suốt kích thước lớn sang JPEG kèm cập nhật liên kết XML `.rels`, truyền dữ liệu ZIP theo luồng (streaming) tiết kiệm RAM.
  2. **Chuyển đổi `.pptx` / `.ppt` → `.pdf` nhẹ**: Xuất PDF chuẩn qua tiến trình Microsoft PowerPoint COM cô lập có timeout an toàn (tự động dọn tiến trình ẩn, không rò rỉ `POWERPNT.EXE`, kèm cơ chế dự phòng dựng slide thuần Python hỗ trợ font tiếng Việt cho `.pptx`).
  3. **Nhận diện riêng định dạng `.ppt` nhị phân cũ**: Tự động khóa tùy chọn nén trực tiếp sang `.pptx` hoặc trích xuất ảnh nhúng trực tiếp, hướng dẫn chuyển sang xuất PDF nhẹ hoặc tách bộ ảnh slide.

### 2. Giải Nén / Trích Xuất Tài Nguyên (`.ZIP`)
- **Trích xuất toàn bộ ảnh gốc nhúng trong tệp (`images`)**: Tách nguyên bản toàn bộ ảnh độ phân giải gốc bên trong `.pptx` (`ppt/media/*`) hoặc `.pdf` (`XObject images`), tự động khử trùng lặp bằng mã băm SHA-256 và đóng gói thành tệp `.zip`.
- **Tách từng trang / slide thành bộ ảnh HD (`slides`)**: Kết xuất từng trang PDF hoặc từng slide thuyết trình (`.pptx` / `.ppt`) thành bộ ảnh `.jpg` chuẩn Full HD (~1920px) đóng gói gọn trong `.zip`.

### 3. Kiến trúc Backend & Giao diện Web hiện đại
- **Streaming Upload & Dọn dẹp tự động (TTL Cleanup)**: Tải tệp theo từng khối `1 MB` (hỗ trợ tối đa `500 MB`, trả mã HTTP `413` nếu vượt ngưỡng) và tự động dọn các tệp tạm quá hạn 2 giờ trên máy chủ.
- **Giao diện 2 Tab trực quan**: Chuyển đổi mượt mà giữa **🗜️ Nén Thông Minh** và **📦 Giải Nén / Trích Xuất (.ZIP)**, theo dõi tiến độ thời gian thực và so sánh chất lượng Trước / Sau ngay trên trình duyệt.

## Cấu trúc dự án

```text
slide-pdf-compressor/
├── app.py                    # FastAPI server (/api/analyze, /api/compress, /api/extract, /api/job, /api/download)
├── core/
│   ├── pdf_compressor.py     # Engine phân tích & nén PDF (flatten / vector)
│   ├── pptx_compressor.py    # Engine phân tích, nén trực tiếp PPTX & chuyển đổi PPTX/PPT -> PDF
│   └── extractor.py          # Engine giải nén ảnh gốc & tách trang/slide thành gói .ZIP
├── static/
│   ├── index.html            # Giao diện người dùng 2 chế độ (Nén & Giải nén .ZIP)
│   ├── app.js                # Logic tương tác API, cập nhật tiến độ & bảng kết quả
│   └── style.css             # Giao diện Glassmorphism hỗ trợ chế độ Sáng / Tối
└── tests/
    └── test_compressors.py   # Bộ kiểm thử tự động toàn diện (pytest)
```

## Cài đặt & Khởi chạy

```bash
pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8080
```

Hoặc trên Windows, nhấp đúp vào tệp `run.bat` để tự động khởi chạy và mở trình duyệt tại `http://127.0.0.1:8080`.

## Kiểm thử (Tests)

```bash
python -m pytest tests/test_compressors.py -v
```
