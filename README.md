# Slide & PDF Smart Compressor

Ứng dụng Web thông minh giúp nén các file thuyết trình (**PDF**, **PowerPoint `.pptx` / `.ppt`**) có nhiều hình ảnh dung lượng lớn (Canva, Slide thiết kế, báo cáo học thuật) về dưới ngưỡng **Dung lượng mục tiêu (Target MB - VD: < 5 MB, < 10 MB)** mà vẫn giữ độ sắc nét tối đa.

## Tính năng nổi bật

- **Tự động căn chỉnh theo Dung lượng mục tiêu (Auto-Target MB)**: Nhập giới hạn dung lượng mong muốn (ví dụ `10 MB`), thuật toán tự động lấy mẫu trang và tìm kiếm thông số `(scale, jpeg_quality)` tối ưu nhất.
- **2 Chế độ nén PDF chuyên biệt**:
  1. **Gộp phẳng trang (`flatten`)**: Giải quyết triệt để các file PDF xuất từ Canva có hàng trăm lớp ảnh PNG trong suốt chồng lớp (giảm tới 90–96% dung lượng).
  2. **Giữ nguyên chữ Vector (`vector`)**: Nén các ảnh nhúng (`XObject`) và loại bỏ tài nguyên trùng lặp nhưng giữ nguyên 100% chữ vector tiếng Việt có thể bôi đen/tìm kiếm.
- **2 Chế độ xử lý PowerPoint (`.pptx` / `.ppt`)**:
  1. **Nén trực tiếp `.pptx` → `.pptx` nhẹ**: Tối ưu hóa toàn bộ ảnh trong `ppt/media/*` (JPEG & PNG 256-color Alpha quantization) giữ nguyên cấu trúc slide.
  2. **Chuyển đổi `.pptx/.ppt` → `.pdf` nhẹ**: Xuất PDF chuẩn qua Microsoft PowerPoint COM (kèm fallback thuần Python) rồi nén theo `Target MB`.
- **Giao diện Web trực quan**: Hỗ trợ kéo thả, theo dõi tiến độ từng trang và khung **So sánh chất lượng Trước / Sau** trực tiếp trên trình duyệt.

## Cài đặt & Khởi chạy

```bash
pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8080
```

Hoặc trên Windows, nhấp đúp vào file `run.bat` để tự động khởi chạy và mở trình duyệt tại `http://127.0.0.1:8080`.

## Kiểm thử (Tests)

```bash
python -m pytest tests/test_compressors.py -v
```
