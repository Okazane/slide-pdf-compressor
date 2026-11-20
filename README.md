# Slide & PDF Smart Compressor (100% Offline Static PWA + Optional Hardened Python Engine)

Ứng dụng Web PWA thông minh giúp **nén dung lượng theo mức mục tiêu (Target MB)** và **giải nén / trích xuất tài nguyên (`.ZIP`)** cho các tệp thuyết trình (**PDF**, **PowerPoint `.pptx`**) **100% cục bộ ngay trên trình duyệt Offline** — không cần máy chủ backend, không tải tệp nhạy cảm lên đám mây, triển khai trực tiếp lên **Vercel Static Hosting** (`outputDirectory: "static"`).

Dự án đồng thời đi kèm máy chủ **FastAPI Python cục bộ tùy chọn (`app.py`)** đã được gia cố bảo mật toàn diện (chống Path/Glob Injection, ZIP/XML/Image Bomb, giới hạn hàng đợi Worker Pool và làm sạch thông báo lỗi), hỗ trợ thêm định dạng `.ppt` nhị phân cũ (PowerPoint 97–2003) qua COM cô lập trên Windows.

---

## Điểm nổi bật của Phiên bản 2.0 (Offline-First Static PWA)

### 1. Xử lý 100% Cục bộ trên Trình duyệt Offline (`static/offline-engine.js`)
- **Không phụ thuộc CDN hay Internet**: Toàn bộ thư viện lõi (`pdf.js`, `pdf-lib`, `JSZip`, `Dexie.js`), bộ font tiếng Việt `Be Vietnam Pro` (`.woff2` đầy đủ dải Unicode tiếng Việt, trọng lượng 400–800) và bộ icon PWA (`192x192`, `512x512`, `favicon.svg`) được đóng gói hoàn toàn cục bộ trong `static/`.
- **Service Worker thông minh (`static/sw.js`)**: Precache toàn bộ giao diện, font chữ và thư viện xử lý; áp dụng chiến lược **Network-First** (sao lưu đồng bộ `res.clone()`) cho HTML và **Cache-First / Stale-While-Revalidate** kèm `Response.error()` an toàn cho tài nguyên tĩnh.
- **Nén PDF theo Dung lượng mục tiêu (`flatten` & `vector`)**:
  - `flatten`: Lấy mẫu 4 trang tìm cặp `(scale, quality)` tối ưu `<= targetMb * 0.90`, dựng từng trang một qua Canvas tiết kiệm RAM, giữ nguyên kích thước vật lý `(w_pt, h_pt)` từng trang bằng `pdf-lib`, tự động hạ bậc lặp nếu vượt ngưỡng và có **chốt chặn chống phình dung lượng (non-inflation guard)**.
  - `vector`: Tái đóng gói luồng đối tượng PDF và nén chọn lọc ảnh nhúng khi giảm được dung lượng mà vẫn giữ nguyên 100% lớp chữ Vector tiếng Việt để bôi đen, sao chép và tìm kiếm.
- **Nén trực tiếp `.PPTX` & Chuyển đổi `.PPTX → .PDF` Offline**:
  - Tối ưu ảnh trong `ppt/media/*` theo cấu hình thích ứng, tự động chuyển ảnh PNG/BMP không trong suốt kích thước lớn sang JPEG kèm cập nhật an toàn `*.rels` & `[Content_Types].xml` qua `DOMParser` (kiểm tra `parsererror` và chặn `<!DOCTYPE` / `<!ENTITY`).
  - Dựng slide `.pptx → .pdf` ngay trên trình duyệt (`p:sldSz`, nền slide, ảnh định vị `p:pic` và khung văn bản `a:t` tiếng Việt).
- **Trích xuất gói `.ZIP` (`images` & `slides`)**:
  - Khử trùng lặp ảnh nhúng bằng mã băm **SHA-256** hoặc kết xuất từng trang/slide thành bộ ảnh `.JPG` chuẩn **Full HD (1920px)**.
- **Giới hạn an toàn & Huỷ tác vụ (`AbortController`)**: Kiểm tra tệp `<= 500 MB`, chống ZIP bomb (`<= 5.000` mục, tổng giải nén `<= 1.5 GB`), giới hạn điểm ảnh Canvas `<= 64 MP` và hỗ trợ nút **Huỷ tác vụ** tức thì.

### 2. Cơ sở dữ liệu Cục bộ trên Thiết bị (`static/db.js` — Dexie.js IndexedDB)
- Lưu trữ bền vững ngay trên trình duyệt (`navigator.storage.persist()`):
  - `preferences`: Ghi nhớ giao diện Sáng/Tối (`theme`), mức dung lượng mục tiêu (`targetMb`), chế độ nén (`mode`) và định dạng xuất (`outputFormat`).
  - `history`: Lưu tối đa **30 bản ghi** nén/giải nén gần nhất kèm ảnh thu nhỏ xem trước.
- **Sao lưu & Khôi phục dữ liệu (`.json`)**:
  - Hộp thoại **🗂️ Lịch sử & Dữ liệu Offline** cho phép **Xuất dữ liệu (`.json`)**, **Nhập dữ liệu (`.json`)** và **Xoá dữ liệu máy**.
  - Bộ kiểm tra lược đồ nghiêm ngặt (`<= 5 MB`, chặn tuyệt đối `__proto__` / `constructor` / `prototype`, kiểm tra biên số học, lọc ký tự điều khiển và chỉ chấp nhận ảnh xem trước chuẩn nhị phân JPEG `data:image/jpeg;base64,` — chống hoàn toàn SVG/Script XSS).

### 3. Triển khai Tĩnh lên Vercel (Zero Backend)
Tệp `vercel.json` đã được cấu hình sẵn `"buildCommand": ""`, `"outputDirectory": "static"` cùng bộ HTTP Security Headers chặt chẽ (`Content-Security-Policy` cấm toàn bộ inline script, `X-Content-Type-Options: nosniff`, `Referrer-Policy`, và `Cache-Control: no-cache` cho `/sw.js`). Chỉ cần kết nối kho mã nguồn với Vercel là ứng dụng hoạt động 100% độc lập.

---

## Cấu trúc dự án

```text
slide-pdf-compressor/
├── vercel.json               # Cấu hình triển khai Static PWA trực tiếp lên Vercel + CSP Headers
├── static/                   # 100% Standalone Offline-First Static PWA Root
│   ├── index.html            # Giao diện chính (đường dẫn tương đối sạch, không inline script)
│   ├── manifest.json         # Web App Manifest chuẩn cài đặt PWA
│   ├── sw.js                 # Service Worker hỗ trợ hoạt động Offline 100%
│   ├── fonts.css             # Khai báo @font-face Be Vietnam Pro tự lưu trữ (.woff2)
│   ├── style.css             # Giao diện Glassmorphism Sáng / Tối
│   ├── db.js                 # Cơ sở dữ liệu IndexedDB (Dexie.js) & trình kiểm duyệt sao lưu .json
│   ├── offline-engine.js     # Engine nén & giải nén PDF / PPTX thuần trình duyệt
│   ├── app.js                # Điều phối UI, AbortController, Lịch sử Offline & fallback .PPT cục bộ
│   ├── fonts/                # Font Be Vietnam Pro (Latin, Latin-Ext, Vietnamese .woff2)
│   ├── icons/                # Icon PWA (icon-192.png, icon-512.png) & favicon.svg
│   └── vendor/               # Thư viện cục bộ: dexie, jszip, pdf-lib, pdf.js + standard_fonts
├── app.py                    # Máy chủ FastAPI cục bộ tùy chọn (đã gia cố 4 lớp bảo mật)
├── core/                     # Engine Python cục bộ tùy chọn (hỗ trợ thêm .ppt 97-2003 qua COM)
└── tests/
    ├── test_compressors.py   # Bộ kiểm thử pytest (nén, trích xuất & kiểm thử bảo mật backend)
    └── verify_offline_pwa.mjs# Bộ kiểm thử Node xác thực Offline PWA, CSP & JSON validator
```

---

## Chạy Máy chủ Python Cục bộ Tùy chọn (Hỗ trợ `.ppt` 97–2003)

```bash
pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8080
```

Hoặc trên Windows, nhấp đúp `run.bat` để khởi chạy tại `http://127.0.0.1:8080`. Có thể tùy biến danh sách CORS qua biến môi trường `ALLOWED_ORIGINS`.

## Kiểm thử Tự động (Pytest + Node Verification)

```bash
python -m pytest tests/test_compressors.py -v
node tests/verify_offline_pwa.mjs
```
