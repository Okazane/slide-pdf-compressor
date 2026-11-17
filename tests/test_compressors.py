import io
import os
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pypdf
import pypdfium2 as pdfium
import pytest
from fastapi.testclient import TestClient
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt
from pypdf.generic import (
    ArrayObject,
    DecodedStreamObject,
    DictionaryObject,
    NameObject,
)


# ---------------------------------------------------------------------------
# Helpers for generating synthetic test files
# ---------------------------------------------------------------------------


def _make_photographic_image(width: int = 1920, height: int = 1080, mode: str = "RGB", seed_byte: int = 42) -> Image.Image:
    """Create a realistic high-entropy image that starts large at high quality and compresses smoothly."""
    # Generate smooth multi-scale noise + shapes so JPEG/PNG sizes are large (~1MB+) and realistic
    small_w, small_h = 240, 135
    channels = 4 if mode == "RGBA" else 3
    raw = bytearray(os.urandom(small_w * small_h * channels))
    if mode == "RGBA":
        # Make half opaque and half semi-transparent or full alpha
        for idx in range(3, len(raw), 4):
            raw[idx] = 255 if (idx % 8 == 3) else 220
    base = Image.frombytes(mode, (small_w, small_h), bytes(raw))
    img = base.resize((width, height), Image.Resampling.BILINEAR)
    draw = ImageDraw.Draw(img)
    draw.rectangle([100, 100, width - 100, height - 100], outline=(255, 255, 255) if mode == "RGB" else (255, 255, 255, 255), width=6)
    return img


def _add_vietnamese_text_to_pdf_page(writer: pypdf.PdfWriter, page: pypdf.PageObject, vn_text: str) -> None:
    """Attach a Type0 Identity-H font with a ToUnicode CMap so pypdf.PdfReader extracts exact Vietnamese Unicode."""
    chars = list(dict.fromkeys(vn_text))
    bfchar_lines = []
    cid_hex = ""
    for idx, ch in enumerate(chars, start=1):
        bfchar_lines.append(f"<{idx:04X}> <{ord(ch):04X}>")
    for ch in vn_text:
        idx = chars.index(ch) + 1
        cid_hex += f"{idx:04X}"

    cmap_str = (
        "/CIDInit /ProcSet findresource begin\n"
        "12 dict begin\n"
        "begincmap\n"
        "/CMapName /CustomUTF16 def\n"
        "/CMapType 2 def\n"
        "1 begincodespacerange\n"
        "<0000> <FFFF>\n"
        "endcodespacerange\n"
        f"{len(bfchar_lines)} beginbfchar\n"
        + "\n".join(bfchar_lines)
        + "\nendbfchar\n"
        "endcmap\n"
        "CMapName currentdict /CMap defineresource pop\n"
        "end\nend\n"
    )
    cmap_stream = DecodedStreamObject()
    cmap_stream.set_data(cmap_str.encode("ascii"))

    font_dict = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type0"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
            NameObject("/Encoding"): NameObject("/Identity-H"),
            NameObject("/ToUnicode"): writer._add_object(cmap_stream),
            NameObject("/DescendantFonts"): ArrayObject([]),
        }
    )
    res = page[NameObject("/Resources")]
    if NameObject("/Font") not in res:
        res[NameObject("/Font")] = DictionaryObject()
    res[NameObject("/Font")][NameObject("/F1")] = writer._add_object(font_dict)

    old_contents = page.get_contents()
    old_data = old_contents.get_data() if old_contents is not None else b""
    new_stream = DecodedStreamObject()
    new_stream.set_data(old_data + f"\nBT /F1 24 Tf 50 50 Td <{cid_hex}> Tj ET\n".encode("ascii"))
    page[NameObject("/Contents")] = writer._add_object(new_stream)


@pytest.fixture(scope="module")
def synthetic_multipage_pdf(tmp_path_factory) -> Path:
    """Create a 6-page high-resolution (~5-6 MB) synthetic slide PDF."""
    tmp_dir = tmp_path_factory.mktemp("synthetic_pdf")
    pdf_path = tmp_dir / "high_res_slides.pdf"
    pages = [_make_photographic_image(1920, 1080, "RGB", i) for i in range(6)]
    pages[0].save(
        str(pdf_path),
        format="PDF",
        save_all=True,
        append_images=pages[1:],
        resolution=150.0,
        quality=98,
    )
    assert pdf_path.exists()
    assert pdf_path.stat().st_size > 2.5 * 1024 * 1024, "Synthetic PDF should exceed 2.5 MB before compression"
    return pdf_path


@pytest.fixture(scope="module")
def synthetic_vector_vn_pdf(tmp_path_factory) -> tuple[Path, list[str]]:
    """Create a 2-page PDF containing embedded RGBA and RGB images plus Vietnamese vector text."""
    tmp_dir = tmp_path_factory.mktemp("vector_pdf")
    pdf_path = tmp_dir / "vietnamese_vector.pdf"

    # Page 1: RGBA image, Page 2: RGB image
    im_rgba = _make_photographic_image(1600, 900, "RGBA", 1)
    im_rgb = _make_photographic_image(1600, 900, "RGB", 2)

    buf1 = io.BytesIO()
    im_rgba.save(buf1, format="PDF", resolution=150.0)
    buf1.seek(0)

    buf2 = io.BytesIO()
    im_rgb.save(buf2, format="PDF", resolution=150.0, quality=98)
    buf2.seek(0)

    r1 = pypdf.PdfReader(buf1)
    r2 = pypdf.PdfReader(buf2)

    writer = pypdf.PdfWriter()
    writer.append(r1)
    writer.append(r2)

    vn_texts = [
        "Nghệ thuật thanh sắc Việt Nam 2026 - Bảo tồn di sản văn hóa",
        "Ứng dụng nén tài liệu thông minh giữ nguyên chữ Vector sắc nét",
    ]
    _add_vietnamese_text_to_pdf_page(writer, writer.pages[0], vn_texts[0])
    _add_vietnamese_text_to_pdf_page(writer, writer.pages[1], vn_texts[1])

    with open(pdf_path, "wb") as f:
        writer.write(f)

    assert pdf_path.exists()
    return pdf_path, vn_texts


@pytest.fixture(scope="module")
def synthetic_pptx(tmp_path_factory) -> Path:
    """Create a 3-slide .pptx presentation with large embedded PNG and JPEG images."""
    tmp_dir = tmp_path_factory.mktemp("synthetic_pptx")
    pptx_path = tmp_dir / "heavy_presentation.pptx"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for idx in range(3):
        slide = prs.slides.add_slide(blank_layout)
        # Add a large PNG image and a large high-quality JPEG image
        img_png = _make_photographic_image(1600, 900, "RGBA" if idx == 0 else "RGB", idx + 10)
        png_buf = io.BytesIO()
        img_png.save(png_buf, format="PNG", compress_level=1)
        png_buf.seek(0)
        slide.shapes.add_picture(png_buf, Inches(0.5), Inches(0.5), width=Inches(6.0))

        img_jpg = _make_photographic_image(1920, 1080, "RGB", idx + 20)
        jpg_buf = io.BytesIO()
        img_jpg.save(jpg_buf, format="JPEG", quality=98)
        jpg_buf.seek(0)
        slide.shapes.add_picture(jpg_buf, Inches(6.8), Inches(0.5), width=Inches(6.0))

        tx_box = slide.shapes.add_textbox(Inches(1), Inches(6.2), Inches(11), Inches(1))
        tf = tx_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Slide thuyết trình số {idx + 1} - Nghệ thuật thanh sắc"
        p.font.size = Pt(24)

    prs.save(str(pptx_path))
    assert pptx_path.exists()
    assert pptx_path.stat().st_size > 2.0 * 1024 * 1024, "Synthetic PPTX should exceed 2 MB"
    return pptx_path


# ---------------------------------------------------------------------------
# 1. Flatten Mode PDF Compression Tests
# ---------------------------------------------------------------------------


def test_flatten_pdf_compression_targets(synthetic_multipage_pdf: Path, tmp_path: Path):
    from core.pdf_compressor import compress_pdf, inspect_pdf

    info = inspect_pdf(str(synthetic_multipage_pdf))
    assert info["page_count"] == 6
    assert info["original_size_mb"] > 2.5
    assert info["preview_base64"].startswith("data:image/jpeg;base64,")

    for target_mb in [2.0, 1.0]:
        out_pdf = tmp_path / f"compressed_flatten_{target_mb}mb.pdf"
        progress_events = []

        def on_progress(curr, total, msg):
            progress_events.append((curr, total, msg))

        res = compress_pdf(
            str(synthetic_multipage_pdf),
            str(out_pdf),
            target_mb=target_mb,
            mode="flatten",
            progress_callback=on_progress,
        )

        assert out_pdf.exists(), f"Output PDF missing for target_mb={target_mb}"
        assert res["page_count"] == 6
        assert res["compressed_size_mb"] < target_mb, (
            f"Expected compressed_size_mb < {target_mb}, got {res['compressed_size_mb']}"
        )
        assert out_pdf.stat().st_size < target_mb * 1024 * 1024
        assert res["reduction_percent"] > 0
        assert res["before_preview_base64"].startswith("data:image/jpeg;base64,")
        assert res["after_preview_base64"].startswith("data:image/jpeg;base64,")
        assert len(progress_events) > 0

        # Verify page count with pypdfium2
        doc = pdfium.PdfDocument(str(out_pdf))
        assert len(doc) == 6
        doc.close()


# ---------------------------------------------------------------------------
# 2. Real File Test (D:\Nghệ thuật thanh sắc.pdf)
# ---------------------------------------------------------------------------


@pytest.mark.skipif(
    not os.path.exists(r"D:\Nghệ thuật thanh sắc.pdf"),
    reason=r"D:\Nghệ thuật thanh sắc.pdf not present on this machine",
)
def test_real_file_nghe_thuat_thanh_sac(tmp_path: Path):
    from core.pdf_compressor import compress_pdf, inspect_pdf

    real_path = r"D:\Nghệ thuật thanh sắc.pdf"
    info = inspect_pdf(real_path)
    assert info["page_count"] == 52
    assert info["original_size_bytes"] > 0
    assert info["preview_base64"].startswith("data:image/jpeg;base64,")

    # Extract a 5-page subset to test fast end-to-end compression on real slides
    subset_path = tmp_path / "real_subset_5pages.pdf"
    reader = pypdf.PdfReader(real_path)
    writer = pypdf.PdfWriter()
    for i in range(5):
        writer.add_page(reader.pages[i])
    with open(subset_path, "wb") as f:
        writer.write(f)

    compressed_subset = tmp_path / "real_subset_compressed.pdf"
    res = compress_pdf(
        str(subset_path),
        str(compressed_subset),
        target_mb=0.6,
        mode="flatten",
    )
    assert compressed_subset.exists()
    assert res["page_count"] == 5
    assert res["compressed_size_mb"] < 0.6


# ---------------------------------------------------------------------------
# 3. Vector Mode PDF Compression with RGBA/RGB Images & Vietnamese Text
# ---------------------------------------------------------------------------


def test_vector_mode_preserves_vietnamese_text(
    synthetic_vector_vn_pdf: tuple[Path, list[str]], tmp_path: Path
):
    from core.pdf_compressor import compress_pdf

    pdf_path, expected_texts = synthetic_vector_vn_pdf
    out_pdf = tmp_path / "compressed_vector.pdf"

    res = compress_pdf(
        str(pdf_path),
        str(out_pdf),
        target_mb=1.0,
        mode="vector",
    )

    assert out_pdf.exists()
    assert res["page_count"] == 2
    assert res["compressed_size_bytes"] < res["original_size_bytes"], (
        f"Vector compression did not reduce size: {res['compressed_size_bytes']} >= {res['original_size_bytes']}"
    )
    assert res["compressed_size_mb"] < 1.0

    # Verify Vietnamese text is still intact and extractable via pypdf.PdfReader
    reader = pypdf.PdfReader(str(out_pdf))
    assert len(reader.pages) == 2
    for idx, expected in enumerate(expected_texts):
        extracted = reader.pages[idx].extract_text()
        assert expected in extracted, (
            f"Vietnamese text lost on page {idx}! Expected {expected!r}, got {extracted!r}"
        )


# ---------------------------------------------------------------------------
# 4. Synthetic PPTX Direct Compression & PPTX-to-PDF Conversion Tests
# ---------------------------------------------------------------------------


def test_pptx_direct_and_pdf_conversion(synthetic_pptx: Path, tmp_path: Path):
    from core.pptx_compressor import (
        compress_pptx_direct,
        convert_pptx_to_compressed_pdf,
        inspect_pptx,
    )

    info = inspect_pptx(str(synthetic_pptx))
    assert info["slide_count"] == 3
    assert info["media_count"] >= 6
    assert info["original_size_mb"] > 2.0

    # 4a. Test compress_pptx_direct
    out_pptx = tmp_path / "compressed_slides.pptx"
    res_pptx = compress_pptx_direct(
        str(synthetic_pptx),
        str(out_pptx),
        target_mb=1.5,
    )
    assert out_pptx.exists()
    assert res_pptx["compressed_size_bytes"] < res_pptx["original_size_bytes"]
    assert res_pptx["reduction_percent"] > 30.0, (
        f"Expected significant PPTX size drop (>30%), got {res_pptx['reduction_percent']}%"
    )

    # Verify the compressed PPTX opens cleanly in python-pptx
    reopened = Presentation(str(out_pptx))
    assert len(reopened.slides) == 3

    # 4b. Test convert_pptx_to_compressed_pdf
    out_pdf = tmp_path / "converted_slides.pdf"
    res_pdf = convert_pptx_to_compressed_pdf(
        str(synthetic_pptx),
        str(out_pdf),
        target_mb=1.5,
        mode="flatten",
    )
    assert out_pdf.exists()
    assert res_pdf["compressed_size_mb"] < 1.5
    doc = pdfium.PdfDocument(str(out_pdf))
    assert len(doc) == 3
    doc.close()


# ---------------------------------------------------------------------------
# 5. FastAPI Endpoints Tests
# ---------------------------------------------------------------------------


def test_fastapi_endpoints(synthetic_multipage_pdf: Path, synthetic_pptx: Path):
    from app import app

    client = TestClient(app)

    # GET /
    r_index = client.get("/")
    assert r_index.status_code == 200
    assert "text/html" in r_index.headers.get("content-type", "")

    # POST /api/analyze with PDF
    with open(synthetic_multipage_pdf, "rb") as f:
        r_analyze = client.post(
            "/api/analyze",
            files={"file": ("Thuyết trình mẫu.pdf", f, "application/pdf")},
        )
    assert r_analyze.status_code == 200, r_analyze.text
    data_analyze = r_analyze.json()
    assert "file_id" in data_analyze
    assert data_analyze.get("page_count") == 6
    file_id = data_analyze["file_id"]

    # POST /api/compress
    r_compress = client.post(
        "/api/compress",
        json={
            "file_id": file_id,
            "target_mb": 1.5,
            "mode": "flatten",
            "output_format": "pdf",
        },
    )
    # Fallback to form data if endpoint expects Form instead of JSON
    if r_compress.status_code in (400, 422):
        r_compress = client.post(
            "/api/compress",
            data={
                "file_id": file_id,
                "target_mb": 1.5,
                "mode": "flatten",
                "output_format": "pdf",
            },
        )
    assert r_compress.status_code == 200, r_compress.text
    job_id = r_compress.json()["job_id"]

    # Poll GET /api/job/{job_id}
    job_data = None
    for _ in range(120):
        r_job = client.get(f"/api/job/{job_id}")
        assert r_job.status_code == 200
        job_data = r_job.json()
        if job_data["status"] in ("done", "error"):
            break
        time.sleep(0.1)

    assert job_data is not None
    assert job_data["status"] == "done", f"Job failed: {job_data}"

    # Check result fields (either top-level or inside job_data['result'])
    res_payload = job_data.get("result") or job_data
    assert res_payload["compressed_size_mb"] < 1.5

    # GET /api/download/{job_id}
    r_dl = client.get(f"/api/download/{job_id}")
    assert r_dl.status_code == 200
    assert r_dl.content[:4] == b"%PDF"
    assert len(r_dl.content) < 1.5 * 1024 * 1024

    # Also test POST /api/analyze with PPTX
    with open(synthetic_pptx, "rb") as f:
        r_pptx = client.post(
            "/api/analyze",
            files={
                "file": (
                    "Bài giảng.pptx",
                    f,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            },
        )
    assert r_pptx.status_code == 200, r_pptx.text
    pptx_info = r_pptx.json()
    assert "file_id" in pptx_info
    assert (pptx_info.get("slide_count") or pptx_info.get("page_count")) == 3
