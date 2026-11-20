import io
import os
import posixpath
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
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
# Process & Workspace Hygiene Helpers / Fixtures
# ---------------------------------------------------------------------------


def _get_powerpnt_pids() -> set[int]:
    """Snapshot running POWERPNT.EXE PIDs on Windows."""
    if os.name != "nt":
        return set()
    try:
        proc = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "@(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -join ','",
            ],
            capture_output=True,
            text=True,
            timeout=10,
        )
        raw = (proc.stdout or "").strip()
        if not raw:
            return set()
        return {int(tok.strip()) for tok in raw.split(",") if tok.strip().isdigit()}
    except Exception:
        return set()


@pytest.fixture(scope="session", autouse=True)
def guard_zero_powerpnt_leaks():
    """Ensure the entire test suite leaves ZERO newly spawned POWERPNT.EXE processes."""
    before_pids = _get_powerpnt_pids()
    yield
    after_pids = _get_powerpnt_pids()
    leaked = after_pids - before_pids
    assert not leaked, f"Leaked POWERPNT.EXE process PIDs detected after test session: {leaked}"


@pytest.fixture(autouse=True)
def cleanup_test_uploads_and_outputs():
    """Automatically remove any files created in uploads/ and outputs/ during a test."""
    uploads_dir = PROJECT_ROOT / "uploads"
    outputs_dir = PROJECT_ROOT / "outputs"
    before_uploads = set(uploads_dir.glob("*")) if uploads_dir.exists() else set()
    before_outputs = set(outputs_dir.glob("*")) if outputs_dir.exists() else set()

    yield

    created_paths: set[str] = set()
    if uploads_dir.exists():
        for p in set(uploads_dir.glob("*")) - before_uploads:
            created_paths.add(str(p.resolve()))
            try:
                p.unlink()
            except OSError:
                pass
    if outputs_dir.exists():
        for p in set(outputs_dir.glob("*")) - before_outputs:
            created_paths.add(str(p.resolve()))
            try:
                p.unlink()
            except OSError:
                pass

    # Prune in-memory app registries referencing cleaned-up test files
    if "app" in sys.modules:
        app_mod = sys.modules["app"]
        files_reg = getattr(app_mod, "FILES", {})
        jobs_reg = getattr(app_mod, "JOBS", {})
        for fid in list(files_reg.keys()):
            fpath = files_reg[fid].get("path")
            if fpath and (not Path(fpath).exists() or str(Path(fpath).resolve()) in created_paths):
                files_reg.pop(fid, None)
        for jid in list(jobs_reg.keys()):
            opath = jobs_reg[jid].get("output_path")
            if opath and (not Path(opath).exists() or str(Path(opath).resolve()) in created_paths):
                jobs_reg.pop(jid, None)


# ---------------------------------------------------------------------------
# Helpers for generating synthetic test files & validating OPC packages
# ---------------------------------------------------------------------------


def _make_photographic_image(
    width: int = 1920, height: int = 1080, mode: str = "RGB", seed_byte: int = 42
) -> Image.Image:
    """Create a realistic high-entropy image that starts large at high quality and compresses smoothly."""
    small_w, small_h = 240, 135
    channels = 4 if mode == "RGBA" else 3
    raw = bytearray(os.urandom(small_w * small_h * channels))
    if mode == "RGBA":
        for idx in range(3, len(raw), 4):
            raw[idx] = 255 if (idx % 8 == 3) else 220
    base = Image.frombytes(mode, (small_w, small_h), bytes(raw))
    img = base.resize((width, height), Image.Resampling.BILINEAR)
    draw = ImageDraw.Draw(img)
    draw.rectangle(
        [100, 100, width - 100, height - 100],
        outline=(255, 255, 255) if mode == "RGB" else (255, 255, 255, 255),
        width=6,
    )
    return img


def _add_vietnamese_text_to_pdf_page(
    writer: pypdf.PdfWriter, page: pypdf.PageObject, vn_text: str
) -> None:
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
    if NameObject("/Resources") not in page:
        page[NameObject("/Resources")] = DictionaryObject()
    res = page[NameObject("/Resources")]
    if NameObject("/Font") not in res:
        res[NameObject("/Font")] = DictionaryObject()
    res[NameObject("/Font")][NameObject("/F1")] = writer._add_object(font_dict)

    old_contents = page.get_contents()
    old_data = old_contents.get_data() if old_contents is not None else b""
    new_stream = DecodedStreamObject()
    new_stream.set_data(
        old_data + f"\nBT /F1 24 Tf 50 50 Td <{cid_hex}> Tj ET\n".encode("ascii")
    )
    page[NameObject("/Contents")] = writer._add_object(new_stream)


def _verify_pptx_opc_relationships(pptx_path: Path) -> None:
    """Verify every internal Target in all *.rels exists in the zip and [Content_Types].xml is valid."""
    assert zipfile.is_zipfile(str(pptx_path)), f"{pptx_path} is not a valid ZIP archive"
    with zipfile.ZipFile(str(pptx_path), "r") as zf:
        names_set = set(zf.namelist())
        assert "[Content_Types].xml" in names_set

        # Validate [Content_Types].xml
        ct_root = ET.fromstring(zf.read("[Content_Types].xml"))
        default_exts: set[str] = set()
        override_parts: set[str] = set()
        for elem in ct_root:
            local_tag = elem.tag.split("}")[-1]
            if local_tag == "Default":
                ext = (elem.attrib.get("Extension") or "").lower().lstrip(".")
                if ext:
                    default_exts.add(ext)
            elif local_tag == "Override":
                part_name = (elem.attrib.get("PartName") or "").lstrip("/")
                if part_name:
                    override_parts.add(part_name)
                    assert part_name in names_set, (
                        f"Stale Override PartName in [Content_Types].xml missing from zip: {part_name}"
                    )

        for entry_name in names_set:
            if entry_name.startswith("ppt/media/") and not entry_name.endswith("/"):
                ext = Path(entry_name).suffix.lower().lstrip(".")
                assert ext in default_exts or entry_name in override_parts, (
                    f"Media entry {entry_name} has no ContentType registration in [Content_Types].xml"
                )

        # Validate all *.rels files
        rels_files = [n for n in names_set if n.endswith(".rels")]
        assert len(rels_files) > 0
        for rels_name in rels_files:
            rels_root = ET.fromstring(zf.read(rels_name))
            # e.g. "ppt/slides/_rels/slide1.xml.rels" -> base_dir = "ppt/slides"
            # e.g. "_rels/.rels" -> base_dir = ""
            rels_dir = posixpath.dirname(rels_name)
            base_dir = posixpath.dirname(rels_dir) if rels_dir.endswith("_rels") else rels_dir

            for rel in rels_root:
                if not rel.tag.endswith("Relationship"):
                    continue
                target_mode = (rel.attrib.get("TargetMode") or "").strip()
                if target_mode.lower() == "external":
                    continue
                target = (rel.attrib.get("Target") or "").strip()
                if not target:
                    continue
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    resolved = posixpath.normpath(posixpath.join(base_dir, target))
                assert resolved in names_set, (
                    f"Broken relationship in {rels_name}: Target={target!r} resolved to {resolved!r}, not found in zip!"
                )


def _extract_all_pptx_texts(prs: Presentation) -> list[str]:
    """Extract text from all shapes in a Presentation, including nested GroupShapes."""
    texts: list[str] = []

    def _visit_shapes(shapes):
        for sh in shapes:
            if getattr(sh, "has_text_frame", False):
                t = sh.text_frame.text.strip()
                if t:
                    texts.append(t)
            if hasattr(sh, "shapes"):
                _visit_shapes(sh.shapes)

    for slide in prs.slides:
        _visit_shapes(slide.shapes)
    return texts


def _poll_job_until_done(client: TestClient, job_id: str, max_wait_sec: float = 20.0) -> dict:
    deadline = time.time() + max_wait_sec
    job_data = None
    while time.time() < deadline:
        r_job = client.get(f"/api/job/{job_id}")
        assert r_job.status_code == 200, r_job.text
        job_data = r_job.json()
        if job_data["status"] in ("done", "error"):
            return job_data
        time.sleep(0.1)
    pytest.fail(f"Job {job_id} timed out; last state: {job_data}")


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
    assert (
        pdf_path.stat().st_size > 2.5 * 1024 * 1024
    ), "Synthetic PDF should exceed 2.5 MB before compression"
    return pdf_path


@pytest.fixture(scope="module")
def synthetic_vector_vn_pdf(tmp_path_factory) -> tuple[Path, list[str]]:
    """Create a 2-page PDF containing embedded RGBA and RGB images plus Vietnamese vector text."""
    tmp_dir = tmp_path_factory.mktemp("vector_pdf")
    pdf_path = tmp_dir / "vietnamese_vector.pdf"

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
        img_png = _make_photographic_image(
            1600, 900, "RGBA" if idx == 0 else "RGB", idx + 10
        )
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
# 1. Flatten Mode PDF Compression & Page-Dimension Fidelity Tests
# ---------------------------------------------------------------------------


def test_flatten_pdf_compression_targets(synthetic_multipage_pdf: Path, tmp_path: Path):
    from core.pdf_compressor import compress_pdf, inspect_pdf

    info = inspect_pdf(str(synthetic_multipage_pdf))
    assert info["page_count"] == 6
    assert info["original_size_mb"] > 2.5
    assert info["media_count"] == 6
    assert info["preview_base64"].startswith("data:image/jpeg;base64,")

    orig_doc = pdfium.PdfDocument(str(synthetic_multipage_pdf))
    orig_sizes = [orig_doc[i].get_size() for i in range(len(orig_doc))]
    orig_doc.close()

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
        assert res["target_met"] is True
        assert res["compressed_size_mb"] < target_mb, (
            f"Expected compressed_size_mb < {target_mb}, got {res['compressed_size_mb']}"
        )
        assert out_pdf.stat().st_size < target_mb * 1024 * 1024
        assert res["reduction_percent"] > 0
        assert res["before_preview_base64"].startswith("data:image/jpeg;base64,")
        assert res["after_preview_base64"].startswith("data:image/jpeg;base64,")
        assert len(progress_events) > 0

        # Verify page count AND exact physical page dimensions within 0.5 pt
        doc = pdfium.PdfDocument(str(out_pdf))
        try:
            assert len(doc) == 6
            for idx in range(6):
                out_w, out_h = doc[idx].get_size()
                in_w, in_h = orig_sizes[idx]
                assert abs(out_w - in_w) <= 0.5 and abs(out_h - in_h) <= 0.5, (
                    f"Page {idx} size mismatch: expected ({in_w}, {in_h}), got ({out_w}, {out_h})"
                )
        finally:
            doc.close()


def test_flatten_mixed_page_dimensions_fidelity(tmp_path: Path):
    """Verify mixed-page-size PDFs preserve every page's exact physical (width_pt, height_pt) within 0.5 pt."""
    from core.pdf_compressor import compress_pdf

    expected_dims = [
        (921.6, 518.4),  # 16:9 widescreen slide
        (595.0, 842.0),  # A4 portrait
        (842.0, 595.0),  # A4 landscape
    ]
    mixed_pdf = tmp_path / "mixed_dimensions.pdf"
    writer = pypdf.PdfWriter()

    for idx, (w_pt, h_pt) in enumerate(expected_dims):
        px_w = 1920 if w_pt >= h_pt else 1240
        px_h = 1080 if w_pt >= h_pt else 1754
        im = _make_photographic_image(px_w, px_h, "RGB", seed_byte=idx + 50)
        buf = io.BytesIO()
        im.save(buf, format="PDF", resolution=150.0, quality=98)
        buf.seek(0)
        page = pypdf.PdfReader(buf).pages[0]
        added_page = writer.add_page(page)
        added_page.scale_to(w_pt, h_pt)

    with open(mixed_pdf, "wb") as f:
        writer.write(f)

    in_doc = pdfium.PdfDocument(str(mixed_pdf))
    in_sizes = [in_doc[i].get_size() for i in range(len(in_doc))]
    in_doc.close()

    for target_mb in [0.8, 5.0]:
        out_pdf = tmp_path / f"mixed_out_{target_mb}mb.pdf"
        res = compress_pdf(str(mixed_pdf), str(out_pdf), target_mb=target_mb, mode="flatten")
        assert out_pdf.exists()
        assert res["page_count"] == 3
        assert res["target_met"] is True

        out_doc = pdfium.PdfDocument(str(out_pdf))
        try:
            assert len(out_doc) == 3
            for i, (exp_w, exp_h) in enumerate(in_sizes):
                got_w, got_h = out_doc[i].get_size()
                assert abs(got_w - exp_w) <= 0.5, (
                    f"Page {i} width distorted for target_mb={target_mb}: expected {exp_w}, got {got_w}"
                )
                assert abs(got_h - exp_h) <= 0.5, (
                    f"Page {i} height distorted for target_mb={target_mb}: expected {exp_h}, got {got_h}"
                )
        finally:
            out_doc.close()


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
    assert "media_count" in info
    assert info["preview_base64"].startswith("data:image/jpeg;base64,")

    subset_path = tmp_path / "real_subset_5pages.pdf"
    reader = pypdf.PdfReader(real_path)
    writer = pypdf.PdfWriter()
    for i in range(5):
        writer.add_page(reader.pages[i])
    with open(subset_path, "wb") as f:
        writer.write(f)

    in_doc = pdfium.PdfDocument(str(subset_path))
    in_sizes = [in_doc[i].get_size() for i in range(len(in_doc))]
    in_doc.close()

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
    assert res["target_met"] is True

    out_doc = pdfium.PdfDocument(str(compressed_subset))
    try:
        for i, (ew, eh) in enumerate(in_sizes):
            gw, gh = out_doc[i].get_size()
            assert abs(gw - ew) <= 0.5 and abs(gh - eh) <= 0.5
    finally:
        out_doc.close()


# ---------------------------------------------------------------------------
# 3. Vector Mode PDF Compression, Non-Inflation & Vietnamese Text
# ---------------------------------------------------------------------------


def test_vector_mode_preserves_vietnamese_text(
    synthetic_vector_vn_pdf: tuple[Path, list[str]], tmp_path: Path
):
    from core.pdf_compressor import compress_pdf, inspect_pdf

    pdf_path, expected_texts = synthetic_vector_vn_pdf
    info = inspect_pdf(str(pdf_path))
    assert info["media_count"] >= 2

    out_pdf = tmp_path / "compressed_vector.pdf"
    res = compress_pdf(
        str(pdf_path),
        str(out_pdf),
        target_mb=1.0,
        mode="vector",
    )

    assert out_pdf.exists()
    assert res["page_count"] == 2
    assert res["target_met"] is True
    assert res["compressed_size_bytes"] < res["original_size_bytes"], (
        f"Vector compression did not reduce size: {res['compressed_size_bytes']} >= {res['original_size_bytes']}"
    )
    assert res["compressed_size_mb"] < 1.0

    reader = pypdf.PdfReader(str(out_pdf))
    assert len(reader.pages) == 2
    for idx, expected in enumerate(expected_texts):
        extracted = reader.pages[idx].extract_text()
        assert expected in extracted, (
            f"Vietnamese text lost on page {idx}! Expected {expected!r}, got {extracted!r}"
        )


def test_pdf_non_inflation_and_target_met_flags(
    synthetic_multipage_pdf: Path, tmp_path: Path
):
    """Verify already-small PDFs never inflate when target_mb is generous, and unreachable targets set target_met=False."""
    from core.pdf_compressor import compress_pdf

    # Create an already-compact 2-page PDF with low-quality JPEGs + Vietnamese Unicode vector text
    compact_pdf = tmp_path / "already_small_vn.pdf"
    vn_lines = [
        "Trường Đại học Bách Khoa - Tài liệu học thuật chuẩn Vector",
        "Không bao giờ làm phình dung lượng tệp PDF đã tối ưu sẵn",
    ]
    writer = pypdf.PdfWriter()
    for idx in range(2):
        small_im = _make_photographic_image(480, 270, "RGB", seed_byte=idx + 80)
        buf = io.BytesIO()
        small_im.save(buf, format="PDF", resolution=72.0, quality=45)
        buf.seek(0)
        writer.append(pypdf.PdfReader(buf))
        _add_vietnamese_text_to_pdf_page(writer, writer.pages[idx], vn_lines[idx])

    with open(compact_pdf, "wb") as f:
        writer.write(f)
    orig_bytes = compact_pdf.stat().st_size
    assert orig_bytes < 512 * 1024

    # 1. Vector mode with generous target_mb=10.0 -> must NOT inflate, preserves Vietnamese text
    out_vec = tmp_path / "non_inflated_vector.pdf"
    res_vec = compress_pdf(str(compact_pdf), str(out_vec), target_mb=10.0, mode="vector")
    assert res_vec["compressed_size_bytes"] <= orig_bytes, (
        f"Vector mode inflated already-small PDF: {res_vec['compressed_size_bytes']} > {orig_bytes}"
    )
    assert out_vec.stat().st_size == res_vec["compressed_size_bytes"]
    assert res_vec["target_met"] is True

    reader = pypdf.PdfReader(str(out_vec))
    for idx, expected in enumerate(vn_lines):
        assert expected in reader.pages[idx].extract_text()

    # 2. Flatten mode with generous target_mb=10.0 -> must NOT inflate either
    out_flat = tmp_path / "non_inflated_flatten.pdf"
    res_flat = compress_pdf(str(compact_pdf), str(out_flat), target_mb=10.0, mode="flatten")
    assert res_flat["compressed_size_bytes"] <= orig_bytes, (
        f"Flatten mode inflated already-small PDF: {res_flat['compressed_size_bytes']} > {orig_bytes}"
    )
    assert res_flat["target_met"] is True

    # 3. Unreachable tiny target (1 KB = ~0.001 MB) on 6-page high-res PDF -> target_met must be False
    out_tiny = tmp_path / "unreachable_tiny.pdf"
    res_tiny = compress_pdf(
        str(synthetic_multipage_pdf), str(out_tiny), target_mb=0.001, mode="flatten"
    )
    assert out_tiny.exists()
    assert res_tiny["target_met"] is False


# ---------------------------------------------------------------------------
# 4. PPTX Multi-Format Media, Opaque Conversion, GroupShape & .PPT Safety
# ---------------------------------------------------------------------------


def test_pptx_direct_and_pdf_conversion(synthetic_pptx: Path, tmp_path: Path):
    from core.pptx_compressor import (
        compress_pptx_direct,
        convert_pptx_to_compressed_pdf,
        inspect_pptx,
    )

    pids_before = _get_powerpnt_pids()

    info = inspect_pptx(str(synthetic_pptx))
    assert info["slide_count"] == 3
    assert info["media_count"] >= 6
    assert info["original_size_mb"] > 2.0
    assert info["is_legacy_ppt"] is False

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
    assert "target_met" in res_pptx
    _verify_pptx_opc_relationships(out_pptx)

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
    assert res_pdf["target_met"] is True
    doc = pdfium.PdfDocument(str(out_pdf))
    assert len(doc) == 3
    doc.close()

    pids_after = _get_powerpnt_pids()
    assert pids_after - pids_before == set(), (
        f"POWERPNT.EXE leaked during PPTX conversion: {pids_after - pids_before}"
    )


def test_pptx_multiformat_opaque_png_bmp_tiff_groupshape_and_rels(tmp_path: Path):
    """Test opaque PNG/BMP/TIFF -> JPEG smart conversion, transparent PNG preservation, GroupShape Vietnamese text, and XML rels."""
    from core.pptx_compressor import compress_pptx_direct, inspect_pptx

    pptx_in = tmp_path / "multiformat_deck.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1: Opaque large PNG + Transparent RGBA PNG overlay + GroupShape with Vietnamese text
    s1 = prs.slides.add_slide(blank)
    opaque_png = _make_photographic_image(1600, 900, "RGB", seed_byte=101)
    b_opaque = io.BytesIO()
    opaque_png.save(b_opaque, format="PNG", compress_level=1)
    b_opaque.seek(0)
    s1.shapes.add_picture(b_opaque, Inches(0.5), Inches(0.5), width=Inches(6.0))

    trans_im = Image.new("RGBA", (500, 350), (0, 0, 0, 0))
    draw_t = ImageDraw.Draw(trans_im)
    draw_t.ellipse([40, 40, 460, 310], fill=(30, 140, 240, 160), outline=(255, 255, 255, 255), width=5)
    b_trans = io.BytesIO()
    trans_im.save(b_trans, format="PNG")
    b_trans.seek(0)
    s1.shapes.add_picture(b_trans, Inches(4.5), Inches(1.5), width=Inches(3.5))

    grp = s1.shapes.add_group_shape()
    grp_tb = grp.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(10.0), Inches(1.0))
    grp_tb.text_frame.text = "Nhóm biểu đồ & chữ Tiếng Việt trong GroupShape"

    # Slide 2: Uncompressed BMP + Uncompressed TIFF + Vietnamese textbox
    s2 = prs.slides.add_slide(blank)
    bmp_im = _make_photographic_image(1280, 720, "RGB", seed_byte=102)
    b_bmp = io.BytesIO()
    bmp_im.save(b_bmp, format="BMP")
    b_bmp.seek(0)
    s2.shapes.add_picture(b_bmp, Inches(0.5), Inches(0.5), width=Inches(5.8))

    tiff_im = _make_photographic_image(1280, 720, "RGB", seed_byte=103)
    b_tiff = io.BytesIO()
    tiff_im.save(b_tiff, format="TIFF")
    b_tiff.seek(0)
    s2.shapes.add_picture(b_tiff, Inches(6.8), Inches(0.5), width=Inches(5.8))

    tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(10.0), Inches(1.0))
    tb2.text_frame.text = "Tối ưu hóa ảnh BMP và TIFF giữ nguyên liên kết XML"

    prs.save(str(pptx_in))
    assert pptx_in.stat().st_size > 5 * 1024 * 1024

    info = inspect_pptx(str(pptx_in))
    assert info["slide_count"] == 2
    assert info["media_count"] == 4
    assert info["is_legacy_ppt"] is False
    assert info.get("preview_base64", "").startswith("data:image/jpeg;base64,")

    pptx_out = tmp_path / "multiformat_compressed.pptx"
    res = compress_pptx_direct(str(pptx_in), str(pptx_out), target_mb=1.5)

    assert pptx_out.exists()
    assert res["target_met"] is True
    assert res["reduction_percent"] > 65.0, (
        f"Expected >65% reduction on uncompressed PNG/BMP/TIFF PPTX, got {res['reduction_percent']}%"
    )

    # Validate OPC XML relationships and [Content_Types].xml
    _verify_pptx_opc_relationships(pptx_out)

    # Check media entries: opaque images converted to .jpg and transparent .png preserved with alpha
    with zipfile.ZipFile(str(pptx_out), "r") as zf:
        media_files = [n for n in zf.namelist() if n.startswith("ppt/media/") and not n.endswith("/")]
        assert len(media_files) == 4
        jpg_entries = [n for n in media_files if n.lower().endswith(".jpg")]
        png_entries = [n for n in media_files if n.lower().endswith(".png")]
        assert len(jpg_entries) >= 2, f"Expected opaque PNG/BMP/TIFF to convert to .jpg, got {media_files}"
        assert len(png_entries) >= 1, f"Expected transparent PNG to remain .png, got {media_files}"

        has_transparent_png = False
        for png_name in png_entries:
            im_check = Image.open(io.BytesIO(zf.read(png_name)))
            if im_check.mode in ("RGBA", "LA") or ("transparency" in im_check.info):
                rgba_check = im_check.convert("RGBA")
                if rgba_check.getchannel("A").getextrema()[0] < 255:
                    has_transparent_png = True
        assert has_transparent_png, "Transparent PNG overlay lost its alpha channel!"

    # Confirm output PPTX opens cleanly in python-pptx and preserves GroupShape Vietnamese text + image blobs
    reopened = Presentation(str(pptx_out))
    assert len(reopened.slides) == 2
    extracted_texts = _extract_all_pptx_texts(reopened)
    assert any("Nhóm biểu đồ & chữ Tiếng Việt trong GroupShape" in t for t in extracted_texts)
    assert any("Tối ưu hóa ảnh BMP và TIFF giữ nguyên liên kết XML" in t for t in extracted_texts)

    for slide in reopened.slides:
        for shape in slide.shapes:
            if hasattr(shape, "image"):
                assert len(shape.image.blob) > 0

    # Non-inflation check when compressing an already-small PPTX with huge target_mb
    pptx_recomp = tmp_path / "multiformat_recompressed.pptx"
    res_recomp = compress_pptx_direct(str(pptx_out), str(pptx_recomp), target_mb=50.0)
    assert res_recomp["compressed_size_bytes"] <= pptx_out.stat().st_size
    assert res_recomp["target_met"] is True


def test_legacy_ppt_handling_and_zero_powerpnt_leaks(tmp_path: Path):
    """Verify binary .ppt handling across core and API with zero BadZipFile crashes and zero POWERPNT leaks."""
    from app import app
    from core.extractor import extract_archive
    from core.pptx_compressor import compress_pptx_direct, inspect_pptx

    pids_before = _get_powerpnt_pids()

    dummy_ppt = tmp_path / "legacy_binary.ppt"
    dummy_ppt.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)

    corrupt_pptx = tmp_path / "corrupt_archive.pptx"
    corrupt_pptx.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)

    # 1. inspect_pptx marks is_legacy_ppt=True without crashing
    ppt_meta = inspect_pptx(str(dummy_ppt))
    assert ppt_meta["is_legacy_ppt"] is True
    assert ppt_meta["original_size_bytes"] == 520

    # Corrupt .pptx raises clean ValueError
    with pytest.raises(ValueError):
        inspect_pptx(str(corrupt_pptx))

    # 2. compress_pptx_direct raises clean Vietnamese ValueError (never zipfile.BadZipFile)
    with pytest.raises(ValueError) as exc_info:
        compress_pptx_direct(str(dummy_ppt), str(tmp_path / "should_not_exist.pptx"), target_mb=5.0)
    assert not isinstance(exc_info.value, zipfile.BadZipFile)
    assert ".ppt" in str(exc_info.value).lower() or "pdf" in str(exc_info.value).lower()

    # 3. extract_archive with extract_mode="images" on .ppt raises clean Vietnamese ValueError
    with pytest.raises(ValueError) as ext_exc:
        extract_archive(
            str(dummy_ppt),
            str(tmp_path / "should_not_exist.zip"),
            file_type="ppt",
            extract_mode="images",
        )
    assert not isinstance(ext_exc.value, zipfile.BadZipFile)

    # 4. FastAPI /api/analyze, /api/compress, and /api/extract for legacy .ppt
    client = TestClient(app)
    with open(dummy_ppt, "rb") as f:
        r_analyze = client.post(
            "/api/analyze",
            files={"file": ("Bài_giảng_cũ_2003.ppt", f, "application/vnd.ms-powerpoint")},
        )
    assert r_analyze.status_code == 200, r_analyze.text
    analyze_data = r_analyze.json()
    assert analyze_data["file_type"] == "ppt"
    assert analyze_data.get("is_legacy_ppt") is True
    ppt_file_id = analyze_data["file_id"]

    # /api/compress with output_format="pptx" -> HTTP 400 cleanly
    r_comp_bad = client.post(
        "/api/compress",
        json={"file_id": ppt_file_id, "target_mb": 5.0, "output_format": "pptx"},
    )
    assert r_comp_bad.status_code == 400, r_comp_bad.text

    # /api/extract with extract_mode="images" -> HTTP 400 cleanly
    r_ext_bad = client.post(
        "/api/extract",
        json={"file_id": ppt_file_id, "extract_mode": "images"},
    )
    assert r_ext_bad.status_code == 400, r_ext_bad.text

    pids_after = _get_powerpnt_pids()
    assert pids_after - pids_before == set(), (
        f"Leaked POWERPNT.EXE processes during .ppt tests: {pids_after - pids_before}"
    )


# ---------------------------------------------------------------------------
# 5. Resource Extraction Tests (core.extractor & POST /api/extract)
# ---------------------------------------------------------------------------


def test_extract_archive_core_pdf_and_pptx(
    synthetic_vector_vn_pdf: tuple[Path, list[str]],
    synthetic_pptx: Path,
    tmp_path: Path,
):
    from core import extract_archive

    pdf_path, _ = synthetic_vector_vn_pdf

    # 5a. PDF extract_mode="images" (2 unique embedded images)
    pdf_imgs_zip = tmp_path / "pdf_images.zip"
    res_pdf_imgs = extract_archive(
        str(pdf_path), str(pdf_imgs_zip), file_type="pdf", extract_mode="images"
    )
    assert pdf_imgs_zip.exists()
    assert res_pdf_imgs["extracted_count"] == 2
    assert res_pdf_imgs["zip_size_bytes"] > 0
    assert res_pdf_imgs["extract_mode"] == "images"
    assert res_pdf_imgs["preview_base64"].startswith("data:image/jpeg;base64,")
    with zipfile.ZipFile(str(pdf_imgs_zip), "r") as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        assert len(names) == 2
        assert all(n.startswith("images/") for n in names)

    # 5b. PDF extract_mode="images" fallback on vector-only PDF with 0 embedded images -> pages/trang_001.jpg
    blank_pdf = tmp_path / "vector_only_no_imgs.pdf"
    w_blank = pypdf.PdfWriter()
    w_blank.add_blank_page(width=595.0, height=842.0)
    w_blank.add_blank_page(width=595.0, height=842.0)
    with open(blank_pdf, "wb") as f:
        w_blank.write(f)

    fallback_zip = tmp_path / "pdf_fallback_pages.zip"
    res_fallback = extract_archive(
        str(blank_pdf), str(fallback_zip), file_type="pdf", extract_mode="images"
    )
    assert res_fallback["extracted_count"] == 2
    with zipfile.ZipFile(str(fallback_zip), "r") as zf:
        fb_names = sorted(n for n in zf.namelist() if not n.endswith("/"))
        assert fb_names == ["pages/trang_001.jpg", "pages/trang_002.jpg"]

    # 5c. PDF extract_mode="slides" -> slides/trang_001.jpg, slides/trang_002.jpg
    pdf_slides_zip = tmp_path / "pdf_slides.zip"
    res_pdf_slides = extract_archive(
        str(pdf_path), str(pdf_slides_zip), file_type="pdf", extract_mode="slides"
    )
    assert res_pdf_slides["extracted_count"] == 2
    with zipfile.ZipFile(str(pdf_slides_zip), "r") as zf:
        s_names = sorted(n for n in zf.namelist() if not n.endswith("/"))
        assert s_names == ["slides/trang_001.jpg", "slides/trang_002.jpg"]

    # 5d. PPTX extract_mode="images" (6 unique images)
    pptx_imgs_zip = tmp_path / "pptx_images.zip"
    res_pptx_imgs = extract_archive(
        str(synthetic_pptx), str(pptx_imgs_zip), file_type="pptx", extract_mode="images"
    )
    assert res_pptx_imgs["extracted_count"] == 6
    with zipfile.ZipFile(str(pptx_imgs_zip), "r") as zf:
        p_names = [n for n in zf.namelist() if not n.endswith("/")]
        assert len(p_names) == 6
        assert all(n.startswith("images/") for n in p_names)

    # 5e. PPTX extract_mode="slides" (3 slides -> slides/slide_001.jpg..slide_003.jpg)
    pptx_slides_zip = tmp_path / "pptx_slides.zip"
    res_pptx_slides = extract_archive(
        str(synthetic_pptx), str(pptx_slides_zip), file_type="pptx", extract_mode="slides"
    )
    assert res_pptx_slides["extracted_count"] == 3
    with zipfile.ZipFile(str(pptx_slides_zip), "r") as zf:
        ps_names = sorted(n for n in zf.namelist() if not n.endswith("/"))
        assert ps_names == [
            "slides/slide_001.jpg",
            "slides/slide_002.jpg",
            "slides/slide_003.jpg",
        ]

    # 5f. SHA-256 deduplication when identical image bytes appear multiple times in PPTX media
    dup_pptx = tmp_path / "dup_media.pptx"
    with zipfile.ZipFile(str(synthetic_pptx), "r") as zin, zipfile.ZipFile(
        str(dup_pptx), "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        sample_media_bytes = None
        for info in zin.infolist():
            raw = zin.read(info.filename)
            if sample_media_bytes is None and info.filename.startswith("ppt/media/"):
                sample_media_bytes = raw
            zout.writestr(info, raw)
        assert sample_media_bytes is not None
        zout.writestr("ppt/media/duplicate_copy.jpg", sample_media_bytes)

    dup_zip = tmp_path / "dup_extracted.zip"
    res_dup = extract_archive(
        str(dup_pptx), str(dup_zip), file_type="pptx", extract_mode="images"
    )
    assert res_dup["extracted_count"] == 6, "Duplicate media bytes should be deduplicated via SHA-256"


def test_fastapi_extract_endpoints(
    synthetic_vector_vn_pdf: tuple[Path, list[str]], synthetic_pptx: Path
):
    from app import app

    client = TestClient(app)
    pdf_path, _ = synthetic_vector_vn_pdf

    # Upload PDF
    with open(pdf_path, "rb") as f:
        r_pdf = client.post(
            "/api/analyze",
            files={"file": ("Tài_liệu_ảnh.pdf", f, "application/pdf")},
        )
    assert r_pdf.status_code == 200
    pdf_info = r_pdf.json()
    assert pdf_info.get("media_count") == 2
    pdf_fid = pdf_info["file_id"]

    # Upload PPTX
    with open(synthetic_pptx, "rb") as f:
        r_pptx = client.post(
            "/api/analyze",
            files={
                "file": (
                    "Slide_bài_giảng.pptx",
                    f,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            },
        )
    assert r_pptx.status_code == 200
    pptx_fid = r_pptx.json()["file_id"]

    cases = [
        (pdf_fid, "images", 2),
        (pdf_fid, "slides", 2),
        (pptx_fid, "images", 6),
        (pptx_fid, "slides", 3),
    ]
    for fid, ext_mode, expected_count in cases:
        r_ext = client.post(
            "/api/extract",
            json={"file_id": fid, "extract_mode": ext_mode},
        )
        assert r_ext.status_code == 200, r_ext.text
        job_id = r_ext.json()["job_id"]

        job_done = _poll_job_until_done(client, job_id)
        assert job_done["status"] == "done", f"Extract job failed: {job_done}"
        res_payload = job_done.get("result") or job_done
        assert res_payload.get("extracted_count") == expected_count
        assert res_payload.get("output_format") == "zip"

        r_dl = client.get(f"/api/download/{job_id}")
        assert r_dl.status_code == 200
        assert "application/zip" in r_dl.headers.get("content-type", "")
        assert r_dl.content[:4] == b"PK\x03\x04"

        with zipfile.ZipFile(io.BytesIO(r_dl.content), "r") as zf:
            entries = [n for n in zf.namelist() if not n.endswith("/")]
            assert len(entries) == expected_count
            # Verify first image opens cleanly via Pillow
            im0 = Image.open(io.BytesIO(zf.read(entries[0])))
            im0.load()
            assert im0.width > 0 and im0.height > 0


# ---------------------------------------------------------------------------
# 6. FastAPI Compression Endpoints, Target-Met, UI Assets & TTL Cleanup
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

    job_data = _poll_job_until_done(client, job_id)
    assert job_data["status"] == "done", f"Job failed: {job_data}"

    res_payload = job_data.get("result") or job_data
    assert res_payload["compressed_size_mb"] < 1.5
    assert res_payload.get("target_met") is True

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
    assert pptx_info.get("file_type") == "pptx"
    assert (pptx_info.get("slide_count") or pptx_info.get("page_count")) == 3


def test_ui_static_assets_target_met_and_ttl_cleanup(
    synthetic_multipage_pdf: Path, tmp_path: Path, monkeypatch
):
    import app as app_mod

    client = TestClient(app_mod.app)

    # 1. Verify HTML, JS, CSS contain new Extract UI elements, .ppt badge handling & target_met logic
    r_html = client.get("/")
    assert r_html.status_code == 200
    html_text = r_html.text
    for required_id in [
        'id="compressForm"',
        'id="extractForm"',
        'name="extract_mode"',
        'id="startExtractBtn"',
        'id="pdfOutputOptions"',
        'id="pptxOutputOptions"',
        'id="modeConfigGroup"',
        'id="targetPills"',
        'id="targetMbInput"',
        'id="startCompressBtn"',
    ]:
        assert required_id in html_text, f"Missing {required_id} in index.html"

    r_js = client.get("/static/app.js")
    assert r_js.status_code == 200
    js_text = r_js.text
    assert "/api/extract" in js_text
    assert "target_met" in js_text
    assert "ppt-badge" in js_text

    r_css = client.get("/static/style.css")
    assert r_css.status_code == 200
    assert len(r_css.text) > 200

    # 2. Verify /api/compress propagates target_met=False when target_mb is impossibly tiny
    with open(synthetic_multipage_pdf, "rb") as f:
        r_up = client.post(
            "/api/analyze",
            files={"file": ("tiny_target_check.pdf", f, "application/pdf")},
        )
    fid = r_up.json()["file_id"]
    r_comp = client.post(
        "/api/compress",
        json={"file_id": fid, "target_mb": 0.001, "mode": "flatten", "output_format": "pdf"},
    )
    assert r_comp.status_code == 200
    job_done = _poll_job_until_done(client, r_comp.json()["job_id"])
    assert job_done["status"] == "done"
    assert (job_done.get("target_met") is False) or (
        job_done.get("result", {}).get("target_met") is False
    )

    # 3. Verify automatic TTL cleanup (_cleanup_old_files) deletes >2h old files and keeps fresh ones
    iso_uploads = tmp_path / "iso_uploads"
    iso_outputs = tmp_path / "iso_outputs"
    iso_uploads.mkdir()
    iso_outputs.mkdir()
    monkeypatch.setattr(app_mod, "UPLOADS_DIR", iso_uploads)
    monkeypatch.setattr(app_mod, "OUTPUTS_DIR", iso_outputs)

    stale_up = iso_uploads / "stale_upload.pdf"
    stale_out = iso_outputs / "stale_output.pdf"
    fresh_up = iso_uploads / "fresh_upload.pdf"
    stale_up.write_bytes(b"%PDF-1.4 stale")
    stale_out.write_bytes(b"%PDF-1.4 stale out")
    fresh_up.write_bytes(b"%PDF-1.4 fresh")

    old_ts = time.time() - 9000
    os.utime(stale_up, (old_ts, old_ts))
    os.utime(stale_out, (old_ts, old_ts))

    with open(synthetic_multipage_pdf, "rb") as f:
        r_ttl = client.post(
            "/api/analyze",
            files={"file": ("trigger_ttl.pdf", f, "application/pdf")},
        )
    assert r_ttl.status_code == 200
    assert not stale_up.exists(), "TTL cleanup should remove >2h old upload file"
    assert not stale_out.exists(), "TTL cleanup should remove >2h old output file"
    assert fresh_up.exists(), "TTL cleanup must preserve fresh upload file"


# ---------------------------------------------------------------------------
# 7. Security Hardening & Offline PWA Asset Serving Tests
# ---------------------------------------------------------------------------


def test_security_fix1_hex_id_glob_traversal_and_pwa_routes(tmp_path: Path):
    import app as app_mod

    client = TestClient(app_mod.app)

    # /api/health endpoint
    r_health = client.get("/api/health")
    assert r_health.status_code == 200
    hdata = r_health.json()
    assert hdata["status"] == "ok"
    assert hdata["engine"] == "local-python"

    # PWA root static assets served cleanly (including /static/ with html=True)
    for route in [
        "/static/",
        "/manifest.json",
        "/sw.js",
        "/favicon.svg",
        "/fonts.css",
        "/style.css",
        "/db.js",
        "/offline-engine.js",
        "/app.js",
        "/icons/icon-192.png",
        "/icons/icon-512.png",
        "/vendor/dexie.min.js",
    ]:
        resp = client.get(route)
        assert resp.status_code == 200, f"PWA route {route} failed with {resp.status_code}"
    assert "no-cache" in client.get("/sw.js").headers.get("cache-control", "")

    # Reject glob / path traversal / malformed file_id & job_id
    malicious_ids = [
        "*",
        "????????????",
        "[0-9a-f]*",
        "../app.py",
        "..\\app.py",
        "aabbccddeeff/../../secret",
        "AABBCCDDEEFF",
        "short123",
    ]
    for bad_id in malicious_ids:
        r_c = client.post("/api/compress", json={"file_id": bad_id, "target_mb": 5.0})
        assert r_c.status_code == 400, f"Expected 400 for file_id={bad_id!r}, got {r_c.status_code}"

        r_e = client.post("/api/extract", json={"file_id": bad_id, "extract_mode": "images"})
        assert r_e.status_code == 400, f"Expected 400 for file_id={bad_id!r}, got {r_e.status_code}"

    from urllib.parse import quote as url_quote

    for bad_job in ["*", "????????????", "[a-f]*", "AABBCCDDEEFF", "short"]:
        encoded_job = url_quote(bad_job, safe="")
        assert client.get(f"/api/job/{encoded_job}").status_code == 400
        assert client.get(f"/api/download/{encoded_job}").status_code == 400

    # Verify download rejects output_path outside OUTPUTS_DIR even if registry were tampered
    outside_file = tmp_path / "secret_outside.pdf"
    outside_file.write_bytes(b"%PDF-1.4 secret")
    tampered_jid = "deadbeef1234"
    with app_mod.REGISTRY_LOCK:
        app_mod.JOBS[tampered_jid] = {
            "job_id": tampered_jid,
            "file_id": "cafebabe1234",
            "status": "done",
            "output_path": str(outside_file),
            "output_filename": "secret.pdf",
            "created_at": time.time(),
        }
    try:
        r_tampered = client.get(f"/api/download/{tampered_jid}")
        assert r_tampered.status_code == 404
    finally:
        with app_mod.REGISTRY_LOCK:
            app_mod.JOBS.pop(tampered_jid, None)


def test_security_fix2_zip_bomb_path_traversal_and_xml_entity_rejection(
    synthetic_pptx: Path, tmp_path: Path
):
    from app import app
    from core.extractor import extract_archive
    from core.pptx_compressor import (
        _rewrite_content_types_xml,
        _rewrite_rels_xml,
        compress_pptx_direct,
        inspect_pptx,
    )

    assert Image.MAX_IMAGE_PIXELS == 64_000_000

    # 1. XML <!ENTITY / <!DOCTYPE rejection (Billion Laughs / XXE)
    xxe_xml = (
        b'<?xml version="1.0"?>'
        b'<!DOCTYPE lolz [<!ENTITY lol "boom">]>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">&lol;</Types>'
    )
    with pytest.raises(ValueError):
        _rewrite_content_types_xml(xxe_xml, {"ppt/media/a.png": "ppt/media/a.jpg"})
    with pytest.raises(ValueError):
        _rewrite_rels_xml(xxe_xml, "ppt/slides/_rels/slide1.xml.rels", {"ppt/media/a.png": "ppt/media/a.jpg"})

    xxe_pptx = tmp_path / "xxe_bomb.pptx"
    with zipfile.ZipFile(str(synthetic_pptx), "r") as zin, zipfile.ZipFile(
        str(xxe_pptx), "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            if info.filename == "[Content_Types].xml":
                zout.writestr(info.filename, xxe_xml)
            else:
                zout.writestr(info, zin.read(info.filename))

    with pytest.raises(ValueError):
        inspect_pptx(str(xxe_pptx))
    with pytest.raises(ValueError):
        compress_pptx_direct(str(xxe_pptx), str(tmp_path / "no_out.pptx"))

    client = TestClient(app)
    with open(xxe_pptx, "rb") as f:
        r_xxe = client.post(
            "/api/analyze",
            files={"file": ("evil_xxe.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
    assert r_xxe.status_code == 400

    # 2. Zip-Slip path traversal entry inside .pptx
    traversal_pptx = tmp_path / "zip_slip.pptx"
    with zipfile.ZipFile(str(traversal_pptx), "w") as zout:
        zout.writestr("[Content_Types].xml", b"<Types/>")
        zout.writestr("../../evil.txt", b"pwn")

    with pytest.raises(ValueError):
        inspect_pptx(str(traversal_pptx))
    with pytest.raises(ValueError):
        extract_archive(str(traversal_pptx), str(tmp_path / "no_zip.zip"), file_type="pptx")

    # 3. Entry-count ZIP bomb (> 5,000 entries)
    many_entries_pptx = tmp_path / "entry_bomb.pptx"
    with zipfile.ZipFile(str(many_entries_pptx), "w", compression=zipfile.ZIP_STORED) as zout:
        for idx in range(5005):
            zout.writestr(f"ppt/media/tiny_{idx}.txt", b"a")

    with pytest.raises(ValueError):
        inspect_pptx(str(many_entries_pptx))
    with open(many_entries_pptx, "rb") as f:
        r_bomb = client.post(
            "/api/analyze",
            files={"file": ("entry_bomb.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
    assert r_bomb.status_code == 400


def test_security_fix3_and_fix4_bounded_queue_cors_and_sanitized_errors(
    synthetic_multipage_pdf: Path, monkeypatch
):
    import app as app_mod

    client = TestClient(app_mod.app)

    # 1. Safe CORS: external untrusted origin does not get wildcard/credentials reflection
    r_cors = client.options(
        "/api/health",
        headers={
            "Origin": "https://malicious.example.com",
            "Access-Control-Request-Method": "GET",
        },
    )
    assert r_cors.headers.get("access-control-allow-origin") != "*"
    assert r_cors.headers.get("access-control-allow-credentials") != "true"

    # 2. Bounded queue returns HTTP 429 when pending running jobs reach MAX_PENDING_JOBS
    with open(synthetic_multipage_pdf, "rb") as f:
        r_up = client.post(
            "/api/analyze",
            files={"file": ("queue_test.pdf", f, "application/pdf")},
        )
    assert r_up.status_code == 200
    fid = r_up.json()["file_id"]

    monkeypatch.setattr(app_mod, "MAX_PENDING_JOBS", 1)
    dummy_jid = "112233445566"
    with app_mod.REGISTRY_LOCK:
        app_mod.JOBS[dummy_jid] = {
            "job_id": dummy_jid,
            "file_id": fid,
            "status": "running",
            "created_at": time.time(),
        }
    try:
        r_429_comp = client.post("/api/compress", json={"file_id": fid, "target_mb": 2.0})
        assert r_429_comp.status_code == 429
        r_429_ext = client.post("/api/extract", json={"file_id": fid, "extract_mode": "images"})
        assert r_429_ext.status_code == 429
    finally:
        with app_mod.REGISTRY_LOCK:
            app_mod.JOBS.pop(dummy_jid, None)

    # 3. Error sanitization: internal exception never leaks server filesystem paths
    err_msg = app_mod._sanitize_worker_error(
        OSError(r"Permission denied: C:\Users\SecretAdmin\uploads\private.pdf"),
        "nén tệp",
    )
    assert "SecretAdmin" not in err_msg
    assert "C:\\" not in err_msg
