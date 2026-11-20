import hashlib
import io
import os
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Optional

import pypdf
import pypdfium2 as pdfium
from PIL import Image

from .pdf_compressor import _notify_progress, _pil_to_base64_jpeg
from .pptx_compressor import (
    IMAGE_EXTENSIONS,
    _export_pptx_to_pdf_via_com,
    _iter_pptx_slides_pil,
    _validate_zip_safety,
)

Image.MAX_IMAGE_PIXELS = 64_000_000

STORED_MEDIA_EXTS = {".jpg", ".jpeg", ".png", ".webp"}


def _write_zip_bytes(zout: zipfile.ZipFile, arcname: str, data: bytes) -> None:
    ext = Path(arcname).suffix.lower()
    zinfo = zipfile.ZipInfo(filename=arcname)
    zinfo.compress_type = (
        zipfile.ZIP_STORED if ext in STORED_MEDIA_EXTS else zipfile.ZIP_DEFLATED
    )
    zout.writestr(zinfo, data)


def _render_pdf_pages_to_zip(
    pdf_path: str,
    zout: zipfile.ZipFile,
    folder: str,
    prefix: str,
    max_dim_px: int = 1920,
    jpeg_quality: int = 90,
    progress_callback: Optional[Callable] = None,
) -> tuple[int, Optional[str]]:
    doc = pdfium.PdfDocument(pdf_path)
    extracted_count = 0
    preview_b64: Optional[str] = None
    try:
        total_pages = len(doc)
        for idx in range(total_pages):
            page = doc[idx]
            w_pt, h_pt = page.get_size()
            max_pt = max(w_pt, h_pt, 1.0)
            scale = max(0.5, min(4.0, float(max_dim_px) / max_pt))

            bitmap = page.render(scale=scale)
            pil_im = bitmap.to_pil()
            try:
                rgb_im = pil_im.convert("RGB") if pil_im.mode != "RGB" else pil_im
                if preview_b64 is None:
                    preview_b64 = _pil_to_base64_jpeg(rgb_im, quality=82)

                buf = io.BytesIO()
                rgb_im.save(buf, format="JPEG", quality=int(jpeg_quality), optimize=True)
                arcname = f"{folder}/{prefix}_{idx + 1:03d}.jpg"
                _write_zip_bytes(zout, arcname, buf.getvalue())
                extracted_count += 1
            finally:
                try:
                    pil_im.close()
                except Exception:
                    pass
                try:
                    bitmap.close()
                except Exception:
                    pass

            _notify_progress(
                progress_callback,
                idx + 1,
                total_pages,
                f"Đang xuất ảnh {idx + 1}/{total_pages}...",
            )
    finally:
        doc.close()

    return extracted_count, preview_b64


def extract_archive(
    input_path: str,
    output_zip_path: str,
    file_type: str = "pdf",
    extract_mode: str = "images",
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Extract embedded media images ('images') or rendered HD page/slide JPEGs ('slides')
    from a .pdf, .pptx, or .ppt file into a .zip archive.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    out_zip = Path(output_zip_path)
    out_zip.parent.mkdir(parents=True, exist_ok=True)

    ext = in_path.suffix.lower()
    is_zip = zipfile.is_zipfile(str(in_path))

    if ext == ".pptx" and not is_zip:
        raise ValueError("Tệp .pptx bị hỏng hoặc không đúng cấu trúc ZIP.")
    if is_zip:
        with zipfile.ZipFile(str(in_path), "r") as zf:
            _validate_zip_safety(zf)

    ft = (file_type or "").lower().strip().lstrip(".")
    if ext == ".ppt" and not is_zip:
        ft = "ppt"
    elif ft not in ("pdf", "pptx", "ppt"):
        if ext == ".pdf":
            ft = "pdf"
        elif ext == ".ppt" and not is_zip:
            ft = "ppt"
        else:
            ft = "pptx"
    elif ft == "pptx" and not is_zip:
        ft = "ppt"

    mode = (extract_mode or "images").lower().strip()
    if mode not in ("images", "slides"):
        raise ValueError(f"Chế độ trích xuất không hợp lệ: {extract_mode}")

    if mode == "images" and ft == "ppt":
        raise ValueError(
            "Định dạng .ppt cũ không hỗ trợ trích xuất ảnh gốc trực tiếp. "
            "Vui lòng chọn chế độ tách slide thành ảnh hoặc lưu dưới dạng .pptx."
        )

    extracted_count = 0
    preview_b64: Optional[str] = None

    if mode == "images":
        if ft == "pptx":
            seen_hashes: set[str] = set()
            used_names: set[str] = set()
            with zipfile.ZipFile(str(in_path), "r") as zin:
                media_infos = [
                    info
                    for info in zin.infolist()
                    if info.filename.startswith("ppt/media/")
                    and not info.is_dir()
                    and Path(info.filename).suffix.lower() in IMAGE_EXTENSIONS
                ]
                total_media = max(len(media_infos), 1)
                with zipfile.ZipFile(
                    str(out_zip), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
                ) as zout:
                    for idx, info in enumerate(media_infos):
                        raw = zin.read(info.filename)
                        digest = hashlib.sha256(raw).hexdigest()
                        if digest not in seen_hashes:
                            seen_hashes.add(digest)
                            base_name = Path(info.filename).name
                            if base_name.lower() in used_names:
                                stem = Path(base_name).stem
                                sfx = Path(base_name).suffix.lower()
                                base_name = f"{stem}_{extracted_count + 1:03d}{sfx}"
                            used_names.add(base_name.lower())

                            _write_zip_bytes(zout, f"images/{base_name}", raw)
                            extracted_count += 1

                            if preview_b64 is None:
                                try:
                                    with Image.open(io.BytesIO(raw)) as im:
                                        im.load()
                                        preview_b64 = _pil_to_base64_jpeg(im, quality=82)
                                except Exception:
                                    pass

                        _notify_progress(
                            progress_callback,
                            idx + 1,
                            total_media,
                            f"Đang trích xuất ảnh gốc {idx + 1}/{len(media_infos)}...",
                        )

        elif ft == "pdf":
            seen_hashes = set()
            reader = pypdf.PdfReader(str(in_path))
            total_pages = max(len(reader.pages), 1)

            with zipfile.ZipFile(
                str(out_zip), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
            ) as zout:
                for page_idx, page in enumerate(reader.pages):
                    try:
                        img_keys = list(page.images.keys())
                    except Exception:
                        img_keys = []

                    for key in img_keys:
                        try:
                            img_file = page.images[key]
                            raw = getattr(img_file, "data", None)
                            img_ext = Path(getattr(img_file, "name", "") or "").suffix.lower()

                            if not raw or img_ext not in IMAGE_EXTENSIONS:
                                pil_im = img_file.image
                                if pil_im is None:
                                    continue
                                buf = io.BytesIO()
                                has_alpha = pil_im.mode in ("RGBA", "LA") or (
                                    pil_im.mode == "P" and "transparency" in pil_im.info
                                )
                                if has_alpha:
                                    pil_im.save(buf, format="PNG", optimize=True)
                                    img_ext = ".png"
                                else:
                                    rgb = pil_im.convert("RGB") if pil_im.mode != "RGB" else pil_im
                                    rgb.save(buf, format="JPEG", quality=92, optimize=True)
                                    img_ext = ".jpg"
                                raw = buf.getvalue()

                            digest = hashlib.sha256(raw).hexdigest()
                            if digest in seen_hashes:
                                continue
                            seen_hashes.add(digest)

                            extracted_count += 1
                            arcname = f"images/anh_{extracted_count:03d}{img_ext}"
                            _write_zip_bytes(zout, arcname, raw)

                            if preview_b64 is None:
                                try:
                                    with Image.open(io.BytesIO(raw)) as im:
                                        im.load()
                                        preview_b64 = _pil_to_base64_jpeg(im, quality=82)
                                except Exception:
                                    try:
                                        if img_file.image is not None:
                                            preview_b64 = _pil_to_base64_jpeg(
                                                img_file.image, quality=82
                                            )
                                    except Exception:
                                        pass
                        except Exception:
                            continue

                    _notify_progress(
                        progress_callback,
                        page_idx + 1,
                        total_pages,
                        f"Đang quét ảnh gốc trang {page_idx + 1}/{len(reader.pages)}...",
                    )

                if extracted_count == 0 and len(reader.pages) > 0:
                    extracted_count, preview_b64 = _render_pdf_pages_to_zip(
                        str(in_path),
                        zout,
                        folder="pages",
                        prefix="trang",
                        max_dim_px=1920,
                        jpeg_quality=90,
                        progress_callback=progress_callback,
                    )

    else:
        # mode == "slides"
        if ft == "pdf":
            with zipfile.ZipFile(
                str(out_zip), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
            ) as zout:
                extracted_count, preview_b64 = _render_pdf_pages_to_zip(
                    str(in_path),
                    zout,
                    folder="slides",
                    prefix="trang",
                    max_dim_px=1920,
                    jpeg_quality=90,
                    progress_callback=progress_callback,
                )

        elif ft == "pptx":
            with zipfile.ZipFile(
                str(out_zip), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
            ) as zout:
                for idx, total_slides, slide_img in _iter_pptx_slides_pil(
                    str(in_path), width=1920
                ):
                    try:
                        rgb_im = (
                            slide_img.convert("RGB")
                            if slide_img.mode != "RGB"
                            else slide_img
                        )
                        if preview_b64 is None:
                            preview_b64 = _pil_to_base64_jpeg(rgb_im, quality=82)

                        buf = io.BytesIO()
                        rgb_im.save(buf, format="JPEG", quality=90, optimize=True)
                        _write_zip_bytes(
                            zout, f"slides/slide_{idx + 1:03d}.jpg", buf.getvalue()
                        )
                        extracted_count += 1
                    finally:
                        try:
                            slide_img.close()
                        except Exception:
                            pass

                    _notify_progress(
                        progress_callback,
                        idx + 1,
                        max(total_slides, 1),
                        f"Đang tách slide {idx + 1}/{total_slides}...",
                    )

        elif ft == "ppt":
            fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="ppt_slides_")
            os.close(fd)
            try:
                exported_ok = _export_pptx_to_pdf_via_com(str(in_path), temp_pdf_path)
                if not exported_ok:
                    raise RuntimeError(
                        "Không thể tách slide từ tệp .ppt cũ do Microsoft PowerPoint không phản hồi. "
                        "Vui lòng lưu tệp sang .pptx hoặc .pdf rồi thử lại."
                    )
                with zipfile.ZipFile(
                    str(out_zip), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
                ) as zout:
                    extracted_count, preview_b64 = _render_pdf_pages_to_zip(
                        temp_pdf_path,
                        zout,
                        folder="slides",
                        prefix="slide",
                        max_dim_px=1920,
                        jpeg_quality=90,
                        progress_callback=progress_callback,
                    )
            finally:
                if os.path.exists(temp_pdf_path):
                    try:
                        os.remove(temp_pdf_path)
                    except Exception:
                        pass

    zip_size_bytes = out_zip.stat().st_size if out_zip.exists() else 0
    return {
        "output_path": str(out_zip),
        "extracted_count": int(extracted_count),
        "zip_size_bytes": int(zip_size_bytes),
        "zip_size_mb": round(zip_size_bytes / (1024 * 1024), 4),
        "extract_mode": mode,
        "preview_base64": preview_b64,
    }
