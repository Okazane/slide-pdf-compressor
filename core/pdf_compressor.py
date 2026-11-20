import base64
import io
import os
import shutil
from pathlib import Path
from typing import Callable, Optional

import pypdf
import pypdfium2 as pdfium
from PIL import Image
from pypdf.generic import (
    DecodedStreamObject,
    DictionaryObject,
    EncodedStreamObject,
    IndirectObject,
    NameObject,
    NumberObject,
    StreamObject,
)

Image.MAX_IMAGE_PIXELS = 64_000_000

# Buffer tuple: (jpeg_bytes, px_w, px_h, w_pt, h_pt)
PageJpegBuffer = tuple[bytes, int, int, float, float]


def _notify_progress(
    callback: Optional[Callable], current: int, total: int, message: str
) -> None:
    if callback is None:
        return
    try:
        callback(current, total, message)
    except TypeError:
        try:
            callback(current, total)
        except Exception:
            pass
    except Exception:
        pass


def _pil_to_base64_jpeg(img: Image.Image, quality: int = 82) -> str:
    work_img = img
    close_work = False
    try:
        if work_img.mode != "RGB":
            if work_img.mode in ("RGBA", "LA") or (
                work_img.mode == "P" and "transparency" in work_img.info
            ):
                rgba = work_img.convert("RGBA")
                bg = Image.new("RGB", rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.getchannel("A"))
                rgba.close()
                work_img = bg
            else:
                work_img = work_img.convert("RGB")
            close_work = True

        buf = io.BytesIO()
        work_img.save(buf, format="JPEG", quality=int(quality), optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{b64}"
    finally:
        if close_work and work_img is not img:
            work_img.close()


def _render_page_preview_base64(
    pdf_path: str, page_index: int = 0, scale: float = 1.0, quality: int = 82
) -> str:
    doc = pdfium.PdfDocument(pdf_path)
    page = None
    bitmap = None
    pil_img = None
    try:
        if len(doc) == 0:
            return "data:image/jpeg;base64,"
        idx = max(0, min(page_index, len(doc) - 1))
        page = doc[idx]
        w_pt, h_pt = page.get_size()
        max_pt = max(w_pt, h_pt, 1.0)
        eff_scale = min(float(scale), 1440.0 / max_pt)
        bitmap = page.render(scale=eff_scale)
        pil_img = bitmap.to_pil()
        return _pil_to_base64_jpeg(pil_img, quality=quality)
    finally:
        if pil_img is not None:
            pil_img.close()
        if bitmap is not None and hasattr(bitmap, "close"):
            bitmap.close()
        if page is not None and hasattr(page, "close"):
            page.close()
        doc.close()


def _count_unique_pdf_images(pdf_path: str) -> int:
    """
    Count unique embedded XObject images across all pages via pypdf.PdfReader,
    deduplicated by indirect reference idnum without decoding image pixels.
    """
    try:
        reader = pypdf.PdfReader(pdf_path)
        seen_image_ids: set[int] = set()
        seen_form_ids: set[int] = set()
        direct_images = 0

        def _walk_resources(container: DictionaryObject) -> None:
            nonlocal direct_images
            res_raw = container.get("/Resources")
            if res_raw is None:
                return
            res = res_raw.get_object() if hasattr(res_raw, "get_object") else res_raw
            if not isinstance(res, DictionaryObject) or "/XObject" not in res:
                return
            xobj_raw = res["/XObject"]
            xobj = (
                xobj_raw.get_object() if hasattr(xobj_raw, "get_object") else xobj_raw
            )
            if not isinstance(xobj, DictionaryObject):
                return

            for key in xobj.keys():
                raw_entry = xobj.raw_get(key)
                idnum = (
                    raw_entry.idnum if isinstance(raw_entry, IndirectObject) else None
                )
                if idnum is not None and (
                    idnum in seen_image_ids or idnum in seen_form_ids
                ):
                    continue

                try:
                    target = (
                        raw_entry.get_object()
                        if hasattr(raw_entry, "get_object")
                        else raw_entry
                    )
                except Exception:
                    continue

                if not isinstance(target, StreamObject):
                    continue

                subtype = target.get("/Subtype", "")
                if subtype == "/Image":
                    if idnum is not None:
                        seen_image_ids.add(idnum)
                    else:
                        direct_images += 1
                elif subtype == "/Form":
                    if idnum is not None:
                        seen_form_ids.add(idnum)
                    _walk_resources(target)

        for page in reader.pages:
            _walk_resources(page)

        return len(seen_image_ids) + direct_images
    except Exception:
        return 0


def inspect_pdf(pdf_path: str) -> dict:
    """
    Inspect a PDF file and return metadata (including unique embedded XObject media_count)
    plus a base64 JPEG thumbnail of page 0.
    """
    path_obj = Path(pdf_path)
    if not path_obj.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    original_size_bytes = path_obj.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)

    doc = pdfium.PdfDocument(str(path_obj))
    first_page = None
    bitmap = None
    pil_img = None
    try:
        page_count = len(doc)
        if page_count > 0:
            first_page = doc[0]
            width, height = first_page.get_size()
            max_pt = max(width, height, 1.0)
            thumb_scale = min(1.0, 1440.0 / max_pt)
            bitmap = first_page.render(scale=thumb_scale)
            pil_img = bitmap.to_pil()
            preview_base64 = _pil_to_base64_jpeg(pil_img, quality=80)
        else:
            width, height = 0.0, 0.0
            preview_base64 = "data:image/jpeg;base64,"
    finally:
        if pil_img is not None:
            pil_img.close()
        if bitmap is not None and hasattr(bitmap, "close"):
            bitmap.close()
        if first_page is not None and hasattr(first_page, "close"):
            first_page.close()
        doc.close()

    media_count = _count_unique_pdf_images(str(path_obj))

    return {
        "page_count": int(page_count),
        "original_size_bytes": int(original_size_bytes),
        "original_size_mb": float(original_size_mb),
        "width": float(round(width, 2)),
        "height": float(round(height, 2)),
        "media_count": int(media_count),
        "preview_base64": preview_base64,
    }


def _estimate_jpeg_size(pil_img: Image.Image, quality: int) -> int:
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=int(quality), optimize=True)
    return buf.tell()


def _encode_pil_to_jpeg_bytes(pil_img: Image.Image, quality: int) -> bytes:
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=int(quality), optimize=True)
    return buf.getvalue()


def _page_base_factor(w_pt: float, h_pt: float) -> float:
    """
    Normalize rendering base scale so standard slides (720pt-960pt) render sharply (~1440px)
    at scale=1.0 while oversized architectural/poster pages stay bounded in pixel count.
    """
    max_dim_pt = max(float(w_pt), float(h_pt), 1.0)
    if max_dim_pt < 1000.0:
        return min(1.5, 1440.0 / max_dim_pt)
    if max_dim_pt > 2000.0:
        return 1600.0 / max_dim_pt
    return 1.0


def _save_jpeg_buffers_as_pdf(
    page_buffers: list[PageJpegBuffer], output_path: str
) -> int:
    """
    Embed pre-encoded JPEG buffers directly as /DCTDecode Image XObjects onto blank PDF
    pages of exact physical dimensions (w_pt, h_pt) using pypdf.PdfWriter.
    Preserves mixed per-page dimensions accurately regardless of pixel render scale.
    """
    if not page_buffers:
        raise ValueError("Cannot save empty page buffer list as PDF.")

    writer = pypdf.PdfWriter()
    for jpeg_bytes, px_w, px_h, w_pt, h_pt in page_buffers:
        page = writer.add_blank_page(width=float(w_pt), height=float(h_pt))

        img_stream = EncodedStreamObject()
        img_stream._data = jpeg_bytes
        img_stream.update(
            {
                NameObject("/Type"): NameObject("/XObject"),
                NameObject("/Subtype"): NameObject("/Image"),
                NameObject("/Width"): NumberObject(int(px_w)),
                NameObject("/Height"): NumberObject(int(px_h)),
                NameObject("/ColorSpace"): NameObject("/DeviceRGB"),
                NameObject("/BitsPerComponent"): NumberObject(8),
                NameObject("/Filter"): NameObject("/DCTDecode"),
            }
        )
        img_ref = writer._add_object(img_stream)

        page[NameObject("/Resources")] = DictionaryObject(
            {
                NameObject("/XObject"): DictionaryObject(
                    {NameObject("/Im0"): img_ref}
                )
            }
        )

        content_cmd = f"q {w_pt:.4f} 0 0 {h_pt:.4f} 0 0 cm /Im0 Do Q".encode("ascii")
        content_stream = DecodedStreamObject()
        content_stream.set_data(content_cmd)
        page[NameObject("/Contents")] = writer._add_object(content_stream)

    with open(output_path, "wb") as f:
        writer.write(f)

    return os.path.getsize(output_path)


def _reencode_jpeg_buffers(
    source_buffers: list[PageJpegBuffer], scale_ratio: float, quality: int
) -> list[PageJpegBuffer]:
    """
    Re-encode JPEG page buffers one page at a time to lower quality and/or pixel dimensions
    without keeping more than a single decoded PIL image in RAM.
    """
    reencoded: list[PageJpegBuffer] = []
    for jpeg_bytes, orig_px_w, orig_px_h, w_pt, h_pt in source_buffers:
        im = Image.open(io.BytesIO(jpeg_bytes))
        try:
            im.load()
            work_im = im.convert("RGB") if im.mode != "RGB" else im
            try:
                if scale_ratio < 0.999:
                    new_w = max(1, int(round(orig_px_w * scale_ratio)))
                    new_h = max(1, int(round(orig_px_h * scale_ratio)))
                    resized = work_im.resize((new_w, new_h), Image.Resampling.BILINEAR)
                    if work_im is not im:
                        work_im.close()
                    work_im = resized

                new_px_w, new_px_h = work_im.size
                new_jpeg = _encode_pil_to_jpeg_bytes(work_im, quality=int(quality))
                reencoded.append((new_jpeg, new_px_w, new_px_h, w_pt, h_pt))
            finally:
                if work_im is not im:
                    work_im.close()
        finally:
            im.close()
    return reencoded


def _save_pil_pages_as_pdf(
    pages: list[Image.Image], output_path: str, quality: int, resolution: float = 72.0
) -> int:
    """Backward-compatible helper converting PIL images to PDF via /DCTDecode buffers."""
    if not pages:
        raise ValueError("Cannot save empty page list as PDF.")
    scale_pt = 72.0 / float(resolution if resolution > 0 else 72.0)
    buffers: list[PageJpegBuffer] = []
    for p in pages:
        rgb = p.convert("RGB") if p.mode != "RGB" else p
        try:
            px_w, px_h = rgb.size
            jb = _encode_pil_to_jpeg_bytes(rgb, quality=int(quality))
            buffers.append((jb, px_w, px_h, px_w * scale_pt, px_h * scale_pt))
        finally:
            if rgb is not p:
                rgb.close()
    return _save_jpeg_buffers_as_pdf(buffers, output_path)


def _compress_pdf_flatten(
    input_path: str,
    output_path: str,
    target_mb: float,
    progress_callback: Optional[Callable] = None,
) -> tuple[int, float, int]:
    """
    Flatten mode: renders PDF pages one at a time via pypdfium2 + Pillow into low-RAM
    JPEG byte buffers while preserving exact per-page physical dimensions (w_pt, h_pt).
    Samples up to 4 pages to select optimal (scale, jpeg_quality) fitting <= target_mb * 0.90,
    and iteratively lowers quality/scale buffer-by-buffer if final size exceeds target_mb.
    Returns (page_count, scale_used, quality_used).
    """
    target_bytes = int(target_mb * 1024 * 1024)
    budget_bytes = int(target_bytes * 0.90)

    candidates = [
        (1.0, 85),
        (1.0, 78),
        (1.0, 72),
        (1.0, 65),
        (0.9, 75),
        (0.85, 70),
        (0.8, 65),
        (0.75, 62),
        (0.7, 58),
        (0.65, 55),
        (0.6, 50),
        (0.55, 45),
        (0.5, 40),
    ]

    initial_buffers: list[PageJpegBuffer] = []
    doc = pdfium.PdfDocument(input_path)
    try:
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("Input PDF has 0 pages.")

        # Sample up to 4 pages: 0, n//3, 2*n//3, n-1
        raw_indices = [0, total_pages // 3, (2 * total_pages) // 3, total_pages - 1]
        sample_indices = sorted(set(i for i in raw_indices if 0 <= i < total_pages))

        sample_base_imgs: list[Image.Image] = []
        try:
            for s_idx in sample_indices:
                s_page = doc[s_idx]
                s_bmp = None
                s_pil = None
                try:
                    sw_pt, sh_pt = s_page.get_size()
                    s_base = _page_base_factor(sw_pt, sh_pt)
                    s_bmp = s_page.render(scale=1.0 * s_base)
                    s_pil = s_bmp.to_pil()
                    s_rgb = s_pil.convert("RGB") if s_pil.mode != "RGB" else s_pil
                    if s_rgb is s_pil:
                        s_pil = None
                    sample_base_imgs.append(s_rgb)
                finally:
                    if s_pil is not None:
                        s_pil.close()
                    if s_bmp is not None and hasattr(s_bmp, "close"):
                        s_bmp.close()
                    if hasattr(s_page, "close"):
                        s_page.close()

            chosen_scale, chosen_quality = 0.5, 40
            for cand_scale, cand_quality in candidates:
                total_sample_bytes = 0
                for base_im in sample_base_imgs:
                    if abs(cand_scale - 1.0) < 1e-3:
                        total_sample_bytes += _estimate_jpeg_size(base_im, cand_quality)
                    else:
                        new_w = max(1, int(round(base_im.width * cand_scale)))
                        new_h = max(1, int(round(base_im.height * cand_scale)))
                        scaled_im = base_im.resize(
                            (new_w, new_h), Image.Resampling.BILINEAR
                        )
                        try:
                            total_sample_bytes += _estimate_jpeg_size(
                                scaled_im, cand_quality
                            )
                        finally:
                            scaled_im.close()

                avg_page_bytes = total_sample_bytes / len(sample_base_imgs)
                # Include ~1.5 KB per page PDF container/xref overhead
                estimated_total = int((avg_page_bytes + 1536) * total_pages)
                if estimated_total <= budget_bytes:
                    chosen_scale, chosen_quality = cand_scale, cand_quality
                    break
        finally:
            for im in sample_base_imgs:
                im.close()
            sample_base_imgs.clear()

        # Render pages one at a time, immediately JPEG-encoding to bytes and closing bitmaps
        for idx in range(total_pages):
            page = doc[idx]
            bmp = None
            pil_raw = None
            pil_rgb = None
            try:
                w_pt, h_pt = page.get_size()
                render_scale = chosen_scale * _page_base_factor(w_pt, h_pt)
                bmp = page.render(scale=render_scale)
                pil_raw = bmp.to_pil()
                pil_rgb = pil_raw.convert("RGB") if pil_raw.mode != "RGB" else pil_raw
                px_w, px_h = pil_rgb.size
                jpeg_bytes = _encode_pil_to_jpeg_bytes(pil_rgb, quality=chosen_quality)
                initial_buffers.append(
                    (jpeg_bytes, px_w, px_h, float(w_pt), float(h_pt))
                )
            finally:
                if pil_rgb is not None and pil_rgb is not pil_raw:
                    pil_rgb.close()
                if pil_raw is not None:
                    pil_raw.close()
                if bmp is not None and hasattr(bmp, "close"):
                    bmp.close()
                if hasattr(page, "close"):
                    page.close()

            _notify_progress(
                progress_callback,
                idx + 1,
                total_pages,
                f"Đang xử lý trang {idx + 1}/{total_pages}...",
            )
    finally:
        doc.close()

    scale_used = chosen_scale
    quality_used = chosen_quality

    # Write initial PDF directly from JPEG byte buffers
    current_buffers = initial_buffers
    compressed_bytes = _save_jpeg_buffers_as_pdf(current_buffers, output_path)

    # Iteratively lower JPEG quality / scale buffer-by-buffer if size still exceeds target
    strict_limit = max(1024, int(target_bytes * 0.985))
    while compressed_bytes >= strict_limit:
        if quality_used > 40:
            quality_used = max(40, quality_used - 8)
        elif scale_used > 0.5:
            scale_used = max(0.5, round(scale_used - 0.1, 2))
            quality_used = 45
        else:
            # At standard floor (scale=0.5, quality=40). If still > target_bytes, perform
            # controlled emergency step(s) down to safe floor (scale >= 0.30, quality >= 25).
            if compressed_bytes > target_bytes and scale_used > 0.30:
                scale_used = max(0.30, round(scale_used * 0.75, 2))
                quality_used = max(25, quality_used - 10)
                ratio = scale_used / chosen_scale
                current_buffers = _reencode_jpeg_buffers(
                    initial_buffers, ratio, quality_used
                )
                compressed_bytes = _save_jpeg_buffers_as_pdf(
                    current_buffers, output_path
                )
                continue
            break

        ratio = scale_used / chosen_scale
        current_buffers = _reencode_jpeg_buffers(initial_buffers, ratio, quality_used)
        compressed_bytes = _save_jpeg_buffers_as_pdf(current_buffers, output_path)

    return total_pages, float(scale_used), int(quality_used)


def _get_original_embedded_image_size(img_file) -> int:
    """
    Compute the actual byte size occupied by an embedded image stream (plus any /SMask)
    inside the PDF, falling back to len(img_file.data) when raw stream data is unavailable.
    """
    raw_stream_len = 0
    try:
        if img_file.indirect_reference is not None:
            raw_obj = img_file.indirect_reference.get_object()
            if isinstance(raw_obj, StreamObject):
                raw_stream_len = len(getattr(raw_obj, "_data", b"") or b"")
                if "/SMask" in raw_obj:
                    try:
                        smask_obj = raw_obj["/SMask"].get_object()
                        if isinstance(smask_obj, StreamObject):
                            raw_stream_len += len(
                                getattr(smask_obj, "_data", b"") or b""
                            )
                    except Exception:
                        pass
    except Exception:
        raw_stream_len = 0

    if raw_stream_len > 0:
        return raw_stream_len

    try:
        data_bytes = getattr(img_file, "data", b"") or b""
        return len(data_bytes)
    except Exception:
        return 0


def _optimize_pdf_images_in_writer(
    writer: pypdf.PdfWriter,
    max_dim: int,
    jpeg_quality: int,
    force_rgb_for_rgba: bool = False,
    progress_callback: Optional[Callable] = None,
    pass_num: int = 1,
    total_passes: int = 1,
) -> None:
    total_pages = len(writer.pages)
    processed_obj_ids: set[int] = set()

    for idx, page in enumerate(writer.pages):
        try:
            image_keys = list(page.images.keys())
        except Exception:
            image_keys = []

        for key in image_keys:
            pil_im = None
            work_im = None
            try:
                img_file = page.images[key]
                if img_file.indirect_reference is not None:
                    obj_id = img_file.indirect_reference.idnum
                    if obj_id in processed_obj_ids:
                        continue
                    processed_obj_ids.add(obj_id)

                orig_size = _get_original_embedded_image_size(img_file)
                pil_im = img_file.image
                if pil_im is None:
                    continue

                w, h = pil_im.size
                scale = 1.0
                if max(w, h) > max_dim:
                    scale = max_dim / float(max(w, h))

                has_alpha = pil_im.mode in ("RGBA", "LA") or (
                    pil_im.mode == "P" and "transparency" in pil_im.info
                )

                if has_alpha:
                    rgba = pil_im.convert("RGBA")
                    min_alpha, _ = rgba.getchannel("A").getextrema()
                    # Convert opaque or large slide photos, or when forced on aggressive passes
                    if min_alpha == 255 or force_rgb_for_rgba or (w * h > 200_000):
                        bg = Image.new("RGB", rgba.size, (255, 255, 255))
                        bg.paste(rgba, mask=rgba.getchannel("A"))
                        rgba.close()
                        work_im = bg
                    else:
                        work_im = rgba
                else:
                    work_im = (
                        pil_im
                        if pil_im.mode in ("RGB", "L")
                        else pil_im.convert("RGB")
                    )

                was_downscaled = scale < 1.0
                if was_downscaled:
                    new_w = max(1, int(round(w * scale)))
                    new_h = max(1, int(round(h * scale)))
                    resized = work_im.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    if work_im is not pil_im:
                        work_im.close()
                    work_im = resized

                # Encode candidate into in-memory buffer to check byte length before replacing
                cand_buf = io.BytesIO()
                if work_im.mode == "RGBA":
                    work_im.save(cand_buf, format="PNG", optimize=True, compress_level=9)
                else:
                    work_im.save(
                        cand_buf,
                        format="JPEG",
                        quality=int(jpeg_quality),
                        optimize=True,
                    )
                cand_len = cand_buf.tell()

                # Only call img_file.replace(...) when smaller than original or downscaled
                if was_downscaled or (orig_size > 0 and cand_len < orig_size):
                    old_obj = (
                        img_file.indirect_reference.get_object()
                        if img_file.indirect_reference is not None
                        else None
                    )

                    if work_im.mode == "RGBA":
                        img_file.replace(work_im, optimize=True)
                    else:
                        img_file.replace(
                            work_im, quality=int(jpeg_quality), optimize=True
                        )

                    # Guard against PDF stream bloat (e.g. JPXDecode RGBA larger than original Flate)
                    if (
                        old_obj is not None
                        and orig_size > 0
                        and img_file.indirect_reference is not None
                    ):
                        new_obj = img_file.indirect_reference.get_object()
                        if isinstance(new_obj, StreamObject):
                            new_stream_len = len(getattr(new_obj, "_data", b"") or b"")
                            if new_stream_len >= orig_size:
                                writer._objects[
                                    img_file.indirect_reference.idnum - 1
                                ] = old_obj
            except Exception:
                continue
            finally:
                if work_im is not None and work_im is not pil_im:
                    work_im.close()
                if pil_im is not None:
                    pil_im.close()

        try:
            page.compress_content_streams()
        except Exception:
            pass

        if total_passes > 1 and pass_num > 1:
            msg = f"Đang tối ưu vector (lượt {pass_num}/{total_passes}) trang {idx + 1}/{total_pages}..."
        else:
            msg = f"Đang tối ưu vector trang {idx + 1}/{total_pages}..."

        _notify_progress(
            progress_callback,
            idx + 1,
            total_pages,
            msg,
        )

    try:
        writer.compress_identical_objects()
    except Exception:
        pass


def _compress_pdf_vector(
    input_path: str,
    output_path: str,
    target_mb: float,
    original_size_bytes: int,
    progress_callback: Optional[Callable] = None,
) -> tuple[int, float, int]:
    """
    Vector mode: preserves vector text/fonts via pypdf.PdfReader & PdfWriter while
    selectively downscaling and re-encoding embedded images when beneficial.
    Runs a single gentle pass if original_size_bytes <= target_bytes, or adaptive
    multi-pass compression when original_size_bytes > target_bytes.
    """
    target_bytes = int(target_mb * 1024 * 1024)

    if original_size_bytes <= target_bytes:
        passes = [
            (1440, 80, False, 1.0),
        ]
    else:
        ratio = (
            target_bytes / float(original_size_bytes)
            if original_size_bytes > 0
            else 0.5
        )
        if ratio >= 0.75:
            initial_max_dim = 1280
            initial_quality = 75
            scale_used = 0.85
        elif ratio >= 0.4:
            initial_max_dim = 1024
            initial_quality = 68
            scale_used = 0.75
        else:
            initial_max_dim = 800
            initial_quality = 58
            scale_used = 0.6

        passes = [
            (initial_max_dim, initial_quality, False, scale_used),
            (720, 50, True, 0.5),
            (480, 40, True, 0.4),
        ]

    page_count = 0
    best_size: Optional[int] = None
    best_scale = passes[0][3]
    best_quality = passes[0][1]
    total_passes = len(passes)
    tmp_cand_path = f"{output_path}.vec_tmp"

    try:
        for pass_idx, (max_dim, quality, force_rgb, s_used) in enumerate(passes):
            reader = pypdf.PdfReader(input_path)
            page_count = len(reader.pages)
            writer = pypdf.PdfWriter(clone_from=reader)

            _optimize_pdf_images_in_writer(
                writer,
                max_dim=max_dim,
                jpeg_quality=quality,
                force_rgb_for_rgba=force_rgb,
                progress_callback=progress_callback,
                pass_num=pass_idx + 1,
                total_passes=total_passes,
            )

            write_target = output_path if pass_idx == 0 else tmp_cand_path
            with open(write_target, "wb") as f:
                writer.write(f)

            cand_size = os.path.getsize(write_target)
            if best_size is None or cand_size <= best_size:
                if write_target != output_path:
                    os.replace(write_target, output_path)
                best_size = cand_size
                best_scale = s_used
                best_quality = quality

            if best_size <= target_bytes and best_size <= original_size_bytes:
                break
    finally:
        if os.path.exists(tmp_cand_path):
            try:
                os.remove(tmp_cand_path)
            except Exception:
                pass

    return page_count, float(best_scale), int(best_quality)


def compress_pdf(
    input_path: str,
    output_path: str,
    target_mb: float = 10.0,
    mode: str = "flatten",
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Compress a PDF file using either 'flatten' mode (pypdfium2 + Pillow + /DCTDecode PdfWriter)
    or 'vector' mode (pypdf selective image re-encoding + stream compression).
    Includes non-inflation guard when input is already within target_mb.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    original_size_bytes = in_path.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)
    target_bytes = int(float(target_mb) * 1024 * 1024)

    before_preview_base64 = _render_page_preview_base64(
        str(in_path), page_index=0, scale=1.0, quality=85
    )

    mode_clean = (mode or "flatten").lower().strip()
    if mode_clean not in ("flatten", "vector"):
        mode_clean = "flatten"

    if mode_clean == "vector":
        page_count, scale_used, quality_used = _compress_pdf_vector(
            str(in_path),
            str(out_path),
            target_mb=float(target_mb),
            original_size_bytes=original_size_bytes,
            progress_callback=progress_callback,
        )
    else:
        page_count, scale_used, quality_used = _compress_pdf_flatten(
            str(in_path),
            str(out_path),
            target_mb=float(target_mb),
            progress_callback=progress_callback,
        )

    compressed_size_bytes = out_path.stat().st_size

    # Non-inflation guard: never increase file size when input already fits within target_mb
    kept_original = False
    if compressed_size_bytes >= original_size_bytes and original_size_bytes <= target_bytes:
        if in_path.resolve() != out_path.resolve():
            shutil.copy2(str(in_path), str(out_path))
        compressed_size_bytes = out_path.stat().st_size
        scale_used = 1.0
        kept_original = True

    compressed_size_mb = round(compressed_size_bytes / (1024 * 1024), 4)

    if original_size_bytes > 0:
        reduction_percent = round(
            (1.0 - (compressed_size_bytes / float(original_size_bytes))) * 100.0, 2
        )
        if reduction_percent == 0.0:
            reduction_percent = 0.0
    else:
        reduction_percent = 0.0

    if kept_original:
        after_preview_base64 = before_preview_base64
    else:
        after_preview_base64 = _render_page_preview_base64(
            str(out_path), page_index=0, scale=1.0, quality=85
        )

    target_met = bool(compressed_size_bytes <= target_bytes)

    return {
        "output_path": str(out_path),
        "original_size_bytes": int(original_size_bytes),
        "compressed_size_bytes": int(compressed_size_bytes),
        "original_size_mb": float(original_size_mb),
        "compressed_size_mb": float(compressed_size_mb),
        "reduction_percent": float(reduction_percent),
        "page_count": int(page_count),
        "mode_used": mode_clean,
        "scale_used": float(scale_used),
        "quality_used": int(quality_used),
        "before_preview_base64": before_preview_base64,
        "after_preview_base64": after_preview_base64,
        "target_met": target_met,
    }
