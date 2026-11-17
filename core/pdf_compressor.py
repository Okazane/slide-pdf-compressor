import base64
import io
import math
import os
from pathlib import Path
from typing import Callable, Optional

import pypdf
import pypdfium2 as pdfium
from PIL import Image


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
    if img.mode != "RGB":
        if img.mode in ("RGBA", "LA") or (
            img.mode == "P" and "transparency" in img.info
        ):
            rgba = img.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.getchannel("A"))
            img = bg
        else:
            img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{b64}"


def _render_page_preview_base64(
    pdf_path: str, page_index: int = 0, scale: float = 1.0, quality: int = 82
) -> str:
    doc = pdfium.PdfDocument(pdf_path)
    try:
        if len(doc) == 0:
            return "data:image/jpeg;base64,"
        idx = max(0, min(page_index, len(doc) - 1))
        page = doc[idx]
        bitmap = page.render(scale=scale)
        pil_img = bitmap.to_pil()
        return _pil_to_base64_jpeg(pil_img, quality=quality)
    finally:
        doc.close()


def inspect_pdf(pdf_path: str) -> dict:
    """
    Inspect a PDF file and return metadata plus a base64 JPEG thumbnail of page 0.
    """
    path_obj = Path(pdf_path)
    if not path_obj.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    original_size_bytes = path_obj.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)

    doc = pdfium.PdfDocument(str(path_obj))
    try:
        page_count = len(doc)
        if page_count > 0:
            first_page = doc[0]
            width, height = first_page.get_size()
            # Render thumbnail at scale 1.0 (or capped if page is huge)
            max_pt = max(width, height, 1.0)
            thumb_scale = min(1.0, 1440.0 / max_pt)
            bitmap = first_page.render(scale=thumb_scale)
            pil_img = bitmap.to_pil()
            preview_base64 = _pil_to_base64_jpeg(pil_img, quality=80)
        else:
            width, height = 0.0, 0.0
            preview_base64 = "data:image/jpeg;base64,"
    finally:
        doc.close()

    return {
        "page_count": int(page_count),
        "original_size_bytes": int(original_size_bytes),
        "original_size_mb": float(original_size_mb),
        "width": float(round(width, 2)),
        "height": float(round(height, 2)),
        "preview_base64": preview_base64,
    }


def _estimate_jpeg_size(pil_img: Image.Image, quality: int) -> int:
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=quality, optimize=True)
    return len(buf.getvalue())


def _save_pil_pages_as_pdf(
    pages: list[Image.Image], output_path: str, quality: int, resolution: float = 72.0
) -> int:
    if not pages:
        raise ValueError("Cannot save empty page list as PDF.")
    rgb_pages = [p.convert("RGB") if p.mode != "RGB" else p for p in pages]
    rgb_pages[0].save(
        output_path,
        format="PDF",
        resolution=resolution,
        save_all=True,
        append_images=rgb_pages[1:],
        quality=int(quality),
        optimize=True,
    )
    return os.path.getsize(output_path)


def _compress_pdf_flatten(
    input_path: str,
    output_path: str,
    target_mb: float,
    progress_callback: Optional[Callable] = None,
) -> tuple[int, float, int]:
    """
    Flatten mode: renders PDF pages using pypdfium2 + Pillow, samples up to 4 pages
    to pick the highest quality (scale, jpeg_quality) fitting <= target_mb * 0.90,
    and iteratively lowers quality/scale if final size exceeds target_mb.
    Returns (page_count, scale_used, quality_used).
    """
    target_bytes = int(target_mb * 1024 * 1024)
    budget_bytes = int(target_bytes * 0.90)

    doc = pdfium.PdfDocument(input_path)
    try:
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("Input PDF has 0 pages.")

        p0_w, p0_h = doc[0].get_size()
        max_dim_pt = max(p0_w, p0_h, 1.0)
        # If native page points are small (e.g. 720pt or 921pt from 150dpi save),
        # base_scale_factor allows scale=1.0 to correspond to sharp ~1440px slide width
        base_factor = 1.0
        if max_dim_pt < 1000.0:
            base_factor = min(1.5, 1440.0 / max_dim_pt)
        elif max_dim_pt > 2000.0:
            base_factor = 1600.0 / max_dim_pt

        # Sample up to 4 pages: 0, n//3, 2*n//3, n-1
        raw_indices = [0, total_pages // 3, (2 * total_pages) // 3, total_pages - 1]
        sample_indices = sorted(set(i for i in raw_indices if 0 <= i < total_pages))

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

        # Pre-render sample pages at scale 1.0 * base_factor once, then resize for faster estimation
        sample_base_imgs = []
        for s_idx in sample_indices:
            bmp = doc[s_idx].render(scale=1.0 * base_factor)
            pil_im = bmp.to_pil().convert("RGB")
            sample_base_imgs.append(pil_im)

        chosen_scale, chosen_quality = 0.5, 40
        for cand_scale, cand_quality in candidates:
            total_sample_bytes = 0
            for base_im in sample_base_imgs:
                if abs(cand_scale - 1.0) < 1e-3:
                    scaled_im = base_im
                else:
                    new_w = max(1, int(base_im.width * cand_scale))
                    new_h = max(1, int(base_im.height * cand_scale))
                    scaled_im = base_im.resize((new_w, new_h), Image.Resampling.BILINEAR)
                total_sample_bytes += _estimate_jpeg_size(scaled_im, cand_quality)

            avg_page_bytes = total_sample_bytes / len(sample_base_imgs)
            # Include ~2 KB per page PDF container overhead
            estimated_total = int((avg_page_bytes + 2048) * total_pages)
            if estimated_total <= budget_bytes:
                chosen_scale, chosen_quality = cand_scale, cand_quality
                break

        # Render all pages at chosen_scale * base_factor
        rendered_pages: list[Image.Image] = []
        render_scale = chosen_scale * base_factor
        for idx in range(total_pages):
            bmp = doc[idx].render(scale=render_scale)
            pil_im = bmp.to_pil().convert("RGB")
            rendered_pages.append(pil_im)
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

    # Save initial PDF
    current_pages = rendered_pages
    compressed_bytes = _save_pil_pages_as_pdf(current_pages, output_path, quality_used)

    # Iteratively lower JPEG quality / scale from in-memory rendered images if size still exceeds target
    # Aim strictly under target_bytes (with small margin so rounded MB is also < target_mb)
    strict_limit = max(1024, int(target_bytes * 0.985))
    while compressed_bytes >= strict_limit:
        if quality_used > 40:
            quality_used = max(40, quality_used - 8)
        elif scale_used > 0.5:
            scale_used = max(0.5, round(scale_used - 0.1, 2))
            quality_used = 45
            ratio = scale_used / chosen_scale
            current_pages = [
                p.resize(
                    (max(1, int(p.width * ratio)), max(1, int(p.height * ratio))),
                    Image.Resampling.BILINEAR,
                )
                for p in rendered_pages
            ]
        else:
            # Already at floor (scale=0.5, quality=40). If still >= target_bytes, perform emergency step
            if compressed_bytes >= target_bytes and scale_used > 0.25:
                scale_used = round(scale_used * 0.75, 2)
                quality_used = max(25, quality_used - 10)
                ratio = scale_used / chosen_scale
                current_pages = [
                    p.resize(
                        (max(1, int(p.width * ratio)), max(1, int(p.height * ratio))),
                        Image.Resampling.BILINEAR,
                    )
                    for p in rendered_pages
                ]
                compressed_bytes = _save_pil_pages_as_pdf(
                    current_pages, output_path, quality_used
                )
            break

        compressed_bytes = _save_pil_pages_as_pdf(
            current_pages, output_path, quality_used
        )

    return total_pages, float(scale_used), int(quality_used)


def _optimize_pdf_images_in_writer(
    writer: pypdf.PdfWriter,
    max_dim: int,
    jpeg_quality: int,
    force_rgb_for_rgba: bool = False,
    progress_callback: Optional[Callable] = None,
) -> None:
    total_pages = len(writer.pages)
    processed_obj_ids: set[int] = set()

    for idx, page in enumerate(writer.pages):
        try:
            image_keys = list(page.images.keys())
        except Exception:
            image_keys = []

        for key in image_keys:
            try:
                img_file = page.images[key]
                if img_file.indirect_reference is not None:
                    obj_id = img_file.indirect_reference.idnum
                    if obj_id in processed_obj_ids:
                        continue
                    processed_obj_ids.add(obj_id)

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
                    # Convert opaque or semi-opaque large slide backgrounds/photos or when forced
                    if min_alpha == 255 or force_rgb_for_rgba or (w * h > 200_000):
                        bg = Image.new("RGB", rgba.size, (255, 255, 255))
                        bg.paste(rgba, mask=rgba.getchannel("A"))
                        work_im = bg
                    else:
                        work_im = rgba
                        if scale == 1.0 and max(w, h) > 300:
                            scale = 0.75
                else:
                    work_im = pil_im.convert("RGB") if pil_im.mode != "RGB" else pil_im

                if scale < 1.0:
                    new_w = max(1, int(w * scale))
                    new_h = max(1, int(h * scale))
                    work_im = work_im.resize(
                        (new_w, new_h), Image.Resampling.LANCZOS
                    )

                img_file.replace(work_im, quality=jpeg_quality, optimize=True)
            except Exception:
                continue

        try:
            page.compress_content_streams()
        except Exception:
            pass

        _notify_progress(
            progress_callback,
            idx + 1,
            total_pages,
            f"Đang tối ưu vector trang {idx + 1}/{total_pages}...",
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
    downscaling and re-encoding embedded images to meet target_mb.
    """
    target_bytes = int(target_mb * 1024 * 1024)
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
    final_scale = scale_used
    final_quality = initial_quality

    for pass_idx, (max_dim, quality, force_rgb, s_used) in enumerate(passes):
        reader = pypdf.PdfReader(input_path)
        page_count = len(reader.pages)
        writer = pypdf.PdfWriter(clone_from=reader)

        _optimize_pdf_images_in_writer(
            writer,
            max_dim=max_dim,
            jpeg_quality=quality,
            force_rgb_for_rgba=force_rgb,
            progress_callback=progress_callback if pass_idx == 0 else None,
        )

        with open(output_path, "wb") as f:
            writer.write(f)

        final_scale = s_used
        final_quality = quality
        out_size = os.path.getsize(output_path)

        if out_size < target_bytes and out_size < original_size_bytes:
            break

    return page_count, float(final_scale), int(final_quality)


def compress_pdf(
    input_path: str,
    output_path: str,
    target_mb: float = 10.0,
    mode: str = "flatten",
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Compress a PDF file using either 'flatten' mode (pypdfium2 + Pillow)
    or 'vector' mode (pypdf image re-encoding + stream compression).
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    original_size_bytes = in_path.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)

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
    compressed_size_mb = round(compressed_size_bytes / (1024 * 1024), 4)

    if original_size_bytes > 0:
        reduction_percent = round(
            (1.0 - (compressed_size_bytes / float(original_size_bytes))) * 100.0, 2
        )
    else:
        reduction_percent = 0.0

    after_preview_base64 = _render_page_preview_base64(
        str(out_path), page_index=0, scale=1.0, quality=85
    )

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
    }
