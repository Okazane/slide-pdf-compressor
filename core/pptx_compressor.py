import io
import os
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Optional

from PIL import Image, ImageDraw
from pptx import Presentation

from .pdf_compressor import (
    _notify_progress,
    _pil_to_base64_jpeg,
    compress_pdf,
)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp"}


def inspect_pptx(pptx_path: str) -> dict:
    """
    Inspect a PowerPoint (.pptx / .ppt) file and return slide count, media count, and file size.
    """
    path_obj = Path(pptx_path)
    if not path_obj.exists():
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")

    original_size_bytes = path_obj.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)

    slide_count = 0
    media_count = 0

    try:
        prs = Presentation(str(path_obj))
        slide_count = len(prs.slides)
    except Exception:
        pass

    try:
        if zipfile.is_zipfile(str(path_obj)):
            with zipfile.ZipFile(str(path_obj), "r") as zf:
                names = zf.namelist()
                if slide_count == 0:
                    slide_count = sum(
                        1
                        for n in names
                        if n.startswith("ppt/slides/slide") and n.endswith(".xml")
                    )
                media_count = sum(
                    1
                    for n in names
                    if n.startswith("ppt/media/") and not n.endswith("/")
                )
    except Exception:
        pass

    preview_base64 = None
    try:
        slides_imgs = _render_pptx_slides_pil(str(path_obj), max_slides=1, width=960)
        if slides_imgs:
            preview_base64 = _pil_to_base64_jpeg(slides_imgs[0], quality=80)
    except Exception:
        preview_base64 = None

    result = {
        "slide_count": int(slide_count),
        "media_count": int(media_count),
        "original_size_bytes": int(original_size_bytes),
        "original_size_mb": float(original_size_mb),
    }
    if preview_base64:
        result["preview_base64"] = preview_base64
    return result


def _optimize_media_blob(
    data: bytes,
    ext: str,
    max_dim: int,
    jpeg_quality: int,
    png_colors: int = 256,
) -> bytes:
    """
    Optimize an embedded media image from ppt/media/* while keeping its exact extension/format
    so all XML relationships and content-types remain 100% valid.
    """
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return data

    w, h = im.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        new_w = max(1, int(w * scale))
        new_h = max(1, int(h * scale))
        im = im.resize((new_w, new_h), Image.Resampling.LANCZOS)

    out_buf = io.BytesIO()
    ext_lower = ext.lower()

    try:
        if ext_lower in (".jpg", ".jpeg"):
            if im.mode != "RGB":
                if im.mode in ("RGBA", "LA") or (
                    im.mode == "P" and "transparency" in im.info
                ):
                    rgba = im.convert("RGBA")
                    bg = Image.new("RGB", rgba.size, (255, 255, 255))
                    bg.paste(rgba, mask=rgba.getchannel("A"))
                    im = bg
                else:
                    im = im.convert("RGB")
            im.save(out_buf, format="JPEG", quality=int(jpeg_quality), optimize=True)
        elif ext_lower == ".png":
            # Quantize to 256 (or fewer) colors while preserving PNG format & alpha
            if im.mode not in ("RGB", "RGBA", "P", "L"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode in ("RGB", "RGBA"):
                quantized = im.quantize(colors=int(png_colors))
            else:
                quantized = im
            quantized.save(out_buf, format="PNG", optimize=True, compress_level=9)
        elif ext_lower == ".webp":
            im.save(out_buf, format="WEBP", quality=int(jpeg_quality))
        else:
            return data
    except Exception:
        return data

    optimized = out_buf.getvalue()
    return optimized if len(optimized) < len(data) else data


def compress_pptx_direct(
    input_path: str,
    output_path: str,
    target_mb: float = 10.0,
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Compress a .pptx file directly by optimizing all images in ppt/media/*
    and repacking with ZIP_DEFLATED.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PPTX not found: {input_path}")

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    meta = inspect_pptx(str(in_path))
    original_size_bytes = meta["original_size_bytes"]
    original_size_mb = meta["original_size_mb"]
    slide_count = meta["slide_count"]
    before_preview = meta.get("preview_base64")

    target_bytes = int(target_mb * 1024 * 1024)
    ratio = (
        target_bytes / float(original_size_bytes)
        if original_size_bytes > 0
        else 0.5
    )

    if ratio >= 0.7:
        initial_profile = (1440, 78, 256)
    elif ratio >= 0.4:
        initial_profile = (1152, 68, 256)
    else:
        initial_profile = (900, 58, 192)

    profiles = [
        initial_profile,
        (800, 52, 128),
        (600, 42, 64),
    ]

    with zipfile.ZipFile(str(in_path), "r") as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]

    media_indices = [
        i
        for i, (info, _) in enumerate(entries)
        if info.filename.startswith("ppt/media/")
        and Path(info.filename).suffix.lower() in IMAGE_EXTENSIONS
    ]
    total_media = max(len(media_indices), 1)

    for pass_idx, (max_dim, jpg_q, png_colors) in enumerate(profiles):
        with zipfile.ZipFile(
            str(out_path), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
        ) as zout:
            media_done = 0
            for idx, (info, raw_bytes) in enumerate(entries):
                fname = info.filename
                ext = Path(fname).suffix.lower()
                if fname.startswith("ppt/media/") and ext in IMAGE_EXTENSIONS:
                    new_bytes = _optimize_media_blob(
                        raw_bytes,
                        ext=ext,
                        max_dim=max_dim,
                        jpeg_quality=jpg_q,
                        png_colors=png_colors,
                    )
                    media_done += 1
                    if pass_idx == 0:
                        _notify_progress(
                            progress_callback,
                            media_done,
                            total_media,
                            f"Đang nén tài nguyên ảnh {media_done}/{total_media}...",
                        )
                else:
                    new_bytes = raw_bytes

                zout.writestr(fname, new_bytes)

        compressed_size_bytes = out_path.stat().st_size
        if compressed_size_bytes < target_bytes:
            break

    compressed_size_bytes = out_path.stat().st_size
    compressed_size_mb = round(compressed_size_bytes / (1024 * 1024), 4)
    if original_size_bytes > 0:
        reduction_percent = round(
            (1.0 - (compressed_size_bytes / float(original_size_bytes))) * 100.0, 2
        )
    else:
        reduction_percent = 0.0

    after_preview = None
    try:
        after_imgs = _render_pptx_slides_pil(str(out_path), max_slides=1, width=960)
        if after_imgs:
            after_preview = _pil_to_base64_jpeg(after_imgs[0], quality=80)
    except Exception:
        after_preview = before_preview

    return {
        "output_path": str(out_path),
        "original_size_bytes": int(original_size_bytes),
        "compressed_size_bytes": int(compressed_size_bytes),
        "original_size_mb": float(original_size_mb),
        "compressed_size_mb": float(compressed_size_mb),
        "reduction_percent": float(reduction_percent),
        "slide_count": int(slide_count),
        "page_count": int(slide_count),
        "mode_used": "pptx_direct",
        "before_preview_base64": before_preview,
        "after_preview_base64": after_preview,
    }


def _export_pptx_to_pdf_via_com(input_path: str, temp_pdf_path: str) -> bool:
    """
    Attempt to export .pptx/.ppt to PDF using Microsoft PowerPoint COM automation.
    Returns True if successful, False otherwise.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False

    powerpoint = None
    deck = None
    abs_in = str(Path(input_path).resolve())
    abs_out = str(Path(temp_pdf_path).resolve())

    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        # Open presentation without window (WithWindow=False) and ReadOnly=True
        deck = powerpoint.Presentations.Open(
            abs_in, ReadOnly=True, Untitled=False, WithWindow=False
        )
        # 32 = ppSaveAsPDF
        deck.SaveAs(abs_out, 32)
        return os.path.exists(abs_out) and os.path.getsize(abs_out) > 0
    except Exception:
        return False
    finally:
        if deck is not None:
            try:
                deck.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _render_pptx_slides_pil(
    input_path: str, max_slides: Optional[int] = None, width: int = 1440
) -> list[Image.Image]:
    """
    Pure-Python fallback renderer using python-pptx + Pillow to rasterize slides
    (including picture shapes and text boxes) when PowerPoint COM is unavailable.
    """
    prs = Presentation(input_path)
    emu_w = float(prs.slide_width or 12192000)
    emu_h = float(prs.slide_height or 6858000)
    aspect = emu_h / emu_w if emu_w > 0 else (9.0 / 16.0)
    canvas_w = int(width)
    canvas_h = max(1, int(round(canvas_w * aspect)))

    rendered_slides: list[Image.Image] = []

    for idx, slide in enumerate(prs.slides):
        if max_slides is not None and idx >= max_slides:
            break

        canvas = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
        draw = ImageDraw.Draw(canvas)

        for shape in slide.shapes:
            left_px = int(round((float(getattr(shape, "left", 0) or 0) / emu_w) * canvas_w))
            top_px = int(round((float(getattr(shape, "top", 0) or 0) / emu_h) * canvas_h))
            w_px = max(1, int(round((float(getattr(shape, "width", 0) or 0) / emu_w) * canvas_w)))
            h_px = max(1, int(round((float(getattr(shape, "height", 0) or 0) / emu_h) * canvas_h)))

            # Render picture shapes
            if hasattr(shape, "image"):
                try:
                    blob = shape.image.blob
                    pic = Image.open(io.BytesIO(blob))
                    pic.load()
                    pic_resized = pic.resize((w_px, h_px), Image.Resampling.LANCZOS)
                    if pic_resized.mode in ("RGBA", "LA") or (
                        pic_resized.mode == "P" and "transparency" in pic_resized.info
                    ):
                        rgba = pic_resized.convert("RGBA")
                        canvas.paste(rgba, (left_px, top_px), mask=rgba.getchannel("A"))
                    else:
                        canvas.paste(pic_resized.convert("RGB"), (left_px, top_px))
                except Exception:
                    pass

            # Render text frames
            if getattr(shape, "has_text_frame", False):
                try:
                    text = shape.text_frame.text.strip()
                    if text:
                        draw.text(
                            (max(0, left_px + 8), max(0, top_px + 8)),
                            text,
                            fill=(30, 30, 30),
                        )
                except Exception:
                    pass

        rendered_slides.append(canvas)

    if not rendered_slides:
        rendered_slides.append(Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255)))

    return rendered_slides


def _export_pptx_to_pdf_fallback(input_path: str, temp_pdf_path: str) -> None:
    slides = _render_pptx_slides_pil(input_path, width=1440)
    slides[0].save(
        temp_pdf_path,
        format="PDF",
        resolution=72.0,
        save_all=True,
        append_images=slides[1:],
        quality=92,
    )


def convert_pptx_to_compressed_pdf(
    input_path: str,
    output_path: str,
    target_mb: float = 10.0,
    mode: str = "flatten",
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Convert a .pptx / .ppt presentation to a temporary PDF via PowerPoint COM
    (with a pure-Python python-pptx + Pillow fallback), then compress the PDF
    to fit under target_mb.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input presentation not found: {input_path}")

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    pptx_orig_bytes = in_path.stat().st_size
    pptx_orig_mb = round(pptx_orig_bytes / (1024 * 1024), 4)

    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="pptx_export_")
    os.close(fd)

    try:
        exported_ok = _export_pptx_to_pdf_via_com(str(in_path), temp_pdf_path)
        if not exported_ok:
            _export_pptx_to_pdf_fallback(str(in_path), temp_pdf_path)

        res = compress_pdf(
            input_path=temp_pdf_path,
            output_path=str(out_path),
            target_mb=target_mb,
            mode=mode,
            progress_callback=progress_callback,
        )

        comp_bytes = res["compressed_size_bytes"]
        reduction_pct = (
            round((1.0 - (comp_bytes / float(pptx_orig_bytes))) * 100.0, 2)
            if pptx_orig_bytes > 0
            else 0.0
        )

        res["original_size_bytes"] = int(pptx_orig_bytes)
        res["original_size_mb"] = float(pptx_orig_mb)
        res["reduction_percent"] = float(reduction_pct)
        res["slide_count"] = int(res.get("page_count", 0))
        return res
    finally:
        if os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
            except Exception:
                pass
