import ctypes
import io
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Callable, Iterator, Optional

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

from .pdf_compressor import (
    _notify_progress,
    _pil_to_base64_jpeg,
    compress_pdf,
)

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".gif",
    ".tiff",
    ".tif",
    ".webp",
}

STORED_ZIP_EXTENSIONS = {
    ".mp4",
    ".m4a",
    ".mov",
    ".avi",
    ".wmv",
    ".mp3",
    ".aac",
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
}

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_FONT_CACHE: dict[tuple[int, bool], ImageFont.ImageFont] = {}


def _load_unicode_font(px_size: int, bold: bool = False) -> ImageFont.ImageFont:
    px_size = max(10, min(int(px_size), 240))
    cache_key = (px_size, bool(bold))
    if cache_key in _FONT_CACHE:
        return _FONT_CACHE[cache_key]

    candidates = (
        [
            r"C:\Windows\Fonts\segoeuib.ttf",
            r"C:\Windows\Fonts\arialbd.ttf",
            r"C:\Windows\Fonts\tahomabd.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\tahoma.ttf",
        ]
        if bold
        else [
            r"C:\Windows\Fonts\segoeui.ttf",
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\tahoma.ttf",
        ]
    )

    font_obj: Optional[ImageFont.ImageFont] = None
    for font_path in candidates:
        if os.path.exists(font_path):
            try:
                font_obj = ImageFont.truetype(font_path, size=px_size)
                break
            except Exception:
                continue

    if font_obj is None:
        font_obj = ImageFont.load_default()

    _FONT_CACHE[cache_key] = font_obj
    return font_obj


def inspect_pptx(pptx_path: str) -> dict:
    """
    Inspect a PowerPoint (.pptx / .ppt) file and return slide count, media count,
    file size, is_legacy_ppt flag, and preview_base64 (when available).
    """
    path_obj = Path(pptx_path)
    if not path_obj.exists():
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")

    ext = path_obj.suffix.lower()
    is_zip = zipfile.is_zipfile(str(path_obj))

    if ext == ".pptx" and not is_zip:
        raise ValueError("Tệp .pptx bị hỏng hoặc không đúng cấu trúc ZIP.")

    is_legacy_ppt = bool(ext == ".ppt" and not is_zip)

    original_size_bytes = path_obj.stat().st_size
    original_size_mb = round(original_size_bytes / (1024 * 1024), 4)

    slide_count = 0
    media_count = 0
    preview_base64 = None

    if is_zip:
        try:
            prs = Presentation(str(path_obj))
            slide_count = len(prs.slides)
        except Exception:
            pass

        try:
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
        "is_legacy_ppt": is_legacy_ppt,
    }
    if preview_base64:
        result["preview_base64"] = preview_base64
    return result


def _is_image_completely_opaque(im: Image.Image) -> bool:
    if im.mode in ("RGBA", "LA", "PA"):
        rgba = im.convert("RGBA")
        min_alpha, _ = rgba.getchannel("A").getextrema()
        return min_alpha == 255
    if im.mode == "P":
        if "transparency" in im.info:
            rgba = im.convert("RGBA")
            min_alpha, _ = rgba.getchannel("A").getextrema()
            return min_alpha == 255
        return True
    if "transparency" in im.info:
        rgba = im.convert("RGBA")
        min_alpha, _ = rgba.getchannel("A").getextrema()
        return min_alpha == 255
    return True


def _to_rgb_composited(im: Image.Image) -> Image.Image:
    if im.mode == "RGB":
        return im
    if im.mode in ("RGBA", "LA", "PA") or (
        im.mode == "P" and "transparency" in im.info
    ):
        rgba = im.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.getchannel("A"))
        return bg
    return im.convert("RGB")


def _optimize_media_entry(
    data: bytes,
    ext: str,
    max_dim: int,
    jpeg_quality: int,
    png_colors: int = 256,
    allow_jpeg_rename: bool = True,
) -> tuple[bytes, str]:
    """
    Optimize an embedded media image from ppt/media/*.
    Returns (optimized_bytes, final_extension_with_dot).
    When an opaque large .png/.bmp/.tiff/.tif saves >25% as RGB JPEG and allow_jpeg_rename
    is True, returns (jpeg_bytes, ".jpg").
    """
    ext_lower = ext.lower()
    try:
        im = Image.open(io.BytesIO(data))
        if getattr(im, "is_animated", False) and getattr(im, "n_frames", 1) > 1:
            return data, ext_lower
        im.load()
    except Exception:
        return data, ext_lower

    orig_w, orig_h = im.size
    is_opaque = _is_image_completely_opaque(im)
    is_large = (orig_w * orig_h >= 100_000) or (len(data) >= 60_000)

    work_im = im
    resized = False
    if max(orig_w, orig_h) > max_dim:
        scale = max_dim / float(max(orig_w, orig_h))
        new_w = max(1, int(orig_w * scale))
        new_h = max(1, int(orig_h * scale))
        if work_im.mode == "P":
            work_im = work_im.convert("RGBA" if not is_opaque else "RGB")
        work_im = work_im.resize((new_w, new_h), Image.Resampling.LANCZOS)
        resized = True

    # 1. Compute in-place format optimization
    best_inplace = data
    try:
        inplace_buf = io.BytesIO()
        if ext_lower in (".jpg", ".jpeg"):
            rgb_im = _to_rgb_composited(work_im)
            rgb_im.save(
                inplace_buf, format="JPEG", quality=int(jpeg_quality), optimize=True
            )
            cand = inplace_buf.getvalue()
            if len(cand) < len(data):
                best_inplace = cand
        elif ext_lower == ".png":
            png_im = work_im
            if png_im.mode not in ("RGB", "RGBA", "P", "L"):
                png_im = png_im.convert("RGBA" if not is_opaque else "RGB")
            if png_im.mode in ("RGB", "RGBA"):
                quantized = png_im.quantize(colors=int(png_colors))
            else:
                quantized = png_im
            quantized.save(inplace_buf, format="PNG", optimize=True, compress_level=9)
            cand = inplace_buf.getvalue()
            if len(cand) < len(data):
                best_inplace = cand
        elif ext_lower == ".webp":
            work_im.save(inplace_buf, format="WEBP", quality=int(jpeg_quality))
            cand = inplace_buf.getvalue()
            if len(cand) < len(data):
                best_inplace = cand
        elif ext_lower in (".tiff", ".tif"):
            tiff_im = work_im
            if tiff_im.mode == "P":
                tiff_im = tiff_im.convert("RGBA" if not is_opaque else "RGB")
            tiff_im.save(inplace_buf, format="TIFF", compression="tiff_adobe_deflate")
            cand = inplace_buf.getvalue()
            if len(cand) < len(data):
                best_inplace = cand
        elif ext_lower == ".gif":
            gif_im = work_im
            if gif_im.mode not in ("P", "L"):
                gif_im = gif_im.convert("RGBA" if not is_opaque else "RGB").quantize(
                    colors=min(int(png_colors), 256)
                )
            gif_im.save(inplace_buf, format="GIF", optimize=True)
            cand = inplace_buf.getvalue()
            if len(cand) < len(data):
                best_inplace = cand
        elif ext_lower == ".bmp":
            if resized:
                bmp_im = _to_rgb_composited(work_im)
                bmp_im.save(inplace_buf, format="BMP")
                cand = inplace_buf.getvalue()
                if len(cand) < len(data):
                    best_inplace = cand
    except Exception:
        best_inplace = data

    # 2. Smart Opaque PNG/BMP/TIFF -> JPEG conversion
    if (
        allow_jpeg_rename
        and ext_lower in (".png", ".bmp", ".tiff", ".tif")
        and is_opaque
        and is_large
    ):
        try:
            jpg_buf = io.BytesIO()
            rgb_im = _to_rgb_composited(work_im)
            rgb_im.save(
                jpg_buf, format="JPEG", quality=int(jpeg_quality), optimize=True
            )
            jpg_bytes = jpg_buf.getvalue()
            if len(jpg_bytes) < int(len(best_inplace) * 0.75):
                return jpg_bytes, ".jpg"
        except Exception:
            pass

    return best_inplace, ext_lower


def _optimize_media_blob(
    data: bytes,
    ext: str,
    max_dim: int,
    jpeg_quality: int,
    png_colors: int = 256,
) -> bytes:
    """Backward-compatible helper keeping original extension."""
    opt_bytes, _ = _optimize_media_entry(
        data,
        ext=ext,
        max_dim=max_dim,
        jpeg_quality=jpeg_quality,
        png_colors=png_colors,
        allow_jpeg_rename=False,
    )
    return opt_bytes


def _find_hardcoded_media_basenames(zin: zipfile.ZipFile, media_basenames: set[str]) -> set[str]:
    """
    Scan non-.rels XML parts (excluding [Content_Types].xml) to verify whether any XML
    directly hardcodes a media filename outside relationship Target attributes.
    """
    if not media_basenames:
        return set()
    hardcoded: set[str] = set()
    for info in zin.infolist():
        fname = info.filename
        if (
            not fname.lower().endswith(".xml")
            or fname == "[Content_Types].xml"
            or info.is_dir()
        ):
            continue
        try:
            xml_text = zin.read(fname).decode("utf-8", errors="ignore")
            if "media" not in xml_text and "." not in xml_text:
                continue
            for bname in media_basenames:
                if bname in xml_text:
                    hardcoded.add(bname)
        except Exception:
            continue
    return hardcoded


def _rewrite_rels_xml(
    raw_xml: bytes, rels_zip_path: str, rename_map: dict[str, str]
) -> bytes:
    if not rename_map:
        return raw_xml

    try:
        root = ET.fromstring(raw_xml)
    except Exception:
        return raw_xml

    rels_dir = posixpath.dirname(rels_zip_path)
    if posixpath.basename(rels_dir).lower() == "_rels":
        source_dir = posixpath.dirname(rels_dir)
    else:
        source_dir = rels_dir

    changed = False
    for elem in root.iter():
        if not elem.tag.endswith("Relationship"):
            continue
        if (elem.get("TargetMode") or "").strip().lower() == "external":
            continue
        target = elem.get("Target")
        if not target:
            continue

        if target.startswith("/"):
            resolved = posixpath.normpath(target.lstrip("/"))
        else:
            resolved = (
                posixpath.normpath(posixpath.join(source_dir, target))
                if source_dir
                else posixpath.normpath(target)
            )

        if resolved in rename_map:
            new_zip_path = rename_map[resolved]
            if target.startswith("/"):
                new_target = "/" + new_zip_path
            else:
                new_target = (
                    posixpath.relpath(new_zip_path, source_dir)
                    if source_dir
                    else new_zip_path
                )
            elem.set("Target", new_target)
            changed = True

    if not changed:
        return raw_xml

    ns = RELS_NS
    if root.tag.startswith("{") and "}" in root.tag:
        ns = root.tag[1 : root.tag.index("}")]
    ET.register_namespace("", ns)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rewrite_content_types_xml(raw_xml: bytes, rename_map: dict[str, str]) -> bytes:
    if not rename_map:
        return raw_xml

    try:
        root = ET.fromstring(raw_xml)
    except Exception:
        return raw_xml

    ns = CONTENT_TYPES_NS
    if root.tag.startswith("{") and "}" in root.tag:
        ns = root.tag[1 : root.tag.index("}")]

    has_jpg_default = False
    to_remove = []
    for child in list(root):
        local_tag = child.tag.split("}")[-1]
        if local_tag == "Default":
            if (child.get("Extension") or "").strip().lower() == "jpg":
                has_jpg_default = True
        elif local_tag == "Override":
            part_name = (child.get("PartName") or "").lstrip("/")
            if posixpath.normpath(part_name) in rename_map:
                to_remove.append(child)

    for child in to_remove:
        root.remove(child)

    if not has_jpg_default:
        default_elem = ET.Element(
            f"{{{ns}}}Default",
            {"Extension": "jpg", "ContentType": "image/jpeg"},
        )
        root.insert(0, default_elem)

    ET.register_namespace("", ns)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _make_zip_info(filename: str, date_time: tuple) -> zipfile.ZipInfo:
    safe_dt = date_time if (date_time and date_time[0] >= 1980) else (2025, 1, 1, 0, 0, 0)
    zinfo = zipfile.ZipInfo(filename=filename, date_time=safe_dt)
    ext = Path(filename).suffix.lower()
    zinfo.compress_type = (
        zipfile.ZIP_STORED if ext in STORED_ZIP_EXTENSIONS else zipfile.ZIP_DEFLATED
    )
    return zinfo


def compress_pptx_direct(
    input_path: str,
    output_path: str,
    target_mb: float = 10.0,
    progress_callback: Optional[Callable] = None,
) -> dict:
    """
    Compress a .pptx file directly by optimizing embedded media images
    (including smart opaque PNG/BMP/TIFF -> JPEG conversion with XML relationship updates)
    and streaming non-image entries with low RAM usage.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PPTX not found: {input_path}")

    ext_in = in_path.suffix.lower()
    is_zip = zipfile.is_zipfile(str(in_path))
    if ext_in == ".ppt" or not is_zip:
        if ext_in == ".pptx" and not is_zip:
            raise ValueError("Tệp .pptx bị hỏng hoặc không đúng cấu trúc ZIP.")
        raise ValueError(
            "Định dạng .ppt cũ không hỗ trợ nén trực tiếp giữ nguyên định dạng. "
            "Vui lòng chọn xuất sang PDF."
        )

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
        (480, 36, 64),
    ]

    fd, temp_zip_str = tempfile.mkstemp(suffix=".pptx", prefix="pptx_comp_")
    os.close(fd)
    temp_zip_path = Path(temp_zip_str)

    try:
        with zipfile.ZipFile(str(in_path), "r") as zin:
            all_infos = zin.infolist()
            media_infos = [
                info
                for info in all_infos
                if info.filename.startswith("ppt/media/")
                and not info.is_dir()
                and Path(info.filename).suffix.lower() in IMAGE_EXTENSIONS
            ]
            total_media = max(len(media_infos), 1)
            candidate_basenames = {
                Path(info.filename).name
                for info in media_infos
                if Path(info.filename).suffix.lower() in (".png", ".bmp", ".tiff", ".tif")
            }
            hardcoded_basenames = _find_hardcoded_media_basenames(
                zin, candidate_basenames
            )
            base_names_lower = {info.filename.lower() for info in all_infos}

            compressed_size_bytes = original_size_bytes

            for pass_idx, (max_dim, jpg_q, png_colors) in enumerate(profiles):
                used_names_lower = set(base_names_lower)
                rename_map: dict[str, str] = {}
                optimized_media: dict[str, tuple[str, bytes]] = {}

                for m_idx, m_info in enumerate(media_infos):
                    orig_fname = posixpath.normpath(m_info.filename)
                    orig_ext = Path(orig_fname).suffix.lower()
                    bname = Path(orig_fname).name
                    raw_bytes = zin.read(m_info.filename)

                    can_rename = bname not in hardcoded_basenames
                    opt_bytes, final_ext = _optimize_media_entry(
                        raw_bytes,
                        ext=orig_ext,
                        max_dim=max_dim,
                        jpeg_quality=jpg_q,
                        png_colors=png_colors,
                        allow_jpeg_rename=can_rename,
                    )

                    if final_ext != orig_ext and final_ext == ".jpg":
                        used_names_lower.discard(orig_fname.lower())
                        parent_dir = posixpath.dirname(orig_fname)
                        stem = Path(orig_fname).stem
                        cand_name = f"{parent_dir}/{stem}.jpg"
                        suffix_idx = 1
                        while cand_name.lower() in used_names_lower:
                            cand_name = f"{parent_dir}/{stem}_{suffix_idx}.jpg"
                            suffix_idx += 1
                        used_names_lower.add(cand_name.lower())
                        rename_map[orig_fname] = cand_name
                        target_fname = cand_name
                    else:
                        target_fname = m_info.filename

                    optimized_media[m_info.filename] = (target_fname, opt_bytes)

                    _notify_progress(
                        progress_callback,
                        m_idx + 1,
                        total_media,
                        f"Đang nén tài nguyên ảnh {m_idx + 1}/{len(media_infos)}...",
                    )

                with zipfile.ZipFile(
                    str(temp_zip_path),
                    "w",
                    compression=zipfile.ZIP_DEFLATED,
                    compresslevel=9,
                ) as zout:
                    for info in all_infos:
                        if info.is_dir():
                            continue
                        fname = info.filename
                        if fname in optimized_media:
                            target_fname, opt_bytes = optimized_media[fname]
                            zinfo = _make_zip_info(target_fname, info.date_time)
                            zout.writestr(zinfo, opt_bytes)
                        elif fname == "[Content_Types].xml" and rename_map:
                            raw_ct = zin.read(fname)
                            new_ct = _rewrite_content_types_xml(raw_ct, rename_map)
                            zinfo = _make_zip_info(fname, info.date_time)
                            zout.writestr(zinfo, new_ct)
                        elif fname.lower().endswith(".rels") and rename_map:
                            raw_rels = zin.read(fname)
                            new_rels = _rewrite_rels_xml(
                                raw_rels, posixpath.normpath(fname), rename_map
                            )
                            zinfo = _make_zip_info(fname, info.date_time)
                            zout.writestr(zinfo, new_rels)
                        else:
                            zinfo = _make_zip_info(fname, info.date_time)
                            with zin.open(info, "r") as src, zout.open(zinfo, "w") as dst:
                                shutil.copyfileobj(src, dst, length=262144)

                compressed_size_bytes = temp_zip_path.stat().st_size
                if compressed_size_bytes <= target_bytes:
                    break

        # Non-inflation guard when original file is already within target_bytes
        if (
            compressed_size_bytes >= original_size_bytes
            and original_size_bytes <= target_bytes
        ):
            if in_path.resolve() != out_path.resolve():
                shutil.copy2(str(in_path), str(out_path))
        else:
            shutil.copy2(str(temp_zip_path), str(out_path))
    finally:
        if temp_zip_path.exists():
            try:
                temp_zip_path.unlink()
            except Exception:
                pass

    compressed_size_bytes = out_path.stat().st_size
    compressed_size_mb = round(compressed_size_bytes / (1024 * 1024), 4)
    if original_size_bytes > 0:
        reduction_percent = round(
            (1.0 - (compressed_size_bytes / float(original_size_bytes))) * 100.0, 2
        )
    else:
        reduction_percent = 0.0

    target_met = bool(compressed_size_bytes <= target_bytes)

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
        "target_met": target_met,
        "before_preview_base64": before_preview,
        "after_preview_base64": after_preview,
    }


def _get_pids_with_visible_main_window() -> set[int]:
    """Return PIDs that currently own a visible top-level window on Windows."""
    if os.name != "nt":
        return set()
    visible_pids: set[int] = set()
    try:
        user32 = ctypes.windll.user32
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def _enum_cb(hwnd, _lparam):
            try:
                if user32.IsWindowVisible(hwnd) and user32.GetWindow(hwnd, 4) == 0:
                    if user32.GetWindowTextLengthW(hwnd) > 0:
                        pid = ctypes.c_ulong(0)
                        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                        if pid.value:
                            visible_pids.add(int(pid.value))
            except Exception:
                pass
            return True

        user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
    except Exception:
        pass
    return visible_pids


def _list_powerpoint_pids() -> set[int]:
    """Return all currently running POWERPNT.EXE PIDs."""
    pids: set[int] = set()
    try:
        import psutil

        for proc in psutil.process_iter(["pid", "name"]):
            name = (proc.info.get("name") or "").upper()
            if name in ("POWERPNT.EXE", "POWERPNT"):
                pids.add(int(proc.info["pid"]))
        return pids
    except Exception:
        pass

    if os.name == "nt":
        try:
            res = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            for line in res.stdout.splitlines():
                parts = [p.strip('"') for p in line.split('","')]
                if len(parts) >= 2 and parts[0].upper().startswith("POWERPNT"):
                    if parts[1].isdigit():
                        pids.add(int(parts[1]))
        except Exception:
            pass
    return pids


def _terminate_new_headless_powerpoint(pre_pids: set[int]) -> None:
    """Terminate ONLY newly-spawned headless POWERPNT.EXE processes (MainWindowHandle == 0)."""
    post_pids = _list_powerpoint_pids()
    new_pids = post_pids - pre_pids
    if not new_pids:
        return

    visible_pids = _get_pids_with_visible_main_window()
    for pid in new_pids:
        if pid in visible_pids:
            continue
        killed = False
        try:
            import psutil

            p = psutil.Process(pid)
            p.kill()
            p.wait(timeout=3)
            killed = True
        except Exception:
            pass

        if not killed and os.name == "nt":
            try:
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    capture_output=True,
                    timeout=5,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception:
                pass


_COM_BROKEN_UNTIL: float = 0.0

_COM_CHILD_SCRIPT = """
import gc
import os
import sys

abs_in = sys.argv[1]
abs_out = sys.argv[2]

import pythoncom
import win32com.client

pythoncom.CoInitialize()
powerpoint = None
deck = None
try:
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    sys.stdout.write("DISPATCH_OK\\n")
    sys.stdout.flush()
    try:
        powerpoint.DisplayAlerts = 1
    except Exception:
        pass
    deck = powerpoint.Presentations.Open(
        abs_in, ReadOnly=True, Untitled=False, WithWindow=False
    )
    deck.SaveAs(abs_out, 32)
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
    del deck
    del powerpoint
    gc.collect()
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

ok = os.path.exists(abs_out) and os.path.getsize(abs_out) > 0
sys.exit(0 if ok else 1)
"""


def _export_pptx_to_pdf_via_com(
    input_path: str, temp_pdf_path: str, timeout: float = 25.0
) -> bool:
    """
    Export .pptx/.ppt to PDF using Microsoft PowerPoint COM automation inside an isolated
    child subprocess with a strict timeout and zero leaked headless POWERPNT.EXE processes.
    """
    global _COM_BROKEN_UNTIL

    if os.name != "nt":
        return False
    if time.monotonic() < _COM_BROKEN_UNTIL:
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False

    abs_in = str(Path(input_path).resolve())
    abs_out = str(Path(temp_pdf_path).resolve())
    pre_pids = _list_powerpoint_pids()

    proc: Optional[subprocess.Popen] = None
    timed_out = False
    stdout_bytes = b""
    try:
        proc = subprocess.Popen(
            [sys.executable, "-c", _COM_CHILD_SCRIPT, abs_in, abs_out],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        stdout_bytes, _ = proc.communicate(timeout=timeout)
        if proc.returncode != 0 and b"DISPATCH_OK" not in (stdout_bytes or b""):
            _COM_BROKEN_UNTIL = time.monotonic() + 180.0
        return (
            proc.returncode == 0
            and os.path.exists(abs_out)
            and os.path.getsize(abs_out) > 0
        )
    except subprocess.TimeoutExpired:
        timed_out = True
        return False
    except Exception:
        return False
    finally:
        if proc is not None and proc.poll() is None:
            try:
                proc.kill()
                rem_out, _ = proc.communicate(timeout=5)
                if rem_out:
                    stdout_bytes += rem_out
            except Exception:
                try:
                    proc.wait(timeout=3)
                except Exception:
                    pass
        _terminate_new_headless_powerpoint(pre_pids)
        if timed_out:
            time.sleep(0.15)
            _terminate_new_headless_powerpoint(pre_pids)
            if b"DISPATCH_OK" not in (stdout_bytes or b""):
                _COM_BROKEN_UNTIL = time.monotonic() + 180.0


def _extract_fill_rgb(fill_obj) -> Optional[tuple[int, int, int]]:
    try:
        if fill_obj is not None and fill_obj.type == 1:
            rgb = fill_obj.fore_color.rgb
            if rgb is not None:
                return (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    except Exception:
        pass
    return None


def _get_slide_background_rgb(slide) -> tuple[int, int, int]:
    for obj in (
        slide,
        getattr(slide, "slide_layout", None),
        getattr(getattr(slide, "slide_layout", None), "slide_master", None),
    ):
        if obj is None:
            continue
        try:
            rgb = _extract_fill_rgb(obj.background.fill)
            if rgb is not None:
                return rgb
        except Exception:
            continue
    return (255, 255, 255)


def _wrap_paragraph_lines(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
    max_width_px: int,
) -> list[str]:
    if not text:
        return []
    words = text.split(" ")
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        try:
            width = draw.textlength(candidate, font=font)
        except Exception:
            width = len(candidate) * 8
        if width <= max_width_px or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _render_shape_recursive(
    shape,
    canvas: Image.Image,
    draw: ImageDraw.ImageDraw,
    emu_w: float,
    emu_h: float,
    canvas_w: int,
    canvas_h: int,
    slide_bg_rgb: tuple[int, int, int],
    offset_x: float = 0.0,
    offset_y: float = 0.0,
    scale_x: float = 1.0,
    scale_y: float = 1.0,
    ch_off_x: float = 0.0,
    ch_off_y: float = 0.0,
) -> None:
    raw_left = float(getattr(shape, "left", 0) or 0)
    raw_top = float(getattr(shape, "top", 0) or 0)
    raw_w = float(getattr(shape, "width", 0) or 0)
    raw_h = float(getattr(shape, "height", 0) or 0)

    s_left = offset_x + (raw_left - ch_off_x) * scale_x
    s_top = offset_y + (raw_top - ch_off_y) * scale_y
    s_w = raw_w * scale_x
    s_h = raw_h * scale_y

    # Recurse into group shapes
    if hasattr(shape, "shapes"):
        ch_x, ch_y, ch_cx, ch_cy = 0.0, 0.0, 0.0, 0.0
        try:
            xfrm = shape._element.grpSpPr.xfrm
            if xfrm is not None:
                if xfrm.chOff is not None:
                    ch_x = float(xfrm.chOff.x)
                    ch_y = float(xfrm.chOff.y)
                if xfrm.chExt is not None:
                    ch_cx = float(xfrm.chExt.cx)
                    ch_cy = float(xfrm.chExt.cy)
        except Exception:
            pass

        if ch_cx > 0 and ch_cy > 0 and s_w > 0 and s_h > 0:
            child_sx = s_w / ch_cx
            child_sy = s_h / ch_cy
            child_ox = s_left
            child_oy = s_top
            child_chx = ch_x
            child_chy = ch_y
        else:
            child_sx = scale_x
            child_sy = scale_y
            child_ox = s_left
            child_oy = s_top
            child_chx = 0.0
            child_chy = 0.0

        for child in shape.shapes:
            _render_shape_recursive(
                child,
                canvas,
                draw,
                emu_w,
                emu_h,
                canvas_w,
                canvas_h,
                slide_bg_rgb,
                offset_x=child_ox,
                offset_y=child_oy,
                scale_x=child_sx,
                scale_y=child_sy,
                ch_off_x=child_chx,
                ch_off_y=child_chy,
            )
        return

    left_px = int(round((s_left / emu_w) * canvas_w))
    top_px = int(round((s_top / emu_h) * canvas_h))
    w_px = max(1, int(round((s_w / emu_w) * canvas_w)))
    h_px = max(1, int(round((s_h / emu_h) * canvas_h)))

    shape_fill_rgb: Optional[tuple[int, int, int]] = None
    if not hasattr(shape, "image") and hasattr(shape, "fill"):
        shape_fill_rgb = _extract_fill_rgb(shape.fill)
        if shape_fill_rgb is not None:
            draw.rectangle(
                [left_px, top_px, left_px + w_px, top_px + h_px],
                fill=shape_fill_rgb,
            )

    # Render picture shapes
    if hasattr(shape, "image"):
        try:
            blob = shape.image.blob
            pic = Image.open(io.BytesIO(blob))
            pic.load()
            pic_resized = pic.resize((w_px, h_px), Image.Resampling.LANCZOS)
            if pic_resized.mode in ("RGBA", "LA", "PA") or (
                pic_resized.mode == "P" and "transparency" in pic_resized.info
            ):
                rgba = pic_resized.convert("RGBA")
                canvas.paste(rgba, (left_px, top_px), mask=rgba.getchannel("A"))
            else:
                canvas.paste(pic_resized.convert("RGB"), (left_px, top_px))
        except Exception:
            pass

    # Render text frames with Vietnamese Unicode TTF + word wrapping
    if getattr(shape, "has_text_frame", False):
        try:
            slide_w_pt = max(emu_w / 12700.0, 1.0)
            px_per_pt = canvas_w / slide_w_pt

            eff_bg = shape_fill_rgb if shape_fill_rgb is not None else slide_bg_rgb
            bg_lum = 0.299 * eff_bg[0] + 0.587 * eff_bg[1] + 0.114 * eff_bg[2]
            default_fg = (245, 245, 245) if bg_lum < 110 else (30, 30, 30)

            pad_x = max(4, int(round(canvas_w * 0.006)))
            pad_y = max(4, int(round(canvas_h * 0.006)))
            max_text_w = max(40, w_px - 2 * pad_x)
            curr_y = max(0, top_px + pad_y)
            max_y = min(canvas_h - 4, max(top_px + h_px + int(canvas_h * 0.08), curr_y + 40))

            for para in shape.text_frame.paragraphs:
                raw_para_text = para.text
                if not raw_para_text or not raw_para_text.strip():
                    curr_y += max(6, int(round(10 * px_per_pt)))
                    continue

                pt_size = 18.0
                is_bold = False
                fg_rgb = default_fg

                try:
                    if para.font.size is not None:
                        pt_size = float(para.font.size.pt)
                    if para.font.bold is not None:
                        is_bold = bool(para.font.bold)
                    if para.font.color is not None and para.font.color.rgb is not None:
                        c = para.font.color.rgb
                        fg_rgb = (int(c[0]), int(c[1]), int(c[2]))
                except Exception:
                    pass

                for run in para.runs:
                    try:
                        if pt_size == 18.0 and run.font.size is not None:
                            pt_size = float(run.font.size.pt)
                        if not is_bold and run.font.bold:
                            is_bold = True
                        if (
                            fg_rgb == default_fg
                            and run.font.color is not None
                            and run.font.color.rgb is not None
                        ):
                            c = run.font.color.rgb
                            fg_rgb = (int(c[0]), int(c[1]), int(c[2]))
                    except Exception:
                        continue

                font_px = max(11, int(round(pt_size * px_per_pt)))
                font_obj = _load_unicode_font(font_px, bold=is_bold)
                line_step = max(font_px + 4, int(round(font_px * 1.28)))

                for raw_line in raw_para_text.splitlines():
                    wrapped = _wrap_paragraph_lines(draw, raw_line.strip(), font_obj, max_text_w)
                    for line_str in wrapped:
                        if curr_y > max_y:
                            break
                        draw.text(
                            (max(0, left_px + pad_x), curr_y),
                            line_str,
                            fill=fg_rgb,
                            font=font_obj,
                        )
                        curr_y += line_step
        except Exception:
            pass


def _iter_pptx_slides_pil(
    input_path: str, max_slides: Optional[int] = None, width: int = 1440
) -> Iterator[tuple[int, int, Image.Image]]:
    """
    Stream rendered slide images one by one using python-pptx + Pillow + Windows Unicode TTF.
    Yields (slide_index, total_slides, pil_image).
    """
    prs = Presentation(input_path)
    emu_w = float(prs.slide_width or 12192000)
    emu_h = float(prs.slide_height or 6858000)
    aspect = emu_h / emu_w if emu_w > 0 else (9.0 / 16.0)
    canvas_w = int(width)
    canvas_h = max(1, int(round(canvas_w * aspect)))

    total_slides = len(prs.slides)
    if total_slides == 0:
        yield 0, 1, Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
        return

    limit = total_slides if max_slides is None else min(total_slides, int(max_slides))
    for idx, slide in enumerate(prs.slides):
        if idx >= limit:
            break

        bg_rgb = _get_slide_background_rgb(slide)
        canvas = Image.new("RGB", (canvas_w, canvas_h), bg_rgb)
        draw = ImageDraw.Draw(canvas)

        for shape in slide.shapes:
            _render_shape_recursive(
                shape,
                canvas,
                draw,
                emu_w=emu_w,
                emu_h=emu_h,
                canvas_w=canvas_w,
                canvas_h=canvas_h,
                slide_bg_rgb=bg_rgb,
            )

        yield idx, limit, canvas


def _render_pptx_slides_pil(
    input_path: str, max_slides: Optional[int] = None, width: int = 1440
) -> list[Image.Image]:
    """
    Pure-Python fallback renderer using python-pptx + Pillow to rasterize slides
    (including backgrounds, shape fills, group shapes, pictures, and Vietnamese Unicode text).
    """
    return [
        img
        for _, _, img in _iter_pptx_slides_pil(
            input_path, max_slides=max_slides, width=width
        )
    ]


def _export_pptx_to_pdf_fallback(input_path: str, temp_pdf_path: str) -> None:
    prs = Presentation(input_path)
    emu_w = float(prs.slide_width or 12192000)
    inches_w = max(emu_w / 914400.0, 1.0)
    render_w = 1440
    dpi = round(render_w / inches_w, 2)

    slides = _render_pptx_slides_pil(input_path, width=render_w)
    slides[0].save(
        temp_pdf_path,
        format="PDF",
        resolution=dpi,
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
    Convert a .pptx / .ppt presentation to a temporary PDF via isolated PowerPoint COM
    (with a pure-Python python-pptx + Pillow fallback for .pptx), then compress the PDF
    to fit under target_mb.
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input presentation not found: {input_path}")

    ext_in = in_path.suffix.lower()
    is_zip = zipfile.is_zipfile(str(in_path))
    if ext_in == ".pptx" and not is_zip:
        raise ValueError("Tệp .pptx bị hỏng hoặc không đúng cấu trúc ZIP.")

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    pptx_orig_bytes = in_path.stat().st_size
    pptx_orig_mb = round(pptx_orig_bytes / (1024 * 1024), 4)
    target_bytes = int(target_mb * 1024 * 1024)

    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="pptx_export_")
    os.close(fd)

    try:
        exported_ok = _export_pptx_to_pdf_via_com(str(in_path), temp_pdf_path)
        if not exported_ok:
            if not is_zip:
                raise RuntimeError(
                    "Không thể chuyển đổi tệp .ppt cũ do Microsoft PowerPoint không phản hồi. "
                    "Vui lòng lưu tệp sang .pptx hoặc .pdf rồi thử lại."
                )
            _export_pptx_to_pdf_fallback(str(in_path), temp_pdf_path)

        res = compress_pdf(
            input_path=temp_pdf_path,
            output_path=str(out_path),
            target_mb=target_mb,
            mode=mode,
            progress_callback=progress_callback,
        )

        comp_bytes = int(res["compressed_size_bytes"])
        reduction_pct = (
            round((1.0 - (comp_bytes / float(pptx_orig_bytes))) * 100.0, 2)
            if pptx_orig_bytes > 0
            else 0.0
        )

        res["original_size_bytes"] = int(pptx_orig_bytes)
        res["original_size_mb"] = float(pptx_orig_mb)
        res["reduction_percent"] = float(reduction_pct)
        res["slide_count"] = int(res.get("page_count", 0))
        res["target_met"] = bool(comp_bytes <= target_bytes)
        return res
    finally:
        if os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
            except Exception:
                pass
