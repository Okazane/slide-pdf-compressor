import os
import sys
import threading
import time
import traceback
import uuid
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"

UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
STATIC_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(
    title="Slide & PDF Smart Compressor",
    description="Hệ thống nén & trích xuất tài nguyên PDF / PowerPoint thông minh giữ nguyên độ nét",
    version="1.1.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# In-memory registries
FILES: Dict[str, Dict[str, Any]] = {}
JOBS: Dict[str, Dict[str, Any]] = {}

ALLOWED_EXTENSIONS = {".pdf", ".pptx", ".ppt"}
UPLOAD_CHUNK_SIZE = 1024 * 1024  # 1 MB
MAX_UPLOAD_BYTES = 500 * 1024 * 1024  # 500 MB
DEFAULT_CLEANUP_TTL_SECONDS = 7200  # 2 hours


def _ext_to_file_type(ext: str) -> str:
    ext_lower = ext.lower()
    if ext_lower == ".pdf":
        return "pdf"
    if ext_lower == ".ppt":
        return "ppt"
    return "pptx"


def _get_core_functions():
    from core.pdf_compressor import compress_pdf, inspect_pdf
    from core.pptx_compressor import (
        compress_pptx_direct,
        convert_pptx_to_compressed_pdf,
        inspect_pptx,
    )

    return (
        inspect_pdf,
        compress_pdf,
        inspect_pptx,
        compress_pptx_direct,
        convert_pptx_to_compressed_pdf,
    )


def _get_extract_function():
    from core.extractor import extract_archive

    return extract_archive


def _cleanup_old_files(max_age_seconds: int = DEFAULT_CLEANUP_TTL_SECONDS) -> int:
    """Delete files older than max_age_seconds in UPLOADS_DIR/OUTPUTS_DIR and prune stale registries."""
    now = time.time()
    removed = 0

    running_jobs = [j for j in list(JOBS.values()) if j.get("status") == "running"]
    active_file_ids = {str(j.get("file_id")) for j in running_jobs if j.get("file_id")}
    active_job_ids = {str(j.get("job_id")) for j in running_jobs if j.get("job_id")}
    active_paths = set()

    for fid in active_file_ids:
        finfo = FILES.get(fid)
        if finfo and finfo.get("path"):
            try:
                active_paths.add(Path(finfo["path"]).resolve())
            except OSError:
                pass

    for j in running_jobs:
        if j.get("output_path"):
            try:
                active_paths.add(Path(j["output_path"]).resolve())
            except OSError:
                pass

    for directory in (UPLOADS_DIR, OUTPUTS_DIR):
        if not directory.exists():
            continue
        for item in list(directory.iterdir()):
            if not item.is_file() or item.name.startswith("."):
                continue
            try:
                if item.resolve() in active_paths:
                    continue
                if any(item.name.startswith(fid) for fid in active_file_ids):
                    continue
                if any(item.name.startswith(jid) for jid in active_job_ids):
                    continue
                if (now - item.stat().st_mtime) > max_age_seconds:
                    item.unlink(missing_ok=True)
                    removed += 1
            except OSError:
                continue

    for fid, info in list(FILES.items()):
        if fid in active_file_ids:
            continue
        fpath = Path(info["path"]) if info.get("path") else None
        created_at = float(info.get("created_at", now))
        expired = (now - created_at) > max_age_seconds
        missing = fpath is not None and not fpath.exists()
        if expired or missing:
            if fpath is not None and fpath.exists():
                try:
                    fpath.unlink(missing_ok=True)
                    removed += 1
                except OSError:
                    pass
            FILES.pop(fid, None)

    for jid, job in list(JOBS.items()):
        if jid in active_job_ids or job.get("status") == "running":
            continue
        out_p = Path(job["output_path"]) if job.get("output_path") else None
        created_at = float(job.get("created_at", now))
        expired = (now - created_at) > max_age_seconds
        missing_done_output = (
            job.get("status") == "done" and out_p is not None and not out_p.exists()
        )
        if expired or missing_done_output:
            if out_p is not None and out_p.exists():
                try:
                    out_p.unlink(missing_ok=True)
                    removed += 1
                except OSError:
                    pass
            JOBS.pop(jid, None)

    return removed


async def _parse_request_payload(request: Request) -> Dict[str, Any]:
    payload: Dict[str, Any] = {}
    content_type = request.headers.get("content-type", "").lower()

    if "application/json" in content_type:
        try:
            payload = await request.json()
        except Exception:
            payload = {}
    elif "form" in content_type:
        try:
            form_data = await request.form()
            payload = dict(form_data)
        except Exception:
            payload = {}
    else:
        try:
            payload = await request.json()
        except Exception:
            try:
                form_data = await request.form()
                payload = dict(form_data)
            except Exception:
                payload = {}

    if not payload and request.query_params:
        payload = dict(request.query_params)

    return payload if isinstance(payload, dict) else {}


def _resolve_file_info(file_id: str) -> Dict[str, Any]:
    if file_id in FILES:
        info = FILES[file_id]
        if Path(info["path"]).exists():
            return info

    # Fallback: look up in UPLOADS_DIR by prefix
    for candidate in UPLOADS_DIR.glob(f"{file_id}*"):
        if candidate.is_file() and candidate.suffix.lower() in ALLOWED_EXTENSIONS:
            ext = candidate.suffix.lower()
            file_type = _ext_to_file_type(ext)
            size_bytes = candidate.stat().st_size
            info = {
                "file_id": file_id,
                "filename": candidate.name,
                "ext": ext,
                "file_type": file_type,
                "is_legacy_ppt": bool(file_type == "ppt"),
                "path": str(candidate),
                "original_size_bytes": size_bytes,
                "original_size_mb": round(size_bytes / (1024 * 1024), 2),
                "created_at": candidate.stat().st_mtime,
            }
            FILES[file_id] = info
            return info

    raise HTTPException(status_code=404, detail=f"Không tìm thấy file_id: {file_id}")


@app.get("/")
async def index():
    index_path = STATIC_DIR / "index.html"
    if not index_path.exists():
        return JSONResponse({"message": "Slide & PDF Smart Compressor API is running."})
    return FileResponse(str(index_path), media_type="text/html; charset=utf-8")


@app.post("/api/analyze")
async def analyze_file(file: UploadFile = File(...)):
    try:
        _cleanup_old_files(DEFAULT_CLEANUP_TTL_SECONDS)
    except Exception:
        pass

    if not file or not file.filename:
        raise HTTPException(status_code=400, detail="Vui lòng chọn tệp để tải lên.")

    original_filename = Path(file.filename).name
    ext = Path(original_filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Định dạng '{ext}' không được hỗ trợ. Chỉ hỗ trợ .pdf, .pptx, .ppt",
        )

    file_id = uuid.uuid4().hex[:12]
    saved_path = UPLOADS_DIR / f"{file_id}{ext}"

    total_written = 0
    try:
        with open(saved_path, "wb") as f:
            while chunk := await file.read(UPLOAD_CHUNK_SIZE):
                total_written += len(chunk)
                if total_written > MAX_UPLOAD_BYTES:
                    f.close()
                    saved_path.unlink(missing_ok=True)
                    raise HTTPException(
                        status_code=413,
                        detail="Dung lượng tệp vượt quá giới hạn tối đa 500 MB.",
                    )
                f.write(chunk)
    except HTTPException:
        saved_path.unlink(missing_ok=True)
        raise
    except Exception as exc:
        saved_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"Lỗi lưu tệp tải lên: {exc}")
    finally:
        await file.close()

    if total_written == 0:
        saved_path.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail="Tệp tải lên trống (0 byte).")

    file_type = _ext_to_file_type(ext)
    inspect_pdf, _, inspect_pptx, _, _ = _get_core_functions()

    try:
        if file_type == "pdf":
            meta = inspect_pdf(str(saved_path))
            page_count = int(meta.get("page_count", 0))
            slide_count = page_count
        else:
            meta = inspect_pptx(str(saved_path))
            slide_count = int(meta.get("slide_count", 0))
            page_count = slide_count
    except ValueError as exc:
        saved_path.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        saved_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"Không thể phân tích tệp: {exc}")

    is_legacy_ppt = bool(file_type == "ppt" or meta.get("is_legacy_ppt"))
    media_count = int(meta.get("media_count", 0) or 0)

    file_record = {
        "file_id": file_id,
        "filename": original_filename,
        "ext": ext,
        "file_type": file_type,
        "is_legacy_ppt": is_legacy_ppt,
        "path": str(saved_path),
        "original_size_bytes": meta.get("original_size_bytes", total_written),
        "original_size_mb": round(
            float(meta.get("original_size_mb", total_written / (1024 * 1024))),
            2,
        ),
        "page_count": page_count,
        "slide_count": slide_count,
        "media_count": media_count,
        "width": meta.get("width"),
        "height": meta.get("height"),
        "preview_base64": meta.get("preview_base64"),
        "created_at": time.time(),
    }
    FILES[file_id] = file_record

    return {
        "file_id": file_id,
        "filename": original_filename,
        "file_type": file_type,
        "is_legacy_ppt": is_legacy_ppt,
        "original_size_bytes": file_record["original_size_bytes"],
        "original_size_mb": file_record["original_size_mb"],
        "page_count": page_count,
        "slide_count": slide_count,
        "media_count": media_count,
        "width": file_record["width"],
        "height": file_record["height"],
        "preview_base64": file_record["preview_base64"],
    }


def _run_compression_worker(
    job_id: str,
    file_info: Dict[str, Any],
    target_mb: float,
    mode: str,
    output_format: str,
):
    job = JOBS[job_id]
    try:
        (
            _,
            compress_pdf,
            _,
            compress_pptx_direct,
            convert_pptx_to_compressed_pdf,
        ) = _get_core_functions()

        input_path = file_info["path"]
        file_type = file_info["file_type"]
        orig_stem = Path(file_info["filename"]).stem

        def progress_cb(current: int, total: int, message: str = ""):
            job["progress_current"] = int(current)
            job["progress_total"] = max(int(total), 1)
            job["message"] = (
                message
                if message
                else f"Đang xử lý {current}/{total}..."
            )

        if file_type == "pdf":
            final_ext = ".pdf"
            output_path = str(OUTPUTS_DIR / f"{job_id}_compressed{final_ext}")
            result = compress_pdf(
                input_path=input_path,
                output_path=output_path,
                target_mb=target_mb,
                mode=mode,
                progress_callback=progress_cb,
            )
        else:
            if output_format == "pptx":
                final_ext = ".pptx"
                output_path = str(OUTPUTS_DIR / f"{job_id}_compressed{final_ext}")
                result = compress_pptx_direct(
                    input_path=input_path,
                    output_path=output_path,
                    target_mb=target_mb,
                    progress_callback=progress_cb,
                )
            else:
                final_ext = ".pdf"
                output_path = str(OUTPUTS_DIR / f"{job_id}_compressed{final_ext}")
                result = convert_pptx_to_compressed_pdf(
                    input_path=input_path,
                    output_path=output_path,
                    target_mb=target_mb,
                    mode=mode,
                    progress_callback=progress_cb,
                )

        download_filename = f"{orig_stem}_compressed{final_ext}"
        download_url = f"/api/download/{job_id}"

        comp_bytes = result.get("compressed_size_bytes")
        if comp_bytes is None and Path(output_path).exists():
            comp_bytes = Path(output_path).stat().st_size

        target_met = bool(
            result.get(
                "target_met",
                (comp_bytes <= int(target_mb * 1024 * 1024))
                if comp_bytes is not None
                else True,
            )
        )

        orig_mb = round(
            float(
                result.get(
                    "original_size_mb",
                    file_info.get("original_size_mb", 0.0),
                )
            ),
            2,
        )
        comp_mb = round(float(result.get("compressed_size_mb", 0.0)), 2)
        reduction_pct = round(
            float(
                result.get(
                    "reduction_percent",
                    ((orig_mb - comp_mb) / orig_mb * 100.0) if orig_mb > 0 else 0.0,
                )
            ),
            1,
        )

        before_preview = result.get("before_preview_base64") or file_info.get(
            "preview_base64"
        )
        after_preview = result.get("after_preview_base64")

        full_result = {
            **result,
            "job_id": job_id,
            "job_type": "compress",
            "file_id": file_info["file_id"],
            "filename": file_info["filename"],
            "output_filename": download_filename,
            "output_format": final_ext.lstrip("."),
            "target_mb": target_mb,
            "target_met": target_met,
            "original_size_mb": orig_mb,
            "compressed_size_mb": comp_mb,
            "reduction_percent": reduction_pct,
            "before_preview_base64": before_preview,
            "after_preview_base64": after_preview,
            "download_url": download_url,
        }

        total_done = (
            result.get("page_count")
            or result.get("slide_count")
            or job.get("progress_total")
            or 1
        )
        done_msg = (
            "Hoàn tất nén tệp thành công!"
            if target_met
            else f"Đã nén tới ngưỡng an toàn ({comp_mb} MB), cao hơn mục tiêu {target_mb} MB để giữ độ sắc nét."
        )
        job.update(
            {
                "status": "done",
                "job_type": "compress",
                "target_mb": target_mb,
                "target_met": target_met,
                "progress_current": total_done,
                "progress_total": total_done,
                "message": done_msg,
                "output_path": output_path,
                "output_filename": download_filename,
                "original_size_mb": orig_mb,
                "compressed_size_mb": comp_mb,
                "reduction_percent": reduction_pct,
                "before_preview_base64": before_preview,
                "after_preview_base64": after_preview,
                "download_url": download_url,
                "result": full_result,
            }
        )
    except Exception as exc:
        traceback.print_exc()
        job.update(
            {
                "status": "error",
                "message": f"Lỗi khi nén tệp: {exc}",
                "error": str(exc),
            }
        )


def _run_extraction_worker(
    job_id: str,
    file_info: Dict[str, Any],
    extract_mode: str,
):
    job = JOBS[job_id]
    try:
        extract_archive = _get_extract_function()

        input_path = file_info["path"]
        file_type = file_info["file_type"]
        orig_stem = Path(file_info["filename"]).stem
        output_path = str(OUTPUTS_DIR / f"{job_id}_extracted.zip")

        def progress_cb(current: int, total: int, message: str = ""):
            job["progress_current"] = int(current)
            job["progress_total"] = max(int(total), 1)
            job["message"] = (
                message
                if message
                else f"Đang trích xuất {current}/{total} mục..."
            )

        result = extract_archive(
            input_path=input_path,
            output_zip_path=output_path,
            file_type=file_type,
            extract_mode=extract_mode,
            progress_callback=progress_cb,
        )

        output_filename = f"{orig_stem}_{extract_mode}.zip"
        download_url = f"/api/download/{job_id}"

        extracted_count = int(result.get("extracted_count", 0))
        zip_size_bytes = int(
            result.get(
                "zip_size_bytes",
                Path(output_path).stat().st_size if Path(output_path).exists() else 0,
            )
        )
        raw_zip_mb = float(
            result.get("zip_size_mb", zip_size_bytes / (1024 * 1024))
        )
        zip_size_mb = round(raw_zip_mb, 4)
        compressed_size_mb = round(raw_zip_mb, 2)
        orig_mb = round(float(file_info.get("original_size_mb", 0.0)), 2)
        preview_b64 = result.get("preview_base64") or file_info.get("preview_base64")

        full_result = {
            **result,
            "job_id": job_id,
            "job_type": "extract",
            "file_id": file_info["file_id"],
            "filename": file_info["filename"],
            "extract_mode": extract_mode,
            "extracted_count": extracted_count,
            "zip_size_bytes": zip_size_bytes,
            "zip_size_mb": zip_size_mb,
            "original_size_mb": orig_mb,
            "compressed_size_mb": compressed_size_mb,
            "output_filename": output_filename,
            "output_format": "zip",
            "preview_base64": preview_b64,
            "download_url": download_url,
        }

        total_done = max(extracted_count, int(job.get("progress_total") or 1))
        job.update(
            {
                "status": "done",
                "job_type": "extract",
                "extract_mode": extract_mode,
                "extracted_count": extracted_count,
                "zip_size_bytes": zip_size_bytes,
                "zip_size_mb": zip_size_mb,
                "original_size_mb": orig_mb,
                "compressed_size_mb": compressed_size_mb,
                "progress_current": total_done,
                "progress_total": total_done,
                "message": f"Hoàn tất giải nén {extracted_count} mục vào gói .ZIP!",
                "output_path": output_path,
                "output_filename": output_filename,
                "output_format": "zip",
                "preview_base64": preview_b64,
                "download_url": download_url,
                "result": full_result,
            }
        )
    except Exception as exc:
        traceback.print_exc()
        job.update(
            {
                "status": "error",
                "message": f"Lỗi khi giải nén tài nguyên: {exc}",
                "error": str(exc),
            }
        )


@app.post("/api/compress")
async def start_compression(request: Request):
    payload = await _parse_request_payload(request)

    file_id = payload.get("file_id")
    if not file_id:
        raise HTTPException(status_code=400, detail="Thiếu tham số 'file_id'.")

    file_info = _resolve_file_info(str(file_id))

    try:
        target_mb = float(payload.get("target_mb", 10.0))
        if target_mb <= 0:
            target_mb = 10.0
    except (TypeError, ValueError):
        target_mb = 10.0

    mode = str(payload.get("mode", "flatten")).lower().strip()
    if mode not in ("flatten", "vector"):
        mode = "flatten"

    raw_output_format = str(payload.get("output_format", "pdf")).lower().strip()
    if (
        file_info["file_type"] == "ppt" or file_info.get("is_legacy_ppt")
    ) and raw_output_format == "pptx":
        raise HTTPException(
            status_code=400,
            detail="Định dạng .ppt cũ không hỗ trợ nén trực tiếp sang .pptx. Vui lòng chọn xuất sang PDF.",
        )

    output_format = raw_output_format if raw_output_format in ("pdf", "pptx") else "pdf"
    if file_info["file_type"] in ("pdf", "ppt"):
        output_format = "pdf"

    job_id = uuid.uuid4().hex[:12]
    initial_total = int(
        file_info.get("page_count") or file_info.get("slide_count") or 1
    )

    JOBS[job_id] = {
        "job_id": job_id,
        "job_type": "compress",
        "file_id": file_info["file_id"],
        "filename": file_info["filename"],
        "target_mb": target_mb,
        "target_met": None,
        "mode": mode,
        "output_format": output_format,
        "status": "running",
        "progress_current": 0,
        "progress_total": max(initial_total, 1),
        "message": "Đang khởi tạo tiến trình nén...",
        "result": None,
        "error": None,
        "created_at": time.time(),
    }

    worker = threading.Thread(
        target=_run_compression_worker,
        args=(job_id, file_info, target_mb, mode, output_format),
        daemon=True,
    )
    worker.start()

    return {
        "job_id": job_id,
        "job_type": "compress",
        "file_id": file_info["file_id"],
        "target_mb": target_mb,
        "status": "running",
        "progress_current": 0,
        "progress_total": max(initial_total, 1),
        "message": "Đang khởi tạo tiến trình nén...",
    }


@app.post("/api/extract")
async def start_extraction(request: Request):
    payload = await _parse_request_payload(request)

    file_id = payload.get("file_id")
    if not file_id:
        raise HTTPException(status_code=400, detail="Thiếu tham số 'file_id'.")

    file_info = _resolve_file_info(str(file_id))

    extract_mode = str(payload.get("extract_mode", "images")).lower().strip()
    if extract_mode not in ("images", "slides"):
        raise HTTPException(
            status_code=400,
            detail="Chế độ giải nén không hợp lệ. Chỉ hỗ trợ 'images' hoặc 'slides'.",
        )

    if (
        file_info["file_type"] == "ppt" or file_info.get("is_legacy_ppt")
    ) and extract_mode == "images":
        raise HTTPException(
            status_code=400,
            detail="Định dạng .ppt cũ chỉ hỗ trợ tách slide thành bộ ảnh. Vui lòng chọn chế độ tách slide hoặc lưu dưới dạng .pptx.",
        )

    job_id = uuid.uuid4().hex[:12]
    if extract_mode == "images" and file_info.get("media_count"):
        initial_total = max(int(file_info["media_count"]), 1)
    else:
        initial_total = max(
            int(file_info.get("page_count") or file_info.get("slide_count") or 1),
            1,
        )

    JOBS[job_id] = {
        "job_id": job_id,
        "job_type": "extract",
        "file_id": file_info["file_id"],
        "filename": file_info["filename"],
        "extract_mode": extract_mode,
        "output_format": "zip",
        "status": "running",
        "progress_current": 0,
        "progress_total": initial_total,
        "message": "Đang khởi tạo tiến trình giải nén tài nguyên...",
        "result": None,
        "error": None,
        "created_at": time.time(),
    }

    worker = threading.Thread(
        target=_run_extraction_worker,
        args=(job_id, file_info, extract_mode),
        daemon=True,
    )
    worker.start()

    return {
        "job_id": job_id,
        "job_type": "extract",
        "file_id": file_info["file_id"],
        "extract_mode": extract_mode,
        "status": "running",
        "progress_current": 0,
        "progress_total": initial_total,
        "message": "Đang khởi tạo tiến trình giải nén tài nguyên...",
    }


@app.get("/api/job/{job_id}")
async def get_job_status(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy job_id: {job_id}")

    job = JOBS[job_id]
    job_type = job.get("job_type", "compress")

    response: Dict[str, Any] = {
        "job_id": job["job_id"],
        "job_type": job_type,
        "file_id": job["file_id"],
        "status": job["status"],
        "progress_current": job.get("progress_current", 0),
        "progress_total": job.get("progress_total", 1),
        "message": job.get("message", ""),
        "output_format": job.get("output_format"),
        "result": job.get("result"),
    }

    if job_type == "compress":
        response["target_mb"] = job.get("target_mb")
        response["target_met"] = job.get("target_met")
    else:
        response["extract_mode"] = job.get("extract_mode")

    if job["status"] == "done" and job.get("result"):
        res = job["result"]
        if job_type == "extract":
            response.update(
                {
                    "extract_mode": res.get("extract_mode"),
                    "extracted_count": res.get("extracted_count", 0),
                    "zip_size_bytes": res.get("zip_size_bytes", 0),
                    "zip_size_mb": res.get("zip_size_mb", 0.0),
                    "original_size_mb": res.get("original_size_mb"),
                    "compressed_size_mb": res.get("compressed_size_mb"),
                    "preview_base64": res.get("preview_base64"),
                    "download_url": res.get("download_url", f"/api/download/{job_id}"),
                    "output_filename": res.get("output_filename"),
                    "output_format": "zip",
                }
            )
        else:
            response.update(
                {
                    "target_mb": res.get("target_mb", job.get("target_mb")),
                    "target_met": bool(res.get("target_met", True)),
                    "original_size_mb": res.get("original_size_mb"),
                    "compressed_size_mb": res.get("compressed_size_mb"),
                    "reduction_percent": res.get("reduction_percent"),
                    "before_preview_base64": res.get("before_preview_base64"),
                    "after_preview_base64": res.get("after_preview_base64"),
                    "download_url": res.get("download_url", f"/api/download/{job_id}"),
                    "output_filename": res.get("output_filename"),
                    "output_format": res.get("output_format", job.get("output_format")),
                    "page_count": res.get("page_count") or res.get("slide_count"),
                    "mode_used": res.get("mode_used"),
                }
            )
    elif job["status"] == "error":
        response["error"] = job.get("error", job.get("message"))

    return response


@app.get("/api/download/{job_id}")
async def download_compressed_file(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy job_id: {job_id}")

    job = JOBS[job_id]
    if job["status"] != "done":
        raise HTTPException(
            status_code=400,
            detail=f"Tệp chưa sẵn sàng để tải xuống (trạng thái: {job['status']}).",
        )

    output_path = job.get("output_path")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(
            status_code=404, detail="Tệp đầu ra không tồn tại trên máy chủ."
        )

    output_filename = job.get("output_filename") or Path(output_path).name
    ext = Path(output_filename).suffix.lower()
    if ext == ".pdf":
        media_type = "application/pdf"
    elif ext == ".pptx":
        media_type = (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    elif ext == ".zip":
        media_type = "application/zip"
    else:
        media_type = "application/octet-stream"

    encoded_filename = quote(output_filename)
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
    }

    return FileResponse(
        path=str(output_path),
        media_type=media_type,
        filename=output_filename,
        headers=headers,
    )
