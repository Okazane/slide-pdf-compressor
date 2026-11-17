import os
import sys
import threading
import traceback
import uuid
from pathlib import Path
from typing import Any, Dict, Optional
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
    description="Hệ thống nén PDF & PowerPoint thông minh giữ nguyên độ nét",
    version="1.0.0",
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


def _resolve_file_info(file_id: str) -> Dict[str, Any]:
    if file_id in FILES:
        info = FILES[file_id]
        if Path(info["path"]).exists():
            return info

    # Fallback: look up in UPLOADS_DIR by prefix
    for candidate in UPLOADS_DIR.glob(f"{file_id}*"):
        if candidate.is_file() and candidate.suffix.lower() in ALLOWED_EXTENSIONS:
            ext = candidate.suffix.lower()
            file_type = "pdf" if ext == ".pdf" else "pptx"
            info = {
                "file_id": file_id,
                "filename": candidate.name,
                "ext": ext,
                "file_type": file_type,
                "path": str(candidate),
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

    try:
        content = await file.read()
        with open(saved_path, "wb") as f:
            f.write(content)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Lỗi lưu tệp tải lên: {exc}")

    file_type = "pdf" if ext == ".pdf" else "pptx"
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
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Không thể phân tích tệp: {exc}")

    file_record = {
        "file_id": file_id,
        "filename": original_filename,
        "ext": ext,
        "file_type": file_type,
        "path": str(saved_path),
        "original_size_bytes": meta.get("original_size_bytes", saved_path.stat().st_size),
        "original_size_mb": round(
            float(meta.get("original_size_mb", saved_path.stat().st_size / (1024 * 1024))),
            2,
        ),
        "page_count": page_count,
        "slide_count": slide_count,
        "media_count": meta.get("media_count", 0),
        "width": meta.get("width"),
        "height": meta.get("height"),
        "preview_base64": meta.get("preview_base64"),
    }
    FILES[file_id] = file_record

    return {
        "file_id": file_id,
        "filename": original_filename,
        "file_type": file_type,
        "original_size_bytes": file_record["original_size_bytes"],
        "original_size_mb": file_record["original_size_mb"],
        "page_count": page_count,
        "slide_count": slide_count,
        "media_count": file_record["media_count"],
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
                else f"Đang xử lý trang {current}/{total}..."
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
            "file_id": file_info["file_id"],
            "filename": file_info["filename"],
            "output_filename": download_filename,
            "output_format": final_ext.lstrip("."),
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
        job.update(
            {
                "status": "done",
                "progress_current": total_done,
                "progress_total": total_done,
                "message": "Hoàn tất nén tệp thành công!",
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


@app.post("/api/compress")
async def start_compression(request: Request):
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

    default_out = "pdf"
    output_format = str(payload.get("output_format", default_out)).lower().strip()
    if output_format not in ("pdf", "pptx"):
        output_format = "pdf"
    if file_info["file_type"] == "pdf":
        output_format = "pdf"

    job_id = uuid.uuid4().hex[:12]
    initial_total = int(
        file_info.get("page_count") or file_info.get("slide_count") or 1
    )

    JOBS[job_id] = {
        "job_id": job_id,
        "file_id": file_info["file_id"],
        "filename": file_info["filename"],
        "target_mb": target_mb,
        "mode": mode,
        "output_format": output_format,
        "status": "running",
        "progress_current": 0,
        "progress_total": max(initial_total, 1),
        "message": "Đang khởi tạo tiến trình nén...",
        "result": None,
        "error": None,
    }

    worker = threading.Thread(
        target=_run_compression_worker,
        args=(job_id, file_info, target_mb, mode, output_format),
        daemon=True,
    )
    worker.start()

    return {
        "job_id": job_id,
        "file_id": file_info["file_id"],
        "status": "running",
        "progress_current": 0,
        "progress_total": max(initial_total, 1),
        "message": "Đang khởi tạo tiến trình nén...",
    }


@app.get("/api/job/{job_id}")
async def get_job_status(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy job_id: {job_id}")

    job = JOBS[job_id]
    response: Dict[str, Any] = {
        "job_id": job["job_id"],
        "file_id": job["file_id"],
        "status": job["status"],
        "progress_current": job.get("progress_current", 0),
        "progress_total": job.get("progress_total", 1),
        "message": job.get("message", ""),
        "result": job.get("result"),
    }

    if job["status"] == "done" and job.get("result"):
        res = job["result"]
        response.update(
            {
                "original_size_mb": res.get("original_size_mb"),
                "compressed_size_mb": res.get("compressed_size_mb"),
                "reduction_percent": res.get("reduction_percent"),
                "before_preview_base64": res.get("before_preview_base64"),
                "after_preview_base64": res.get("after_preview_base64"),
                "download_url": res.get("download_url", f"/api/download/{job_id}"),
                "output_filename": res.get("output_filename"),
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
        raise HTTPException(status_code=404, detail="Tệp đầu ra không tồn tại trên máy chủ.")

    output_filename = job.get("output_filename") or Path(output_path).name
    ext = Path(output_filename).suffix.lower()
    if ext == ".pdf":
        media_type = "application/pdf"
    elif ext == ".pptx":
        media_type = (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
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
