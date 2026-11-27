import json
import logging
from pathlib import Path
from typing import List, Optional
from fastapi import APIRouter, Depends, Query, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from config.settings import CACHE_DIR
from core.auth import verify_api_key
from core.jobs import (
    create_report_job,
    create_upload_report_job,
    create_l4d_baseline_job,
    get_job,
    active_jobs,
)
from core.logger import get_job_log_lines

router = APIRouter()
log = logging.getLogger("conversion_report_server.routes")


class L4DBaselineRequest(BaseModel):
    file_ids: List[str] = Field(..., min_length=1, description="List of Google Drive file IDs or URLs")
    access_token: Optional[str] = Field(None, description="Optional Google OAuth access token")


def _default_conversion_tabs(sub_type: Optional[str]) -> List[str]:
    if sub_type and str(sub_type).strip():
        raw = str(sub_type).strip()
        clean = raw.lower()
        if clean in ("d1", "d-1", "d_1", "nextday"):
            prefix = "D-1"
        elif clean == "sameday":
            prefix = "Sameday"
        else:
            prefix = raw
        return [f"{prefix} DC_View", f"{prefix} Agent_View"]
    return ["DC_View", "Agent_View"]


@router.get("/")
async def root():
    return {
        "server": "Conversion Report Stream Server",
        "status": "ready",
        "version": "1.0.0",
        "engine": "In-Flight Rust Calamine / XML Streaming",
    }


@router.get("/health")
async def health(x_api_key: Optional[str] = Depends(verify_api_key)):
    active_count = sum(1 for j in active_jobs.values() if j.get("status") == "processing")
    return {
        "status": "ok",
        "active_jobs": active_count,
        "total_tracked_jobs": len(active_jobs),
        "cache_dir": str(CACHE_DIR),
    }


@router.get("/convert-async")
async def convert_async(
    drive_url: str = Query(..., description="Google Drive URL or File ID"),
    sub_type: Optional[str] = Query(None, description="Report sub-type (e.g. sameday, d1)"),
    access_token: Optional[str] = Query(None, description="Optional Google OAuth access token"),
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    if access_token and "access_token=" not in drive_url:
        sep = "&" if "?" in drive_url else "?"
        drive_url = f"{drive_url}{sep}access_token={access_token}"
    return create_report_job(drive_url, sub_type=sub_type)


@router.post("/convert-upload")
@router.post("/reports/conversion")
@router.post("/reports/async/conversion")
async def convert_upload(
    file: UploadFile = File(...),
    sub_type: Optional[str] = Query(None, description="Report sub-type (e.g. sameday, d1)"),
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    content = await file.read()
    return create_upload_report_job(content, file.filename or "upload.xlsx", sub_type=sub_type)


@router.get("/job/{job_id}")
@router.get("/jobs/{job_id}")
async def get_job_status_route(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    job = get_job(job_id)
    logs = get_job_log_lines(job_id)
    tabs = job.get("tabs")
    if not tabs:
        tabs = _default_conversion_tabs(job.get("sub_type"))

    return {
        "job_id": job_id,
        "status": job.get("status"),
        "progress": job.get("progress"),
        "error": job.get("error"),
        "report_type": job.get("report_type", "conversion"),
        "sub_type": job.get("sub_type"),
        "report_date": job.get("report_date"),
        "file_name": job.get("file_name"),
        "tabs": tabs,
        "created_at": job.get("created_at"),
        "completed_at": job.get("completed_at"),
        "logs": logs,
    }


@router.get("/job/{job_id}/logs")
@router.get("/jobs/{job_id}/logs")
async def get_job_logs_route(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    logs = get_job_log_lines(job_id)
    return {
        "job_id": job_id,
        "logs": logs,
    }


@router.get("/job/{job_id}/result")
@router.get("/jobs/{job_id}/result")
@router.get("/reports/download/{job_id}")
async def get_job_result(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    job = get_job(job_id)
    if job.get("status") != "done":
        raise HTTPException(
            status_code=400,
            detail=f"Job is not ready yet (status: {job.get('status')})",
        )

    out_path = Path(job.get("output_path", ""))
    if not out_path.exists():
        raise HTTPException(
            status_code=404,
            detail="Generated report file has expired or was removed from cache.",
        )

    if out_path.stat().st_size < 1000:
        raise HTTPException(
            status_code=500,
            detail=f"Generated report file is incomplete or empty ({out_path.stat().st_size} bytes).",
        )

    download_name = job.get("file_name") or f"Conversion_Summary_Report_{job_id}.xlsx"
    return FileResponse(
        path=str(out_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name,
    )


@router.post("/l4d-historical-baseline")
async def l4d_historical_baseline(
    payload: L4DBaselineRequest,
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    return create_l4d_baseline_job(payload.file_ids, access_token=payload.access_token)


@router.get("/l4d-job/{job_id}/result")
async def get_l4d_job_result(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key),
):
    job = get_job(job_id)
    if job.get("status") != "done":
        raise HTTPException(
            status_code=400,
            detail=f"Job is not ready yet (status: {job.get('status')})",
        )

    out_path = Path(job.get("output_path") or (CACHE_DIR / f"L4D_BASELINE_{job_id}.json"))
    if not out_path.exists():
        raise HTTPException(
            status_code=404,
            detail="Generated L4D baseline JSON file has expired or was removed from cache.",
        )

    with open(out_path, "r", encoding="utf-8") as f:
        return json.load(f)


@router.post("/clear-cache")
@router.get("/clear-cache")
async def clear_cache_endpoint(x_api_key: Optional[str] = Depends(verify_api_key)):
    """Wipes all cached jobs and temporary report files from disk."""
    patterns = [
        "JOB_*.json",
        "L4D_BASELINE_*.json",
        "REPORT_*.xlsx",
        "*.xlsx",
        "*.xlsb",
        "*.ods",
        "*.csv",
        "*.tsv",
        "*.xls",
        "*.xlsm",
    ]
    deleted = 0
    for pattern in patterns:
        for f in CACHE_DIR.glob(pattern):
            try:
                f.unlink(missing_ok=True)
                deleted += 1
            except Exception:
                pass
    active_jobs.clear()
    log.info(f"Full cache wipe executed: deleted {deleted} files.")
    return {"status": "ok", "deleted_files": deleted, "cache_dir": str(CACHE_DIR)}
