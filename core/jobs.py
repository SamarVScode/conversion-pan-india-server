import uuid
import time
import json
import hashlib
import logging
import threading
import re
import gc
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional
from fastapi import HTTPException

from config.settings import (
    CACHE_DIR,
    CACHE_TTL,
    CACHE_MAX_AGE,
    MAX_CONCURRENT_JOBS,
    MAX_QUEUED_JOBS,
)
from core.downloader import extract_file_id, download_from_url
from core.logger import (
    print_job_start,
    print_job_step,
    print_job_success,
    print_job_error,
    log_job_message,
)
from generators import (
    generate_conversion_report,
    extract_report_date_from_agent_view,
    aggregate_l4d_baseline,
)

log = logging.getLogger("conversion_report_server.jobs")

active_jobs: Dict[str, Dict[str, Any]] = {}
conversion_semaphore = threading.Semaphore(MAX_CONCURRENT_JOBS)
_jobs_lock = threading.Lock()


def _job_meta_path(job_id: str) -> Path:
    return CACHE_DIR / f"JOB_{job_id}.json"


def _save_job(job_id: str, job: Dict[str, Any]) -> None:
    try:
        meta_path = _job_meta_path(job_id)
        with open(meta_path, "w") as f:
            json.dump({"job_id": job_id, **job}, f)
    except Exception as e:
        log.warning(f"Could not persist job {job_id} to disk: {e}")


def _set_job(job_id: str, job: Dict[str, Any]) -> None:
    with _jobs_lock:
        active_jobs[job_id] = job
    _save_job(job_id, job)


def recover_jobs_from_disk() -> None:
    recovered = 0
    now = time.time()
    try:
        for meta_file in CACHE_DIR.glob("JOB_*.json"):
            try:
                with open(meta_file, "r") as f:
                    data = json.load(f)
                jid = data.get("job_id")
                created_at = data.get("created_at", 0)
                if not jid:
                    continue

                if now - created_at > CACHE_MAX_AGE:
                    meta_file.unlink(missing_ok=True)
                    out_p = data.get("output_path")
                    if out_p:
                        Path(out_p).unlink(missing_ok=True)
                    continue

                if data.get("status") == "processing":
                    data["status"] = "error"
                    data["error"] = "Job interrupted by server restart."
                    data["progress"] = "Failed (server restarted)"
                    _save_job(jid, data)

                active_jobs[jid] = data
                recovered += 1
            except Exception as e:
                log.warning(f"Could not recover {meta_file.name}: {e}")
        if recovered:
            log.info(f"Recovered {recovered} jobs from disk on boot.")
    except Exception as e:
        log.warning(f"Error during job recovery: {e}")


def get_job(job_id: str) -> Dict[str, Any]:
    if job_id in active_jobs:
        return active_jobs[job_id]

    meta_file = _job_meta_path(job_id)
    if meta_file.exists():
        try:
            with open(meta_file, "r") as f:
                data = json.load(f)
            active_jobs[job_id] = data
            return data
        except Exception:
            pass

    raise HTTPException(status_code=404, detail="Job not found")


def cleanup_old_jobs() -> None:
    evict_old_jobs()


def evict_old_jobs() -> None:
    now = time.time()
    to_delete = []
    for jid, job in list(active_jobs.items()):
        created_at = job.get("created_at", 0)
        if now - created_at > CACHE_MAX_AGE:
            to_delete.append(jid)
    for jid in to_delete:
        meta_file = _job_meta_path(jid)
        meta_file.unlink(missing_ok=True)
        job = active_jobs.get(jid, {})
        out = job.get("output_path")
        if out:
            Path(out).unlink(missing_ok=True)
        del active_jobs[jid]
        log.info(f"Evicted expired job {jid} from cache")


def generate_proper_report_filename(
    original_filename: Optional[str] = None,
    sub_type: Optional[str] = None,
    report_date: Optional[str] = None,
) -> str:
    if sub_type and str(sub_type).strip():
        clean_sub = str(sub_type).strip().lower()
        if clean_sub in ("d1", "d-1", "d_1", "nextday"):
            prefix = "Conversion_Summary_D-1_Report"
        elif clean_sub == "sameday":
            prefix = "Conversion_Summary_Sameday_Report"
        else:
            prefix = f"Conversion_Summary_{str(sub_type).strip().replace(' ', '_')}_Report"
    else:
        prefix = "Conversion_Summary_Report"

    gen_date_str = datetime.now().strftime("%d-%b-%Y")
    if report_date and str(report_date).strip():
        clean_report_date = str(report_date).strip().replace("/", "-").replace("_", "-")
        return f"{prefix}_{clean_report_date}_{gen_date_str}.xlsx"
    elif original_filename:
        date_match = re.search(
            r"(\d{1,2}[-_][a-zA-Z]{3}[-_]\d{4}|\d{4}[-_]\d{2}[-_]\d{2}|\d{1,2}[-_]\d{1,2}[-_]\d{4})",
            original_filename,
        )
        if date_match:
            extracted_date = date_match.group(1).replace("_", "-")
            return f"{prefix}_{extracted_date}_{gen_date_str}.xlsx"
    return f"{prefix}_{gen_date_str}.xlsx"


def _cleanup_upload_inputs(file_id: str) -> None:
    if not str(file_id).startswith("upload_"):
        return
    for path in CACHE_DIR.glob(f"{file_id}.*"):
        if path.suffix.lower() not in (".json", ".log") and not path.name.startswith("REPORT_"):
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass


def ensure_queue_capacity() -> None:
    """Rejects new heavy jobs with HTTP 429 once in-flight + queued jobs hit MAX_CONCURRENT_JOBS + MAX_QUEUED_JOBS."""
    max_total = max(1, MAX_CONCURRENT_JOBS + max(0, MAX_QUEUED_JOBS))
    with _jobs_lock:
        in_flight = sum(1 for j in active_jobs.values() if j.get("status") in ("processing", "queued"))
    if in_flight >= max_total:
        raise HTTPException(
            status_code=429,
            detail=(
                f"Server busy: {in_flight} jobs already active/queued (limit {max_total}). "
                "Please retry in a moment."
            ),
        )


def background_report_job(job_id: str, file_id: str, output_path: Path) -> None:
    log_job_message(job_id, f"Waiting for execution lock (Max concurrent: {MAX_CONCURRENT_JOBS})...")
    acquired = conversion_semaphore.acquire(timeout=600)
    if not acquired:
        _cleanup_upload_inputs(file_id)
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = "Server busy: concurrent limit reached."
        job["progress"] = "Failed: server busy"
        _set_job(job_id, job)
        print_job_error(job_id, "Server busy: concurrent limit reached")
        return

    if output_path.exists() and output_path.stat().st_size > 0:
        conversion_semaphore.release()
        job = active_jobs.get(job_id, {})
        job["status"] = "done"
        job["progress"] = "Complete (cached)"
        job["completed_at"] = output_path.stat().st_mtime
        _set_job(job_id, job)
        log_job_message(job_id, f"Report already cached ({output_path.name}). Returning instant result.")
        return

    t_start = time.time()
    source_desc = "Direct Upload" if str(file_id).startswith("upload_") else f"Google Drive ({file_id})"
    print_job_start(job_id, "conversion", source_desc)

    tmp_input = None
    try:
        job = active_jobs.get(job_id, {})
        job["status"] = "processing"
        job["progress"] = "Locating input file..."
        _set_job(job_id, job)

        for path in CACHE_DIR.glob(f"{file_id}.*"):
            if path.suffix.lower() not in (".json", ".log") and not path.name.startswith("REPORT_"):
                tmp_input = path
                break

        if tmp_input is None:
            tmp_input = CACHE_DIR / f"{file_id}.xlsx"

        if not file_id.startswith("upload_"):
            needs_download = not tmp_input.exists() or (time.time() - tmp_input.stat().st_mtime > CACHE_TTL)
            if needs_download:
                t0 = time.time()
                raw_url = active_jobs.get(job_id, {}).get("source_url", "")
                download_url = raw_url if raw_url else f"https://drive.google.com/uc?export=download&id={file_id}"
                print_job_step(job_id, 1, f"Downloading source file (ID: {file_id})...")
                download_from_url(download_url, tmp_input, job_id=job_id)
                print_job_step(
                    job_id,
                    1,
                    f"Source file downloaded ({tmp_input.stat().st_size / 1024 / 1024:.2f} MB in {time.time() - t0:.2f}s)",
                )
            else:
                print_job_step(job_id, 1, f"Using cached source file ({tmp_input.name})")

        job["progress"] = "Generating Conversion Stream Report..."
        _set_job(job_id, job)

        t_gen = time.time()
        print_job_step(job_id, 2, "Executing Conversion stream report generator...")

        sub_type = job.get("sub_type") or "sameday"
        generate_conversion_report(tmp_input, output_path, sub_type=sub_type)

        try:
            report_date = extract_report_date_from_agent_view(tmp_input)
            if not report_date and output_path.exists():
                report_date = extract_report_date_from_agent_view(output_path)
        except Exception as ex:
            log.warning(f"Could not extract report date from Agent_view: {ex}")
            report_date = None

        if report_date:
            job["report_date"] = report_date
            job["file_name"] = generate_proper_report_filename(
                original_filename=job.get("original_filename"),
                sub_type=job.get("sub_type"),
                report_date=report_date,
            )

        if not output_path.exists() or output_path.stat().st_size < 100:
            raise RuntimeError(
                f"Report generator failed to produce a valid file ({output_path.name}). Output size is 0 bytes."
            )

        out_size_kb = output_path.stat().st_size / 1024
        print_job_step(job_id, 2, f"Stream report generated ({out_size_kb:.1f} KB in {time.time() - t_gen:.2f}s)")

        tabs = []
        try:
            with zipfile.ZipFile(str(output_path), "r") as z:
                tree = ET.fromstring(z.read("xl/workbook.xml"))
                tabs = [elem.attrib.get("name") for elem in tree.findall(".//{*}sheet") if elem.attrib.get("name")]
            job["tabs"] = tabs
            log_job_message(job_id, f"Output sheets created: {tabs}")
        except Exception as tab_err:
            log.warning(f"Could not inspect output tabs: {tab_err}")

        if tmp_input and tmp_input.exists():
            try:
                tmp_input.unlink()
                log_job_message(job_id, f"Disk Reclaimed: Temporary raw input file deleted ({tmp_input.name})")
            except Exception as del_err:
                log.warning(f"Could not delete temporary input file {tmp_input.name}: {del_err}")

        job["status"] = "done"
        job["progress"] = "Complete"
        job["completed_at"] = time.time()
        _set_job(job_id, job)

        duration = time.time() - t_start
        print_job_success(job_id, duration, f"{output_path.name} ({out_size_kb:.1f} KB, tabs: {tabs})")

    except Exception as e:
        log.exception(f"Job {job_id} failed: {e}")
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = str(e)
        job["progress"] = f"Failed: {str(e)[:100]}"
        _set_job(job_id, job)
        print_job_error(job_id, str(e))

        if tmp_input and tmp_input.exists():
            try:
                tmp_input.unlink()
            except Exception:
                pass
    finally:
        gc.collect()
        conversion_semaphore.release()


def process_job_async(job_id: str, file_id: str, output_path: Path) -> None:
    thread = threading.Thread(
        target=background_report_job,
        args=(job_id, file_id, output_path),
        daemon=True,
    )
    thread.start()


def create_report_job(drive_url: str, sub_type: Optional[str] = None) -> Dict[str, str]:
    evict_old_jobs()
    file_id = extract_file_id(drive_url)
    cache_key = f"{file_id}_conversion_{sub_type}" if sub_type else f"{file_id}_conversion"
    output_path = CACHE_DIR / f"REPORT_{cache_key}.xlsx"
    formatted_name = generate_proper_report_filename(sub_type=sub_type)

    if output_path.exists() and (time.time() - output_path.stat().st_mtime <= CACHE_TTL):
        log_job_message(cache_key, "Instant cache hit for conversion report.")
        cached_report_date = None
        try:
            cached_report_date = extract_report_date_from_agent_view(output_path)
        except Exception:
            pass
        formatted_name = generate_proper_report_filename(sub_type=sub_type, report_date=cached_report_date)
        active_jobs[cache_key] = {
            "status": "done",
            "output_path": str(output_path),
            "error": None,
            "progress": "Complete (cached)",
            "created_at": output_path.stat().st_mtime,
            "report_type": "conversion",
            "sub_type": sub_type,
            "report_date": cached_report_date,
            "file_name": formatted_name,
        }
        return {"job_id": cache_key, "status": "done"}

    job_id = cache_key
    with _jobs_lock:
        existing = active_jobs.get(job_id)
        if existing and existing.get("status") in ("processing", "queued"):
            log_job_message(job_id, "Coalescing concurrent conversion request onto existing in-flight job.")
            return {"job_id": job_id, "status": existing.get("status", "processing"), "deduplicated": True}

    ensure_queue_capacity()

    job = {
        "status": "processing",
        "output_path": str(output_path),
        "error": None,
        "progress": "Job created, waiting to download...",
        "created_at": time.time(),
        "source_url": drive_url,
        "report_type": "conversion",
        "sub_type": sub_type,
        "file_name": formatted_name,
    }
    _set_job(job_id, job)
    log_job_message(job_id, f"Initialized async job for conversion (source: {drive_url[:80]}...)")

    process_job_async(job_id, file_id, output_path)
    return {"job_id": job_id, "status": "processing"}


def create_upload_report_job(
    file_bytes: Optional[bytes] = None,
    filename: str = "upload.xlsx",
    sub_type: Optional[str] = None,
    staged_path: Optional[Path] = None,
    job_id: Optional[str] = None,
    file_size: Optional[int] = None,
) -> Dict[str, str]:
    evict_old_jobs()
    ensure_queue_capacity()

    job_id = job_id or str(uuid.uuid4())[:8]
    file_id = f"upload_{job_id}"

    if staged_path is not None:
        tmp_input = Path(staged_path)
        size_bytes = file_size if file_size is not None else (tmp_input.stat().st_size if tmp_input.exists() else 0)
    else:
        payload = file_bytes or b""
        size_bytes = len(payload)
        ext = Path(filename).suffix.lower() if filename and Path(filename).suffix else ".xlsx"
        tmp_input = CACHE_DIR / f"{file_id}{ext}"
        with open(tmp_input, "wb") as f:
            f.write(payload)

    log.info(f"Creating upload stream job {job_id} (filename={filename}, size={size_bytes} bytes)")

    output_path = CACHE_DIR / f"REPORT_{job_id}.xlsx"
    formatted_name = generate_proper_report_filename(original_filename=filename, sub_type=sub_type)
    job = {
        "status": "processing",
        "output_path": str(output_path),
        "error": None,
        "progress": "Uploaded, starting generation...",
        "created_at": time.time(),
        "report_type": "conversion",
        "sub_type": sub_type,
        "original_filename": filename,
        "file_name": formatted_name,
    }
    _set_job(job_id, job)
    log_job_message(job_id, f"Upload received: {filename} ({size_bytes} bytes). Starting generator...")

    process_job_async(job_id, file_id, output_path)
    return {"job_id": job_id, "status": "processing"}


def background_l4d_baseline_job(
    job_id: str,
    file_ids: List[str],
    access_token: Optional[str],
    output_path: Path,
) -> None:
    log_job_message(job_id, f"Waiting for execution lock (Max concurrent: {MAX_CONCURRENT_JOBS})...")
    acquired = conversion_semaphore.acquire(timeout=600)
    if not acquired:
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = "Server busy: concurrent limit reached."
        job["progress"] = "Failed: server busy"
        _set_job(job_id, job)
        print_job_error(job_id, "Server busy: concurrent limit reached")
        return

    t_start = time.time()
    print_job_start(job_id, "l4d_baseline", f"{len(file_ids)} file(s)")
    downloaded_paths: List[Path] = []

    try:
        total = len(file_ids)
        for idx, raw_id in enumerate(file_ids, 1):
            fid = extract_file_id(raw_id)
            raw_str = raw_id.strip()
            if raw_str.startswith("http"):
                download_url = raw_str
            else:
                download_url = f"https://drive.google.com/uc?export=download&id={fid}"

            if access_token and "access_token=" not in download_url:
                sep = "&" if "?" in download_url else "?"
                download_url = f"{download_url}{sep}access_token={access_token}"

            tmp_path = CACHE_DIR / f"L4D_RAW_{job_id}_{idx}_{fid}.xlsx"
            job = active_jobs.get(job_id, {})
            job["status"] = "processing"
            job["progress"] = f"Downloading file {idx}/{total}..."
            _set_job(job_id, job)

            print_job_step(job_id, 1, f"Downloading file {idx}/{total} (ID: {fid})...")
            download_from_url(download_url, tmp_path, job_id=job_id)
            downloaded_paths.append(tmp_path)

        job = active_jobs.get(job_id, {})
        job["progress"] = f"Aggregating {len(downloaded_paths)} Agent_view files..."
        _set_job(job_id, job)
        print_job_step(job_id, 2, f"Streaming & aggregating {len(downloaded_paths)} Agent_view files...")

        def _on_file_done(done_idx: int, total_cnt: int, p: Path) -> None:
            log_job_message(job_id, f"Processed file {done_idx}/{total_cnt}: {p.name}")

        result = aggregate_l4d_baseline(
            downloaded_paths,
            delete_after_read=True,
            progress_cb=_on_file_done,
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f)

        job = active_jobs.get(job_id, {})
        job["status"] = "done"
        job["progress"] = "Complete"
        job["total_agents"] = result.get("total_agents", 0)
        job["dates_processed"] = result.get("dates_processed", [])
        job["completed_at"] = time.time()
        _set_job(job_id, job)

        duration = time.time() - t_start
        print_job_success(
            job_id,
            duration,
            f"{output_path.name} (agents={result.get('total_agents', 0)}, dates={len(result.get('dates_processed', []))})",
        )

    except Exception as e:
        log.exception(f"L4D baseline job {job_id} failed: {e}")
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = str(e)
        job["progress"] = f"Failed: {str(e)[:100]}"
        _set_job(job_id, job)
        print_job_error(job_id, str(e))
    finally:
        for p in downloaded_paths:
            if p.exists():
                try:
                    p.unlink(missing_ok=True)
                except Exception:
                    pass
        gc.collect()
        conversion_semaphore.release()


def create_l4d_baseline_job(
    file_ids: List[str],
    access_token: Optional[str] = None,
) -> Dict[str, Any]:
    evict_old_jobs()

    canonical_ids = sorted({extract_file_id(fid) for fid in file_ids if str(fid or "").strip()})
    baseline_key = hashlib.sha256("|".join(canonical_ids).encode("utf-8")).hexdigest()[:12]
    max_total = max(1, MAX_CONCURRENT_JOBS + max(0, MAX_QUEUED_JOBS))

    with _jobs_lock:
        # 1. Single-flight / Coalesce: if same file_ids set is already processing OR done within CACHE_TTL,
        # return the existing job_id without queuing a duplicate job on Render.
        for existing_id, existing_job in active_jobs.items():
            if existing_job.get("report_type") != "l4d_baseline":
                continue
            if existing_job.get("baseline_key") != baseline_key:
                continue

            status = existing_job.get("status")
            if status in ("processing", "queued"):
                log_job_message(existing_id, "Coalescing duplicate L4D baseline request onto active job.")
                return {
                    "job_id": existing_id,
                    "status": status,
                    "file_count": len(file_ids),
                    "deduplicated": True,
                }
            if status == "done":
                out_p = Path(existing_job.get("output_path", ""))
                age = time.time() - existing_job.get("completed_at", existing_job.get("created_at", 0))
                if out_p.exists() and age <= CACHE_TTL:
                    log_job_message(existing_id, "Serving fresh cached L4D baseline job (deduplicated).")
                    return {
                        "job_id": existing_id,
                        "status": "done",
                        "file_count": len(file_ids),
                        "deduplicated": True,
                    }

        in_flight = sum(1 for j in active_jobs.values() if j.get("status") in ("processing", "queued"))
        if in_flight >= max_total:
            raise HTTPException(
                status_code=429,
                detail=(
                    f"Server busy: {in_flight} jobs already active/queued (limit {max_total}). "
                    "Please retry in a moment."
                ),
            )

        job_id = str(uuid.uuid4())[:8]
        output_path = CACHE_DIR / f"L4D_BASELINE_{job_id}.json"

        job = {
            "status": "processing",
            "output_path": str(output_path),
            "error": None,
            "progress": "Job created, waiting to download files...",
            "created_at": time.time(),
            "report_type": "l4d_baseline",
            "baseline_key": baseline_key,
            "file_count": len(file_ids),
            "tabs": ["Agent_view"],
        }
        active_jobs[job_id] = job

    _save_job(job_id, job)
    log_job_message(job_id, f"Initialized L4D baseline job for {len(file_ids)} file(s).")

    thread = threading.Thread(
        target=background_l4d_baseline_job,
        args=(job_id, list(file_ids), access_token, output_path),
        daemon=True,
    )
    thread.start()
    return {"job_id": job_id, "status": "processing", "file_count": len(file_ids)}
