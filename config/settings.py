import os
import tempfile
import logging
from pathlib import Path

log = logging.getLogger("conversion_report_server.config")

# Production API Key for client authentication (X-API-KEY header)
REQUIRED_API_KEY = os.getenv("API_KEY", "OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ")

# Cache Directory (defaults to OS temporary folder / conversion_report_cache)
env_cache = os.getenv("CACHE_DIR")
if env_cache:
    CACHE_DIR = Path(env_cache)
else:
    CACHE_DIR = Path(tempfile.gettempdir()) / "conversion_report_cache"

try:
    CACHE_DIR.mkdir(exist_ok=True, parents=True)
except Exception as e:
    log.warning(f"Could not create CACHE_DIR {CACHE_DIR}: {e}")

# Cache & Cleanup Timers
CACHE_TTL = int(os.getenv("CACHE_TTL", "1800"))          # 30 minutes output retention
CACHE_MAX_AGE = int(os.getenv("CACHE_MAX_AGE", "7200"))  # 2 hours job metadata eviction

# Concurrency Containment: Defaults to 1 for 512MB RAM free instances
MAX_CONCURRENT_JOBS = int(os.getenv("MAX_CONCURRENT_JOBS", "1"))
MAX_QUEUED_JOBS = int(os.getenv("MAX_QUEUED_JOBS", "2"))
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "250"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
UPLOAD_CHUNK_BYTES = 1024 * 1024  # 1 MB streaming chunks

log.info(f"CACHE_DIR: {CACHE_DIR}")
log.info(f"MAX_CONCURRENT_JOBS: {MAX_CONCURRENT_JOBS}, MAX_QUEUED_JOBS: {MAX_QUEUED_JOBS}, MAX_UPLOAD_MB: {MAX_UPLOAD_MB}")
