import os
import re
import uuid
import shutil
import threading
from typing import Dict, Any, Optional

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK_DIR = os.path.join(BASE_DIR, "workspace")
DEFAULT_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads", "MediaConverterPro")
COOKIES_FILE = os.path.join(BASE_DIR, "cookies.txt")

os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(DEFAULT_DOWNLOADS, exist_ok=True)

jobs_lock = threading.Lock()
jobs: Dict[str, Dict[str, Any]] = {}


def create_job(job_type: str, title: str = "جاري التجهيز...") -> str:
    job_id = str(uuid.uuid4())[:8]
    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "type": job_type,
            "status": "queued",
            "percent": "0%",
            "speed": "0MiB/s",
            "eta": "--:--",
            "filename": title,
            "message": "تم إنشاء المهمة بنجاح",
        }
    return job_id


def update_job(job_id: str, **kwargs):
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id].update(kwargs)


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    with jobs_lock:
        return dict(jobs[job_id]) if job_id in jobs else None


def safe_cleanup(path: str):
    if not path or not os.path.exists(path):
        return
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        else:
            os.remove(path)
    except Exception:
        pass


def sanitize_name(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip() or "output"


def resolve_save_dir(user_path: Optional[str]) -> str:
    if not user_path or not user_path.strip():
        return DEFAULT_DOWNLOADS
    cleaned = os.path.abspath(os.path.expanduser(user_path.strip()))
    os.makedirs(cleaned, exist_ok=True)
    return cleaned