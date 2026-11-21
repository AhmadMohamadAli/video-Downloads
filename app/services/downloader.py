import os
import re
from typing import Dict, Any
import yt_dlp
from app.config import COOKIES_FILE, update_job


def make_yt_hook(job_id: str):
    def progress_hook(d):
        status = d.get("status")
        if status == "downloading":
            clean_pct = re.sub(r"\x1b\[[0-9;]*m", "", d.get("_percent_str", "0%")).strip()
            clean_spd = re.sub(r"\x1b\[[0-9;]*m", "", d.get("_speed_str", "0MiB/s")).strip()
            clean_eta = re.sub(r"\x1b\[[0-9;]*m", "", d.get("_eta_str", "--:--")).strip()
            fname = os.path.basename(
                d.get("filename") or d.get("info_dict", {}).get("title") or "جاري التحميل..."
            )
            update_job(
                job_id,
                status="downloading",
                percent=clean_pct,
                speed=clean_spd,
                eta=clean_eta,
                filename=fname,
            )
        elif status == "finished":
            update_job(
                job_id,
                status="processing",
                percent="100%",
                message="اكتمل تنزيل الجزء، جاري الدمج والمعالجة...",
            )
        elif status == "error":
            update_job(job_id, status="error", message="حدث خطأ أثناء التنزيل")

    return progress_hook


def run_universal_download(job_id: str, url: str, mode: str, bitrate: str, save_dir: str):
    update_job(job_id, status="downloading", message="جاري تحليل الرابط والاتصال...")

    is_playlist = ("list=" in url) or ("playlist" in url)
    outtmpl = (
        os.path.join(save_dir, "%(playlist_title)s", "%(playlist_index)s - %(title)s.%(ext)s")
        if is_playlist
        else os.path.join(save_dir, "%(title)s.%(ext)s")
    )

    ydl_opts: Dict[str, Any] = {
        "outtmpl": outtmpl,
        "noplaylist": False,
        "concurrent_fragment_downloads": min(os.cpu_count() or 4, 8),
        "ignoreerrors": True,
        "nocolor": True,
        "progress_hooks": [make_yt_hook(job_id)],
        "extractor_args": {
            "youtube": {
                "player_client": ["android", "tv", "web_embedded"]
            }
        },
    }

    if os.path.exists(COOKIES_FILE):
        ydl_opts["cookiefile"] = COOKIES_FILE

    if mode == "audio_mp3":
        br_num = bitrate.replace("k", "") if bitrate.endswith("k") else "192"
        ydl_opts.update({
            "format": "bestaudio/best",
            "postprocessors": [{
                "key": "FFmpegExtractAudio",
                "preferredcodec": "mp3",
                "preferredquality": br_num,
            }],
        })
    elif mode == "video_1080":
        ydl_opts.update({
            "format": "bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080][ext=mp4]/best",
            "merge_output_format": "mp4",
        })
    else:
        ydl_opts.update({
            "format": "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
            "merge_output_format": "mp4",
        })

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        update_job(job_id, status="done", percent="100%", message=f"تم الحفظ في: {save_dir}")
    except Exception as exc:
        update_job(job_id, status="error", message=f"فشل التحميل: {str(exc)}")