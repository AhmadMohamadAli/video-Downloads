import os
import threading
import webbrowser
import subprocess
import zipfile
from pathlib import Path

import yt_dlp
import fitz
import uvicorn
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pptx import Presentation
from pptx.util import Inches

app = FastAPI(title="All-in-One Media Converter")

TMP_DIR = Path("tmp_pages")
TMP_DIR.mkdir(exist_ok=True)

# ── Thread-safe download state ──────────────────────────────────────────────

_progress_lock = threading.Lock()
_download_progress: dict = {
    "status": "idle",   # idle | starting | downloading | processing | done | error
    "percent": "0%",
    "speed": "0 MiB/s",
    "eta": "--:--",
    "filename": "",
}


def _set_progress(**kwargs) -> None:
    with _progress_lock:
        _download_progress.update(kwargs)


def _get_progress() -> dict:
    with _progress_lock:
        return dict(_download_progress)


# ── yt-dlp progress hook ────────────────────────────────────────────────────

def progress_hook(d: dict) -> None:
    status = d.get("status")
    if status == "downloading":
        _set_progress(
            status="downloading",
            percent=d.get("_percent_str", "0%").strip(),
            speed=d.get("_speed_str", "0 MiB/s").strip(),
            eta=d.get("_eta_str", "--:--").strip(),
            filename=Path(d.get("filename", "")).name,
        )
    elif status == "finished":
        _set_progress(status="processing", percent="100%")
    elif status == "error":
        _set_progress(status="error")


# ── Helpers ─────────────────────────────────────────────────────────────────

def cleanup(path: str) -> None:
    """Remove a file or directory, ignoring errors."""
    import shutil
    try:
        p = Path(path)
        if p.is_dir():
            shutil.rmtree(p)
        elif p.exists():
            p.unlink()
    except Exception:
        pass


def validate_save_path(save_path: str) -> Path:
    """Return a resolved Path and raise HTTPException on invalid input."""
    if not save_path or not save_path.strip():
        raise HTTPException(status_code=400, detail="save_path mag niet leeg zijn.")
    resolved = Path(save_path.strip()).expanduser().resolve()
    # Prevent directory traversal to sensitive system paths
    forbidden = [Path("/etc"), Path("/sys"), Path("/proc")]
    if any(str(resolved).startswith(str(f)) for f in forbidden):
        raise HTTPException(status_code=400, detail="Opgegeven pad is niet toegestaan.")
    return resolved


# ── HTML frontend ────────────────────────────────────────────────────────────

HTML_CONTENT = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>المحول الخارق | All-in-One Pro</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Tajawal', sans-serif; background: #0f172a; color: white; }
    .card { background: #1e293b; border: 1px solid #334155; transition: border-color .3s; }
    .card:hover { border-color: #3b82f6; }
    .btn-grad { background: linear-gradient(90deg, #3b82f6, #8b5cf6); }
    .btn-grad:hover { opacity: .9; transform: scale(.99); }
    .progress-bar-stripes {
      background-image: linear-gradient(45deg, rgba(255,255,255,.15) 25%, transparent 25%,
        transparent 50%, rgba(255,255,255,.15) 50%, rgba(255,255,255,.15) 75%, transparent 75%);
      background-size: 1rem 1rem;
      animation: stripe-move 1s linear infinite;
    }
    @keyframes stripe-move { from { background-position: 1rem 0; } to { background-position: 0 0; } }
  </style>
</head>
<body class="min-h-screen p-5 md:p-10">
  <div class="max-w-4xl mx-auto text-center">
    <h1 class="text-4xl font-extrabold mb-2 bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-purple-400">
      المحول الذكي الشامل
    </h1>
    <p class="text-gray-400 mb-10">تحميل بأقصى سرعة • قوائم تشغيل • أعلى جودة • اختيار مسار الحفظ</p>

    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">

      <!-- Video downloader -->
      <div class="card p-6 rounded-2xl shadow-xl flex flex-col justify-between border-blue-500 border-2 md:col-span-2">
        <div class="text-red-500 text-4xl mb-4">🎬</div>
        <h2 class="text-2xl font-bold mb-4">التحميل الصاروخي (فيديو / قوائم تشغيل)</h2>

        <div class="space-y-4 text-right max-w-2xl mx-auto w-full">
          <div>
            <label class="block text-sm text-gray-400 mb-1">الرابط (Playlist أو Video):</label>
            <input id="urlInput" type="text" placeholder="https://youtube.com/..."
              class="w-full p-3 rounded-lg bg-slate-700 border border-slate-600 focus:outline-none focus:border-blue-500 text-left" dir="ltr">
          </div>
          <div>
            <label class="block text-sm text-gray-400 mb-1">مسار الحفظ:</label>
            <input id="pathInput" type="text" placeholder="مثال: C:\Downloads أو /home/user/Downloads"
              class="w-full p-3 rounded-lg bg-slate-700 border border-slate-600 focus:outline-none focus:border-blue-500 text-left" dir="ltr">
          </div>
          <button id="downloadBtn" onclick="startDownload()"
            class="btn-grad w-full py-3 rounded-lg font-bold text-lg transition shadow-lg">
            🚀 بدء التحميل الآن
          </button>
        </div>

        <!-- Progress -->
        <div id="progressContainer" class="hidden mt-6 bg-slate-800 p-4 rounded-xl border border-slate-600 text-right">
          <div class="flex justify-between text-sm text-gray-300 mb-2">
            <span id="progressFile" class="truncate max-w-[70%] text-blue-300 font-bold">جاري التجهيز...</span>
            <span id="progressPercent" class="font-bold text-green-400">0%</span>
          </div>
          <div class="w-full bg-slate-700 rounded-full h-4 mb-2 overflow-hidden">
            <div id="progressBar" class="bg-blue-500 h-4 rounded-full progress-bar-stripes transition-all duration-300" style="width:0%"></div>
          </div>
          <div class="flex justify-between text-xs text-gray-400">
            <span id="progressSpeed">السرعة: 0 MiB/s</span>
            <span id="progressEta">الوقت المتبقي: --:--</span>
          </div>
        </div>
      </div>

      <!-- PDF → PPTX -->
      <div class="card p-6 rounded-2xl shadow-xl flex flex-col justify-between">
        <div class="text-blue-400 text-3xl mb-4">📄 ➔ 📊</div>
        <h2 class="text-xl font-bold mb-4">تحويل PDF إلى PowerPoint</h2>
        <form action="/convert-pdf/" method="post" enctype="multipart/form-data" class="space-y-4">
          <input type="file" name="files" multiple accept=".pdf" required
            class="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0
                   file:text-sm file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer">
          <button class="bg-blue-600 hover:bg-blue-700 w-full py-3 rounded-lg font-bold transition">بدء التحويل</button>
        </form>
      </div>

      <!-- Video → MP3 -->
      <div class="card p-6 rounded-2xl shadow-xl">
        <div class="text-green-400 text-3xl mb-4">🎵</div>
        <h2 class="text-xl font-bold mb-4">استخراج الصوت من الفيديو</h2>
        <form action="/convert-video/" method="post" enctype="multipart/form-data" class="space-y-4">
          <input type="file" name="file" required accept="video/*"
            class="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0
                   file:text-sm file:font-semibold file:bg-green-600 file:text-white hover:file:bg-green-700 cursor-pointer">
          <button class="bg-green-600 hover:bg-green-700 w-full py-3 rounded-lg font-bold transition">تحويل للصوت (MP3)</button>
        </form>
      </div>

    </div>
  </div>

  <script>
    let pollInterval = null;

    async function startDownload() {
      const url = document.getElementById('urlInput').value.trim();
      const savePath = document.getElementById('pathInput').value.trim();
      if (!url || !savePath) { alert('يرجى ملء الرابط ومسار الحفظ.'); return; }

      const btn = document.getElementById('downloadBtn');
      btn.disabled = true;
      btn.innerHTML = '⏳ جاري التحميل...';
      btn.classList.add('opacity-50', 'cursor-not-allowed');
      document.getElementById('progressContainer').classList.remove('hidden');

      pollInterval = setInterval(updateProgress, 1000);

      const formData = new FormData();
      formData.append('url', url);
      formData.append('save_path', savePath);

      try {
        const res = await fetch('/download-video/', { method: 'POST', body: formData });
        if (!res.ok) {
          const err = await res.json();
          alert('❌ Fout: ' + (err.detail || 'Onbekende fout'));
        }
      } catch (e) {
        console.error(e);
      }
    }

    async function updateProgress() {
      try {
        const data = await (await fetch('/progress/')).json();
        const pct = data.percent || '0%';
        document.getElementById('progressBar').style.width = pct;
        document.getElementById('progressPercent').textContent = pct;
        document.getElementById('progressSpeed').textContent = 'السرعة: ' + data.speed;
        document.getElementById('progressEta').textContent = 'الوقت المتبقي: ' + data.eta;
        document.getElementById('progressFile').textContent =
          data.status === 'processing' ? 'جاري دمج الفيديو والصوت... ⏳' : (data.filename || '');

        if (data.status === 'done' || data.status === 'error') {
          clearInterval(pollInterval);
          const btn = document.getElementById('downloadBtn');
          btn.disabled = false;
          btn.innerHTML = '🚀 بدء التحميل الآن';
          btn.classList.remove('opacity-50', 'cursor-not-allowed');
          if (data.status === 'done') {
            document.getElementById('progressFile').textContent = '✅ تمت العملية بنجاح!';
            alert('✅ اكتمل التحميل!');
          } else {
            document.getElementById('progressFile').textContent = '❌ حدث خطأ';
            alert('❌ حدث خطأ أثناء التحميل.');
          }
        }
      } catch (e) { console.error(e); }
    }
  </script>
</body>
</html>
"""


# ── Routes ───────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def home() -> str:
    return HTML_CONTENT


@app.get("/progress/")
async def get_progress() -> JSONResponse:
    return JSONResponse(content=_get_progress())


@app.post("/download-video/")
def download_video(url: str = Form(...), save_path: str = Form(...)) -> dict:
    """Download a video or playlist via yt-dlp (runs in a FastAPI thread pool)."""
    save_dir = validate_save_path(save_path)
    save_dir.mkdir(parents=True, exist_ok=True)

    _set_progress(status="starting", percent="0%", filename="جاري قراءة الرابط...")

    is_playlist = "playlist" in url.lower()
    outtmpl = str(
        save_dir / ("%(playlist_title)s/%(title)s.%(ext)s" if is_playlist else "%(title)s.%(ext)s")
    )

    ydl_opts = {
        "format": "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
        "outtmpl": outtmpl,
        "noplaylist": not is_playlist,
        "concurrent_fragment_downloads": os.cpu_count() or 4,
        "merge_output_format": "mp4",
        "ignoreerrors": True,
        "nocolor": True,
        "progress_hooks": [progress_hook],
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        _set_progress(status="done", percent="100%")
        return {"status": "success"}
    except Exception as exc:
        _set_progress(status="error")
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/convert-pdf/")
async def convert_pdf(
    background_tasks: BackgroundTasks,
    files: list[UploadFile] = File(...),
) -> FileResponse:
    """Convert one or more PDFs to PPTX files."""
    output_files: list[str] = []

    for upload in files:
        base_name = Path(upload.filename).stem
        temp_pdf = Path(f"temp_{upload.filename}")
        temp_ppt = Path(f"{base_name}.pptx")

        temp_pdf.write_bytes(await upload.read())

        doc = fitz.open(str(temp_pdf))
        prs = Presentation()

        for page_num, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = TMP_DIR / f"{base_name}_{page_num}.png"
            pix.save(str(img_path))

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            prs.slide_width = Inches(page.rect.width / 72)
            prs.slide_height = Inches(page.rect.height / 72)
            slide.shapes.add_picture(str(img_path), 0, 0, width=prs.slide_width, height=prs.slide_height)
            img_path.unlink(missing_ok=True)

        prs.save(str(temp_ppt))
        doc.close()
        temp_pdf.unlink(missing_ok=True)
        output_files.append(str(temp_ppt))

    if len(output_files) == 1:
        background_tasks.add_task(cleanup, output_files[0])
        return FileResponse(output_files[0], filename=Path(output_files[0]).name)

    zip_name = "converted_docs.zip"
    with zipfile.ZipFile(zip_name, "w") as zf:
        for f in output_files:
            zf.write(f)
            Path(f).unlink(missing_ok=True)

    background_tasks.add_task(cleanup, zip_name)
    return FileResponse(zip_name, filename=zip_name)


@app.post("/convert-video/")
async def convert_video(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
) -> FileResponse:
    """Extract audio from an uploaded video file and return MP3."""
    input_path = Path(f"temp_{file.filename}")
    output_path = Path(f"{Path(file.filename).stem}.mp3")

    input_path.write_bytes(await file.read())

    result = subprocess.run(
        [
            "ffmpeg", "-i", str(input_path),
            "-threads", str(os.cpu_count() or 4),
            "-vn", "-ab", "192k", "-ar", "44100", "-y",
            str(output_path),
        ],
        capture_output=True,
    )
    input_path.unlink(missing_ok=True)

    if result.returncode != 0:
        raise HTTPException(status_code=500, detail="ffmpeg conversion failed.")

    background_tasks.add_task(cleanup, str(output_path))
    return FileResponse(str(output_path), filename=output_path.name)


# ── Entry point ───────────────────────────────────────────────────────────────

def _open_browser() -> None:
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    threading.Timer(1.5, _open_browser).start()
    uvicorn.run(app, host="127.0.0.1", port=8000)
