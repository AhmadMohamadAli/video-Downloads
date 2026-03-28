
import os
import threading
import webbrowser
import subprocess
import yt_dlp
import fitz
import zipfile
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
TMP_DIR = "tmp_pages"

if not os.path.exists(TMP_DIR):
    os.makedirs(TMP_DIR)

def cleanup(file_path: str):
    if os.path.exists(file_path):
        try:
            if os.path.isdir(file_path):
                import shutil
                shutil.rmtree(file_path)
            else:
                os.remove(file_path)
        except: pass

# ---------------- متغيرات تتبع التحميل (Global State) ---------------- #
# سنستخدم هذا القاموس لتخزين حالة التحميل وإرسالها للمتصفح
download_progress = {
    "status": "idle",
    "percent": "0%",
    "speed": "0MiB/s",
    "eta": "00:00",
    "filename": "جاري بدء التحميل..."
}

def progress_hook(d):
    """دالة تلتقط بيانات التحميل من yt-dlp في الوقت الفعلي"""
    if d['status'] == 'downloading':
        # إزالة أي مسافات أو رموز غير مرغوب فيها
        download_progress['percent'] = d.get('_percent_str', '0%').strip()
        download_progress['speed'] = d.get('_speed_str', '0MiB/s').strip()
        download_progress['eta'] = d.get('_eta_str', '00:00').strip()
        download_progress['filename'] = os.path.basename(d.get('filename', ''))
        download_progress['status'] = 'downloading'
    elif d['status'] == 'finished':
        download_progress['percent'] = '100%'
        download_progress['status'] = 'processing' # جاري دمج الصوت والصورة
    elif d['status'] == 'error':
        download_progress['status'] = 'error'

# ---------------- واجهة المستخدم (HTML & JS) ---------------- #

html_content = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>المحول الخارق | All-in-One Pro</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Tajawal', sans-serif; background: #0f172a; color: white; }
        .card { background: #1e293b; border: 1px solid #334155; transition: 0.3s; }
        .card:hover { border-color: #3b82f6; }
        .btn-grad { background: linear-gradient(90deg, #3b82f6, #8b5cf6); }
        .btn-grad:hover { opacity: 0.9; transform: scale(0.99); }
        /* تصميم شريط التحميل */
        .progress-bar-stripes {
            background-image: linear-gradient(45deg,rgba(255,255,255,.15) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.15) 50%,rgba(255,255,255,.15) 75%,transparent 75%,transparent);
            background-size: 1rem 1rem;
            animation: progress-animation 1s linear infinite;
        }
        @keyframes progress-animation { from {background-position: 1rem 0;} to {background-position: 0 0;} }
    </style>
</head>
<body class="min-h-screen p-5 md:p-10">
    <div class="max-w-4xl mx-auto text-center">
        <h1 class="text-4xl font-extrabold mb-2 bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-purple-400">
            المحول الذكي الشامل
        </h1>
        <p class="text-gray-400 mb-10">تحميل بأقصى سرعة • قوائم تشغيل • أعلى جودة • اختيار مسار الحفظ</p>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            
            <div class="card p-6 rounded-2xl shadow-xl flex flex-col justify-between border-blue-500 border-2 md:col-span-2">
                <div>
                    <div class="text-red-500 text-4xl mb-4">🎬</div>
                    <h2 class="text-2xl font-bold mb-4">التحميل الصاروخي (فيديو / قوائم تشغيل)</h2>
                </div>
                
                <form id="downloadForm" class="space-y-4 text-right max-w-2xl mx-auto w-full">
                    <div>
                        <label class="block text-sm text-gray-400 mb-1">الرابط (Playlist أو Video):</label>
                        <input type="text" name="url" required placeholder="https://youtube.com/..." class="w-full p-3 rounded-lg bg-slate-700 border border-slate-600 focus:outline-none focus:border-blue-500 text-left" dir="ltr">
                    </div>
                    <div>
                        <label class="block text-sm text-gray-400 mb-1">مسار الحفظ (أين تريد حفظ الملفات؟):</label>
                        <input type="text" name="save_path" required placeholder="مثال: C:\\Downloads أو /Users/name/Downloads" class="w-full p-3 rounded-lg bg-slate-700 border border-slate-600 focus:outline-none focus:border-blue-500 text-left" dir="ltr">
                    </div>
                    <button type="submit" id="downloadBtn" class="btn-grad w-full py-3 rounded-lg font-bold text-lg transition shadow-lg">🚀 بدء التحميل الآن</button>
                </form>

                <div id="progressContainer" class="hidden mt-6 bg-slate-800 p-4 rounded-xl border border-slate-600 text-right">
                    <div class="flex justify-between text-sm text-gray-300 mb-2">
                        <span id="progressFile" class="truncate max-w-[70%] text-blue-300 font-bold">جاري التجهيز...</span>
                        <span id="progressPercent" class="font-bold text-green-400">0%</span>
                    </div>
                    <div class="w-full bg-slate-700 rounded-full h-4 mb-2 overflow-hidden">
                        <div id="progressBar" class="bg-blue-500 h-4 rounded-full progress-bar-stripes transition-all duration-300" style="width: 0%"></div>
                    </div>
                    <div class="flex justify-between text-xs text-gray-400">
                        <span id="progressSpeed">السرعة: 0MiB/s</span>
                        <span id="progressEta">الوقت المتبقي: 00:00</span>
                    </div>
                </div>

            </div>

            <div class="card p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <div class="text-blue-400 text-3xl mb-4">📄 ➔ 📊</div>
                    <h2 class="text-xl font-bold mb-4">تحويل PDF إلى PowerPoint</h2>
                </div>
                <form action="/convert-pdf/" method="post" enctype="multipart/form-data" class="space-y-4">
                    <input type="file" name="files" multiple accept=".pdf" required class="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer">
                    <button class="bg-blue-600 hover:bg-blue-700 w-full py-3 rounded-lg font-bold transition">بدء التحويل</button>
                </form>
            </div>

            <div class="card p-6 rounded-2xl shadow-xl">
                <div class="text-green-400 text-3xl mb-4">🎵</div>
                <h2 class="text-xl font-bold mb-4">استخراج الصوت من الفيديو</h2>
                <form action="/convert-video/" method="post" enctype="multipart/form-data" class="space-y-4">
                    <input type="file" name="file" required accept="video/*" class="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-green-600 file:text-white hover:file:bg-green-700 cursor-pointer">
                    <button class="bg-green-600 hover:bg-green-700 w-full py-3 rounded-lg font-bold transition">تحويل للصوت (MP3)</button>
                </form>
            </div>

        </div>
    </div>

    <script>
        document.getElementById('downloadForm').addEventListener('submit', async (e) => {
            e.preventDefault(); // منع تحديث الصفحة
            
            const form = e.target;
            const formData = new FormData(form);
            const btn = document.getElementById('downloadBtn');
            const progressContainer = document.getElementById('progressContainer');
            
            // تعطيل الزر وإظهار شريط التحميل
            btn.disabled = true;
            btn.innerHTML = '⏳ جاري التحميل... يرجى الانتظار';
            btn.classList.add('opacity-50', 'cursor-not-allowed');
            progressContainer.classList.remove('hidden');

            // بدء استعلام حالة التحميل كل ثانية
            const progressInterval = setInterval(async () => {
                try {
                    const res = await fetch('/progress/');
                    const data = await res.json();

                    // تحديث الواجهة
                    document.getElementById('progressBar').style.width = data.percent;
                    document.getElementById('progressPercent').innerText = data.percent;
                    document.getElementById('progressSpeed').innerText = `السرعة: ${data.speed}`;
                    document.getElementById('progressEta').innerText = `الوقت المتبقي: ${data.eta}`;
                    
                    if (data.status === 'downloading' || data.status === 'processing') {
                        document.getElementById('progressFile').innerText = data.filename;
                    }
                    if (data.status === 'processing') {
                        document.getElementById('progressFile').innerText = "جاري دمج الفيديو والصوت بأعلى جودة... ⏳";
                    }

                    // عند الانتهاء
                    if (data.status === 'done' || data.status === 'error') {
                        clearInterval(progressInterval);
                        btn.disabled = false;
                        btn.innerHTML = '🚀 بدء التحميل الآن';
                        btn.classList.remove('opacity-50', 'cursor-not-allowed');
                        
                        if (data.status === 'done') {
                            alert('✅ اكتمل التحميل بنجاح وتم حفظ الملفات في المجلد المحدد!');
                            document.getElementById('progressFile').innerText = "✅ تمت العملية بنجاح!";
                        } else {
                            alert('❌ حدث خطأ أثناء التحميل. تأكد من الرابط والمساحة المتوفرة.');
                            document.getElementById('progressFile').innerText = "❌ حدث خطأ";
                        }
                    }
                } catch (err) {
                    console.error("Error fetching progress", err);
                }
            }, 1000); // تحديث كل ثانية (1000ms)

            // إرسال طلب التحميل الفعلي
            try {
                await fetch('/download-video/', {
                    method: 'POST',
                    body: formData
                });
            } catch (err) {
                console.error("Download Request Failed", err);
            }
        });
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def home():
    return html_content

# نقطة اتصال (Endpoint) لإرسال حالة التحميل الحالية للمتصفح
@app.get("/progress/")
async def get_progress():
    return JSONResponse(content=download_progress)

# ---------------- محمل الفيديو والقوائم ---------------- #
# ملاحظة: تم استخدام `def` بدلاً من `async def` هنا عمداً 
# لكي يقوم FastAPI بتشغيلها في Thread منفصل ولا يوقف استجابة الـ Progress
@app.post("/download-video/")
def download_video(url: str = Form(...), save_path: str = Form(...)):
    
    # إعادة ضبط حالة التحميل
    download_progress['status'] = 'starting'
    download_progress['percent'] = '0%'
    download_progress['filename'] = 'جاري التجهيز وقراءة الرابط...'
    
    os.makedirs(save_path, exist_ok=True)

    ydl_opts = {
        'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best',
        'outtmpl': os.path.join(save_path, '%(playlist_title)s', '%(title)s.%(ext)s') if 'playlist' in url else os.path.join(save_path, '%(title)s.%(ext)s'),
        'noplaylist': False, 
        'concurrent_fragment_downloads': os.cpu_count() or 4, # استخدام كامل أنوية الكمبيوتر
        'merge_output_format': 'mp4',
        'ignoreerrors': True,
        'nocolor': True, # هام جداً لمنع الرموز الغريبة في نسبة التحميل
        'progress_hooks': [progress_hook], # ربط دالة المتابعة
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        download_progress['status'] = 'done'
        return {"status": "success"}
    except Exception as e:
        download_progress['status'] = 'error'
        return {"status": "error", "message": str(e)}

# ---------------- PDF → PPTX ---------------- #
@app.post("/convert-pdf/")
async def convert_pdf(background_tasks: BackgroundTasks, files: list[UploadFile] = File(...)):
    output_files = []
    for file in files:
        base_name = os.path.splitext(file.filename)[0]
        temp_pdf = f"temp_{file.filename}"
        temp_ppt = f"{base_name}.pptx"

        with open(temp_pdf, "wb") as f:
            f.write(await file.read())

        doc = fitz.open(temp_pdf)
        prs = Presentation()

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(2,2)) 
            img_path = f"{TMP_DIR}/{base_name}_{page_num}.png"
            pix.save(img_path)

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            prs.slide_width, prs.slide_height = Inches(page.rect.width/72), Inches(page.rect.height/72)
            slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            os.remove(img_path)

        prs.save(temp_ppt)
        doc.close()
        os.remove(temp_pdf)
        output_files.append(temp_ppt)

    if len(output_files) == 1:
        background_tasks.add_task(cleanup, output_files[0])
        return FileResponse(output_files[0], filename=output_files[0])
    
    zip_name = "converted_docs.zip"
    with zipfile.ZipFile(zip_name, "w") as z:
        for f in output_files:
            z.write(f)
            os.remove(f)
    background_tasks.add_task(cleanup, zip_name)
    return FileResponse(zip_name, filename=zip_name)

# ---------------- Video → MP3 ---------------- #
@app.post("/convert-video/")
async def convert_video(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    input_file = f"temp_{file.filename}"
    output_file = f"{os.path.splitext(file.filename)[0]}.mp3"

    with open(input_file, "wb") as f:
        f.write(await file.read())

    subprocess.run([
        "ffmpeg", "-i", input_file, 
        "-threads", str(os.cpu_count() or 4),
        "-vn", "-ab", "192k", "-ar", "44100", "-y", 
        output_file
    ])

    os.remove(input_file)
    background_tasks.add_task(cleanup, output_file)
    return FileResponse(output_file, filename=output_file)

# ---------------- التشغيل التلقائي ---------------- #
def open_browser():
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    threading.Timer(1.5, open_browser).start()
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)