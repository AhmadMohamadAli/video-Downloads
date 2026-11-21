import os
import uuid
import shutil
import zipfile
from typing import List

from fastapi import FastAPI, UploadFile, File, BackgroundTasks, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

from app.config import (
    WORK_DIR,
    create_job,
    get_job,
    safe_cleanup,
    sanitize_name,
    resolve_save_dir,
)
from app.services.downloader import run_universal_download
from app.services.media import convert_to_mp3
from app.services.pdf_tools import (
    build_editable_pptx,
    build_visual_pptx,
    merge_pdfs,
    extract_pdfs_text,
)

app = FastAPI(title="All-in-One Media Converter Pro", version="2.0.0")
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")


@app.get("/", response_class=HTMLResponse)
async def home():
    if not os.path.exists(TEMPLATE_PATH):
        return HTMLResponse(
            content=f"<h3>ملف الواجهة غير موجود:</h3><code>{TEMPLATE_PATH}</code>",
            status_code=500,
        )
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/progress/{job_id}")
async def get_job_progress(job_id: str):
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="المهمة غير موجودة")
    return JSONResponse(content=job)


@app.post("/download-universal/")
async def start_universal_download(
    background_tasks: BackgroundTasks,
    url: str = Form(...),
    mode: str = Form("video_best"),
    bitrate: str = Form("192k"),
    save_path: str = Form(""),
):
    target_url = url.strip()
    if not target_url.startswith(("http://", "https://")):
        raise HTTPException(status_code=400, detail="الرجاء إدخال رابط يبدأ بـ http:// أو https://")

    save_dir = resolve_save_dir(save_path)
    job_id = create_job("universal_download", title=target_url)
    background_tasks.add_task(run_universal_download, job_id, target_url, mode, bitrate, save_dir)
    return {"job_id": job_id, "status": "queued"}


@app.post("/convert-video/")
async def convert_video_instant(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    bitrate: str = Form("192k"),
):
    session_dir = os.path.join(WORK_DIR, f"mp3_{uuid.uuid4().hex[:8]}")
    os.makedirs(session_dir, exist_ok=True)

    safe_orig = sanitize_name(file.filename or "video_input")
    base_name = os.path.splitext(safe_orig)[0]
    input_path = os.path.join(session_dir, safe_orig)
    output_path = os.path.join(session_dir, f"{base_name}.mp3")

    try:
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        convert_to_mp3(input_path, output_path, bitrate=bitrate)
        safe_cleanup(input_path)

        background_tasks.add_task(safe_cleanup, session_dir)
        return FileResponse(output_path, media_type="audio/mpeg", filename=f"{base_name}.mp3")
    except Exception as exc:
        safe_cleanup(session_dir)
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/pdf-tools/")
async def pdf_tools_endpoint(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    action: str = Form("pptx_editable"),
):
    if not files:
        raise HTTPException(status_code=400, detail="يرجى رفع ملف PDF واحد على الأقل")

    session_dir = os.path.join(WORK_DIR, f"pdf_{uuid.uuid4().hex[:8]}")
    os.makedirs(session_dir, exist_ok=True)
    saved_pdfs: List[str] = []

    try:
        for idx, up in enumerate(files):
            clean_name = sanitize_name(up.filename or f"doc_{idx + 1}.pdf")
            if not clean_name.lower().endswith(".pdf"):
                clean_name += ".pdf"
            p_path = os.path.join(session_dir, f"{idx}_{clean_name}")
            with open(p_path, "wb") as f:
                shutil.copyfileobj(up.file, f)
            saved_pdfs.append(p_path)

        if action == "merge_pdf":
            out_pdf = os.path.join(session_dir, "merged_document.pdf")
            merge_pdfs(saved_pdfs, out_pdf)
            background_tasks.add_task(safe_cleanup, session_dir)
            return FileResponse(out_pdf, media_type="application/pdf", filename="merged_document.pdf")

        if action == "extract_txt":
            out_txt = os.path.join(session_dir, "extracted_text.txt")
            extract_pdfs_text(saved_pdfs, out_txt)
            background_tasks.add_task(safe_cleanup, session_dir)
            return FileResponse(out_txt, media_type="text/plain; charset=utf-8", filename="extracted_text.txt")

        generated_pptx: List[str] = []
        for p_path in saved_pdfs:
            base = sanitize_name(os.path.splitext(os.path.basename(p_path))[0])
            out_ppt = os.path.join(session_dir, f"{base}.pptx")
            if action == "pptx_visual":
                build_visual_pptx(p_path, out_ppt, session_dir)
            else:
                build_editable_pptx(p_path, out_ppt)
            generated_pptx.append(out_ppt)

        if len(generated_pptx) == 1:
            background_tasks.add_task(safe_cleanup, session_dir)
            return FileResponse(
                generated_pptx[0],
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                filename=os.path.basename(generated_pptx[0]),
            )

        zip_path = os.path.join(session_dir, "converted_presentations.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for ppt in generated_pptx:
                zf.write(ppt, arcname=os.path.basename(ppt))

        background_tasks.add_task(safe_cleanup, session_dir)
        return FileResponse(zip_path, media_type="application/zip", filename="converted_presentations.zip")

    except Exception as exc:
        safe_cleanup(session_dir)
        raise HTTPException(status_code=500, detail=f"فشلت عملية PDF: {str(exc)}")