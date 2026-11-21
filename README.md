# 🎬 All-in-One Media & PDF Studio Pro

A local desktop-window web app built with **FastAPI**, **Uvicorn**, and **pywebview** that gives you a universal media downloader (powered by `yt-dlp`), fast audio extraction to MP3 (`ffmpeg`), and PDF utilities including PowerPoint slide generation (`PyMuPDF` + `python-pptx`).

---

## ✨ Features

* **Universal Video & Playlist Downloader:** Download videos and playlists in high quality (`MP4`) or extract audio (`MP3`) directly via `yt-dlp`, with live progress tracking (`percent`, `speed`, `ETA`, status messages) and an optional `cookies.txt` file when needed for restricted links.
* **Fast Video → MP3 Conversion:** Convert uploaded video or audio files to `MP3` (`128k`, `192k`, or `320k`) via `ffmpeg` and download the result directly.
* **PDF Tools:**
  * Convert PDF files into **editable PowerPoint (`.pptx`)** slides extracted from page text.
  * Convert PDF files into **visual PowerPoint (`.pptx`)** slides rendered at `2×` resolution.
  * **Merge** multiple PDF files into a single PDF document.
  * **Extract text** from PDF files into a `.txt` file.
* **Standalone Desktop Window (`run.py`):** Starts the local FastAPI server in the background and opens the UI inside a dedicated `pywebview` window.

---

## 🖥️ Requirements

* **Python 3.10+**
* **[ffmpeg](https://ffmpeg.org/download.html)** installed and available in your system `PATH`
* **Windows / macOS / Linux** (note: `pywebview` and `PyMuPDF` depend on local OS/runtime policies and libraries)

---

## 📁 Project Structure

```text
video-Downloads-pro/
├── run.py                  # Starts FastAPI in background and launches the desktop window
├── requirements.txt        # Python dependencies
├── cookies.txt             # Optional cookies file for authenticated/restricted links
├── README.md
└── app/
    ├── __init__.py
    ├── main.py             # FastAPI app and HTTP endpoints
    ├── config.py           # Paths, job tracking, and file helpers
    ├── services/
    │   ├── __init__.py
    │   ├── downloader.py   # yt-dlp download & MP3 extraction logic
    │   ├── media.py        # ffmpeg video-to-MP3 conversion
    │   └── pdf_tools.py    # PDF-to-PPTX, PDF merge, and text extraction
    └── templates/
        └── index.html      # Web UI (HTML + Tailwind CSS + JavaScript)
		==============================================
🚀 Installation & Usage

# 1. Clone the repository
git clone [https://github.com/AhmadMohamadAli/video-Downloads.git](https://github.com/AhmadMohamadAli/video-Downloads.git)
cd video-Downloads

# 2. Create and activate a virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

# 3. Install dependencies
python -m pip install --upgrade pip
pip install -r requirements.txt

# 4. Run the application
python run.py



🛠️ Tech StackComponentPurposeFastAPI & UvicornLocal backend API and ASGI serverpywebviewStandalone desktop application windowyt-dlpVideo, playlist, and audio downloadingffmpeg (system)Audio extraction and MP3 encodingPyMuPDF (fitz)PDF rendering, merging, and text extractionpython-pptxPowerPoint (.pptx) presentation generation⚠️ Notes & TroubleshootingCustom Save Path: If left empty when starting a download, files are saved under your user Downloads/MediaConverterPro folder.Windows Application Control (_mupdf DLL error): If an Application Control policy blocks _mupdf when importing PyMuPDF, check your local security/AppLocker policy or run inside an allowed virtual environment (.venv).📝 LicenseMIT — free to use, modify, and distribute.