# 🎬 All-in-One Media Converter

A local web app built with **FastAPI** that lets you download videos, convert PDFs to PowerPoint, and extract audio from video files — all from your browser.

---

## ✨ Features

| Feature | Details |
|---|---|
| 🚀 Video & Playlist Download | Downloads YouTube videos and playlists in the highest quality (MP4) via `yt-dlp` |
| 📊 PDF → PowerPoint | Converts one or more PDF files to `.pptx` slides |
| 🎵 Video → MP3 | Extracts audio from any uploaded video file at 192 kbps |
| 📡 Live Progress Bar | Real-time download progress (speed, ETA, filename) shown in the browser |
| 📁 Custom Save Path | Choose exactly where downloaded files are saved on your machine |

---

## 🖥️ Requirements

- Python 3.10+
- [ffmpeg](https://ffmpeg.org/download.html) installed and available in your PATH

---

## 🚀 Installation

```bash
# 1. Clone the repository
git clone https://github.com/AhmadMohamadAli/video-Downloads.git
cd video-Downloads

# 2. (Recommended) Create a virtual environment
python -m venv venv
source venv/bin/activate      # Linux / macOS
venv\Scripts\activate         # Windows

# 3. Install dependencies
pip install fastapi uvicorn yt-dlp pymupdf python-pptx
```

---

## ▶️ Usage

```bash
python app.py
```

The app opens automatically in your browser at `http://127.0.0.1:8000`.

---

## 📁 Project Structure

```
video-Downloads/
├── app.py          # FastAPI backend + embedded HTML frontend
└── README.md
```

---

## ⚙️ How It Works

- **Video download** — A POST request triggers `yt-dlp` inside a FastAPI thread. The browser polls `/progress/` every second to update the live progress bar.
- **PDF → PPTX** — Each PDF page is rasterised with `PyMuPDF` (fitz) at 2× resolution and inserted as a full-slide image in a new `python-pptx` presentation.
- **Video → MP3** — `ffmpeg` strips the audio track and encodes it to MP3 at 192 kbps, using all available CPU cores.

---

## 🛠️ Dependencies

| Package | Purpose |
|---|---|
| `fastapi` | Web framework |
| `uvicorn` | ASGI server |
| `yt-dlp` | Video downloading |
| `PyMuPDF` (fitz) | PDF rendering |
| `python-pptx` | PowerPoint generation |
| `ffmpeg` (system) | Audio extraction |

---

## 📝 License

MIT — free to use, modify, and distribute.
