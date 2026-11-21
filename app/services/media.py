import os
import subprocess


def convert_to_mp3(input_path: str, output_path: str, bitrate: str = "192k"):
    valid_bitrates = {"128k", "192k", "320k"}
    br = bitrate if bitrate in valid_bitrates else "192k"
    threads = str(os.cpu_count() or 4)

    cmd = [
        "ffmpeg",
        "-y",
        "-threads", threads,
        "-i", input_path,
        "-vn",
        "-acodec", "libmp3lame",
        "-ab", br,
        "-ar", "44100",
        "-q:a", "2",
        output_path,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not os.path.exists(output_path):
        raise RuntimeError(f"فشل تحويل الصوت عبر ffmpeg: {proc.stderr[-300:]}")