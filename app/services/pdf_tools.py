import os
from typing import List
from pptx import Presentation
from pptx.util import Inches, Pt
from app.config import safe_cleanup

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
    FITZ_IMPORT_ERROR = ""
except Exception as exc:
    fitz = None
    FITZ_AVAILABLE = False
    FITZ_IMPORT_ERROR = str(exc)


def _require_fitz():
    if not FITZ_AVAILABLE or fitz is None:
        raise RuntimeError(
            f"تعذر تحميل PyMuPDF (fitz) على هذا الجهاز: {FITZ_IMPORT_ERROR}"
        )


def build_editable_pptx(pdf_path: str, pptx_path: str):
    _require_fitz()
    doc = fitz.open(pdf_path)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    try:
        for idx, page in enumerate(doc):
            slide = prs.slides.add_slide(blank_layout)
            text = page.get_text("text").strip()
            lines = [line.strip() for line in text.splitlines() if line.strip()]

            title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.7))
            title_tf = title_box.text_frame
            title_tf.word_wrap = True
            title_p = title_tf.paragraphs[0]
            title_p.text = lines[0][:120] if lines else f"صفحة {idx + 1}"
            title_p.font.size = Pt(24)
            title_p.font.bold = True

            body_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.25), Inches(12.1), Inches(5.7))
            body_tf = body_box.text_frame
            body_tf.word_wrap = True
            body_text = (
                "\n".join(lines[1:])
                if len(lines) > 1
                else ("لا يوجد نص قابل للاستخراج في هذه الصفحة." if not lines else "")
            )
            body_p = body_tf.paragraphs[0]
            body_p.text = body_text[:3500]
            body_p.font.size = Pt(14)

        prs.save(pptx_path)
    finally:
        doc.close()


def build_visual_pptx(pdf_path: str, pptx_path: str, img_tmp_dir: str):
    _require_fitz()
    doc = fitz.open(pdf_path)
    prs = Presentation()

    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = os.path.join(img_tmp_dir, f"page_{page_num}.png")
            pix.save(img_path)

            prs.slide_width = Inches(page.rect.width / 72)
            prs.slide_height = Inches(page.rect.height / 72)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            safe_cleanup(img_path)

        prs.save(pptx_path)
    finally:
        doc.close()


def merge_pdfs(pdf_paths: List[str], output_pdf: str):
    _require_fitz()
    merged_doc = fitz.open()
    try:
        for p_path in pdf_paths:
            src_doc = fitz.open(p_path)
            try:
                merged_doc.insert_pdf(src_doc)
            finally:
                src_doc.close()
        merged_doc.save(output_pdf)
    finally:
        merged_doc.close()


def extract_pdfs_text(pdf_paths: List[str], output_txt: str):
    _require_fitz()
    all_text_parts = []
    for p_path in pdf_paths:
        doc = fitz.open(p_path)
        try:
            all_text_parts.append(f"=== {os.path.basename(p_path)} ===\n")
            for i, page in enumerate(doc):
                all_text_parts.append(f"--- صفحة {i + 1} ---\n{page.get_text('text')}\n")
        finally:
            doc.close()

    with open(output_txt, "w", encoding="utf-8") as tf:
        tf.write("\n".join(all_text_parts))