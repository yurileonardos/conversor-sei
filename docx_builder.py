from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from constants import *
from masking import apply_masking
import pdfplumber

def build_docx(tr_bytes, proposal_files, debug=False):
    doc = Document()

    section = doc.sections[0]
    section.page_width  = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.left_margin = MARGIN_LEFT
    section.right_margin = MARGIN_RIGHT
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM

    # --- TR ---
    pdf = pdfplumber.open(BytesIO(tr_bytes))
    images = pdf_to_images(tr_bytes)
    mask_state = {'active': False, 'cut_x_percent': None}

    for i, img in enumerate(images):
        img, mask_state = apply_masking(img, pdf.pages[i], mask_state, debug)
        img = img.resize((TARGET_WIDTH, TARGET_HEIGHT))

        bio = BytesIO()
        img.save(bio, "JPEG", quality=85, optimize=True)
        bio.seek(0)

        doc.add_picture(bio, width=Cm(18))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i < len(images) - 1:
            doc.add_page_break()

    # --- PROPOSTAS ---
    for proposal in proposal_files:
        images = pdf_to_images(proposal)
        for img in images:
            img = img.resize((TARGET_WIDTH, TARGET_HEIGHT))
            bio = BytesIO()
            img.save(bio, "JPEG", quality=85, optimize=True)
            bio.seek(0)

            doc.add_page_break()
            doc.add_picture(bio, width=Cm(18))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output
