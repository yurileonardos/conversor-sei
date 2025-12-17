from docx import Document
from docx.shared import Cm
from io import BytesIO

import pdfplumber

from pdf_utils import pdf_to_images
from masking_table_price import apply_table_price_mask


def build_docx(tr_bytes, proposal_files, debug=False):
    doc = Document()
    _configure_page(doc)

    # ===== TR SEMPRE PRIMEIRO =====
    tr_images = _process_tr(tr_bytes, debug)

    if not tr_images:
        raise RuntimeError("Nenhuma página gerada para o Termo de Referência.")

    for i, img in enumerate(tr_images):
        _add_image(doc, img)
        if i < len(tr_images) - 1:
            doc.add_page_break()

    # ===== PROPOSTAS (SEM MÁSCARA) =====
    for proposal in proposal_files:
        images = pdf_to_images(proposal)
        for i, img in enumerate(images):
            _add_image(doc, img)
            if i < len(images) - 1:
                doc.add_page_break()
        doc.add_page_break()

    # remove quebra final extra
    if doc.paragraphs:
        p = doc.paragraphs[-1]
        p._element.getparent().remove(p._element)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def _process_tr(tr_bytes, debug):
    images = pdf_to_images(tr_bytes)
    tr_images = []

    with pdfplumber.open(BytesIO(tr_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):

            img = images[page_index]

            # === MÁSCARA CIRÚRGICA AQUI ===
            img = apply_table_price_mask(
                image=img,
                pdf_page=page,
                debug=debug
            )

            tr_images.append(img)

    return tr_images


def _add_image(doc, image):
    bio = BytesIO()
    image.save(bio, format="JPEG", quality=90)
    bio.seek(0)
    doc.add_picture(bio, width=Cm(18))


def _configure_page(doc):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
