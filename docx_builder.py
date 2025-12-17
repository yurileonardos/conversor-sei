from docx import Document
from docx.shared import Cm
from io import BytesIO

from pdf_utils import pdf_to_images
from masking_v25r import apply_masking_v25r


def build_docx(tr_bytes, proposal_files, debug=False):
    doc = Document()
    _configure_page(doc)

    # ===== TR SEMPRE PRIMEIRO =====
    tr_images = _process_tr(tr_bytes, debug)

    if not tr_images:
        raise RuntimeError("Nenhuma página gerada para o Termo de Referência.")

    for img in tr_images:
        _add_image(doc, img)
        doc.add_page_break()

    # ===== PROPOSTAS (SEM MÁSCARA) =====
    for proposal in proposal_files:
        images = pdf_to_images(proposal)
        for img in images:
            _add_image(doc, img)
            doc.add_page_break()

    # remove quebra extra
    if doc.paragraphs:
        p = doc.paragraphs[-1]
        p._element.getparent().remove(p._element)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def _process_tr(pdf_bytes, debug):
    images, pages = pdf_to_images(pdf_bytes, return_pages=True)
    processed = []

    state = {"active": False, "cut_x": None}

    for img, page in zip(images, pages):
        img, state = apply_masking_v25r(
            image=img,
            pdf_page=page,
            state=state,
            debug=debug
        )
        processed.append(img)

    return processed


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
