from docx import Document
from docx.shared import Cm
from io import BytesIO

from pdf_utils import pdf_to_images
from masking_v25r import apply_masking_v25r



def build_docx(tr_bytes, proposal_files, debug=False):
    """
    Garante:
    - TR sempre primeiro
    - Máscara apenas no TR
    - Propostas nunca mascaradas
    """

    doc = Document()

    _configure_page(doc)

    # ==========================
    # PROCESSA TERMO DE REFERÊNCIA
    # ==========================
    tr_images = _process_tr(tr_bytes, debug)

    state = {"active": False, "cut_x": None}

    for page, img in zip(pdf_pages, images):
        img, state = apply_masking_v25r(img, page, state, debug=false)
        processed.append(img)

    
    if not tr_images:
        raise RuntimeError("Nenhuma página gerada para o Termo de Referência.")

    for img in tr_images:
        _add_image(doc, img)
        doc.add_page_break()

    # ==========================
    # PROCESSA PROPOSTAS
    # ==========================
    for proposal_bytes in proposal_files:
        proposal_images = pdf_to_images(proposal_bytes)

        for img in proposal_images:
            _add_image(doc, img)
            doc.add_page_break()

    # Remove quebra final extra
    if doc.paragraphs:
        doc.paragraphs[-1]._element.getparent().remove(
            doc.paragraphs[-1]._element
        )

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ==========================
# FUNÇÕES AUXILIARES
# ==========================

def _process_tr(pdf_bytes, debug):
    images = pdf_to_images(pdf_bytes)
    processed = []

   state = {"active": False, "cut_x": None}

for pdf_page, img in zip(pdf_pages, images):
    img, state = apply_masking_v25r(
        image=img,
        pdf_page=pdf_page,
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
