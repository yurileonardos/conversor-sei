from pdf2image import convert_from_bytes
import pdfplumber
from io import BytesIO


def pdf_to_images(pdf_bytes, return_pages=False):
    """
    Converte PDF em imagens.
    Se return_pages=True, retorna também as páginas do pdfplumber.
    """

    images = convert_from_bytes(pdf_bytes)

    if not return_pages:
        return images

    pdf = pdfplumber.open(BytesIO(pdf_bytes))
    pages = pdf.pages

    return images, pages
