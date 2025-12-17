from pdf2image import convert_from_bytes
from io import BytesIO
import pdfplumber


def pdf_to_images(pdf_bytes, dpi=200):
    """
    Converte PDF (bytes) em lista de imagens PIL.
    Uso único e seguro para Streamlit Cloud.
    """

    images = convert_from_bytes(
        pdf_bytes,
        dpi=dpi
    )

    return images


def pdf_to_images_with_pages(pdf_bytes, dpi=200):
    """
    Converte PDF em imagens PIL e retorna também
    as páginas do pdfplumber (mesma ordem).
    """

    images = convert_from_bytes(
        pdf_bytes,
        dpi=dpi
    )

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        pages = list(pdf.pages)

    return images, pages
