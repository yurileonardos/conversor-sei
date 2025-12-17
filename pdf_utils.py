import fitz  # PyMuPDF
from PIL import Image
from io import BytesIO


def pdf_to_images(pdf_bytes, dpi=200):
    """
    Converte PDF em lista de imagens PIL
    Compatível com Streamlit Cloud (sem poppler)
    """
    images = []

    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)

    for page in pdf:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")
        images.append(img)

    return images
