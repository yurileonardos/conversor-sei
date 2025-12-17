from PIL import Image, ImageDraw
import numpy as np


# ===============================
# CONFIGURAÇÕES INSTITUCIONAIS
# ===============================

# Percentuais seguros para fallback (TR)
TR_TABLE_Y_START = 0.22   # abaixo do cabeçalho textual
TR_TABLE_Y_END   = 0.82   # antes de rodapés / texto jurídico

# Percentual horizontal onde começam preços (últimas colunas)
PRICE_COLUMN_X_START = 0.62

# Limite mínimo de densidade para considerar "região tabular"
TABLE_PIXEL_DENSITY_THRESHOLD = 0.018


# ===============================
# DETECÇÃO VISUAL DE TABELA
# ===============================

def _detect_table_vertical_region(image: Image.Image):
    """
    Detecta verticalmente a região provável de uma tabela,
    analisando densidade de pixels escuros por linha.
    Retorna (y_start, y_end) ou None.
    """

    gray = image.convert("L")
    arr = np.array(gray)

    height, width = arr.shape
    threshold = 180  # pixels escuros
    rows_density = []

    for y in range(height):
        dark_pixels = np.sum(arr[y] < threshold)
        density = dark_pixels / width
        rows_density.append(density)

    candidates = [i for i, d in enumerate(rows_density) if d > TABLE_PIXEL_DENSITY_THRESHOLD]

    if not candidates:
        return None

    y_start = min(candidates)
    y_end   = max(candidates)

    # proteção contra falso positivo (linha isolada)
    if (y_end - y_start) < height * 0.15:
        return None

    return y_start, y_end


# ===============================
# FUNÇÃO PRINCIPAL (V3)
# ===============================

def apply_visual_mask(image: Image.Image, debug: bool = False):
    """
    Aplica máscara APENAS na área de tabela do TR.
    Nunca mascara página inteira.
    Nunca vaza para páginas sem tabela.
    """

    width, height = image.size

    # 1️⃣ Tenta detecção visual real
    table_bbox = _detect_table_vertical_region(image)

    # 2️⃣ Define área vertical de máscara
    if table_bbox:
        y_start, y_end = table_bbox
    else:
        # FALLBACK INSTITUCIONAL (TR)
        y_start = int(height * TR_TABLE_Y_START)
        y_end   = int(height * TR_TABLE_Y_END)

    # Trava de segurança: não mascarar páginas essencialmente textuais
    if (y_end - y_start) < height * 0.20:
        return image

    # 3️⃣ Define área horizontal (colunas de preço)
    x_start = int(width * PRICE_COLUMN_X_START)

    # 4️⃣ Aplica máscara
    draw = ImageDraw.Draw(image, "RGBA")

    if debug:
        fill_color = (255, 0, 0, 90)   # vermelho translúcido
        line_color = (255, 0, 0, 255)
    else:
        fill_color = (255, 255, 255, 255)
        line_color = None

    draw.rectangle(
        [x_start, y_start, width, y_end],
        fill=fill_color,
        outline=None
    )

    # Linha guia apenas em debug
    if debug:
        draw.line(
            [(x_start, y_start), (x_start, y_end)],
            fill=line_color,
            width=3
        )

    return image.convert("RGB")
