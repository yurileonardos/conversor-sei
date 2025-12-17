from PIL import Image, ImageDraw
import numpy as np


# -------------------------------------------------
# PARÂMETROS AJUSTÁVEIS (SEGUROS)
# -------------------------------------------------
DENSITY_THRESHOLD = 0.15     # sensibilidade para detectar tabela
MIN_TABLE_HEIGHT = 80        # px mínimos para considerar tabela válida
RIGHT_CUT_RATIO = 0.70       # início das duas últimas colunas (ajustável)


# -------------------------------------------------
# FUNÇÃO PRINCIPAL
# -------------------------------------------------
def apply_visual_mask(image, debug=False):
    """
    Aplica máscara SOMENTE na região da tabela.
    Se não detectar tabela, retorna a imagem original.
    """

    # 1️⃣ Detectar região vertical da tabela
    table_bbox = _detect_table_vertical_region(image)

    # Se não houver tabela → NÃO mascara
    if table_bbox is None:
        return image

    y_start, y_end = table_bbox

    # Validação de segurança
    if (y_end - y_start) < MIN_TABLE_HEIGHT:
        return image

    # 2️⃣ Calcular início das colunas de preço (X)
    width, height = image.size
    cut_x = int(width * RIGHT_CUT_RATIO)

    # 3️⃣ Aplicar máscara SOMENTE na área da tabela
    draw = ImageDraw.Draw(image, "RGBA")

    if debug:
        fill = (255, 0, 0, 120)     # vermelho translúcido
    else:
        fill = (255, 255, 255, 255)  # branco (produção)

    draw.rectangle(
        [cut_x, y_start, width, y_end],
        fill=fill
    )

    return image


# -------------------------------------------------
# DETECÇÃO VERTICAL DA TABELA
# -------------------------------------------------
def _detect_table_vertical_region(image):
    """
    Detecta início e fim vertical da tabela usando densidade gráfica.
    Retorna (y_start, y_end) ou None.
    """

    # Converte para escala de cinza
    gray = image.convert("L")
    arr = np.array(gray)

    height, width = arr.shape

    # Normaliza (0 = branco, 1 = preto)
    norm = 1 - (arr / 255.0)

    # Calcula densidade horizontal por linha
    row_density = np.mean(norm, axis=1)

    # Linhas consideradas "tabulares"
    table_rows = row_density > DENSITY_THRESHOLD

    if not np.any(table_rows):
        return None

    # Encontrar blocos contínuos
    indices = np.where(table_rows)[0]

    # Agrupar regiões contínuas
    regions = _group_contiguous(indices)

    # Selecionar a maior região (tabela principal)
    largest = max(regions, key=lambda r: r[1] - r[0])

    y_start, y_end = largest

    return y_start, y_end


def _group_contiguous(indices):
    """
    Agrupa índices contínuos.
    Retorna lista de tuplas (start, end).
    """
    regions = []
    start = indices[0]
    prev = indices[0]

    for idx in indices[1:]:
        if idx == prev + 1:
            prev = idx
        else:
            regions.append((start, prev))
            start = idx
            prev = idx

    regions.append((start, prev))
    return regions
