def detect_table_region(image):
    """
    Retorna bounding box da tabela:
    (x_start, y_start, x_end, y_end)
    """
    pass


def detect_last_two_columns(table_bbox):
    """
    Retorna x_start da área a mascarar
    """
    pass


from PIL import ImageDraw

def apply_visual_mask(image, debug=False):
    """
    Aplica mascaramento visual nas duas últimas colunas da tabela.
    O parâmetro debug controla a visualização da máscara.
    """
    draw = ImageDraw.Draw(image, "RGBA")

    # ⚠️ IMPLEMENTAÇÃO ATUAL SIMPLIFICADA
    # (exemplo de placeholder visual)
    width, height = image.size
    cut_x = int(width * 0.70)

    if debug:
        fill = (255, 0, 0, 120)  # vermelho translúcido
    else:
        fill = (255, 255, 255, 255)  # branco (produção)

    draw.rectangle(
        [cut_x, 0, width, height],
        fill=fill
    )

    return image

