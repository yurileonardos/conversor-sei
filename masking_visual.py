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


def apply_visual_mask(image, table_bbox, cut_x):
    """
    Aplica máscara sem sobrepor texto externo
    """
    pass
