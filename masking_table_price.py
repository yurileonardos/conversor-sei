def apply_table_price_mask(image, pdf_page, debug=False):
    """
    Aplica máscara APENAS nas colunas de preço,
    APENAS dentro da área real da tabela.

    - Não mascara página inteira
    - Não mascara texto fora de tabela
    - Funciona nas páginas 1 e 2 (pouca repetição)
    - Para automaticamente quando não há tabela
    """

    import re
    from PIL import ImageDraw

    # ---------------------------
    # 0. Sanidade básica
    # ---------------------------
    if image is None or pdf_page is None:
        return image

    img_width, img_height = image.size
    draw = ImageDraw.Draw(image, "RGBA")

    words = pdf_page.extract_words(use_text_flow=True)
    if not words:
        return image  # Página sem texto → nada a fazer

    # ---------------------------
    # 1. Detectar linhas "tabulares"
    # ---------------------------
    lines = {}

    for w in words:
        # Agrupa por proximidade vertical (linhas visuais)
        y_key = round(w["top"] / 4) * 4
        lines.setdefault(y_key, []).append(w)

    table_lines = []

    for y, ws in lines.items():
        xs = [w["x0"] for w in ws]

        # Critério de tabela:
        # - múltiplas colunas
        # - espalhadas horizontalmente
        if len(xs) >= 3 and (max(xs) - min(xs)) > pdf_page.width * 0.35:
            table_lines.append((y, ws))

    # Se não há indício de tabela → NÃO mascara
    if not table_lines:
        return image

    # ---------------------------
    # 2. Limites verticais reais da tabela
    # ---------------------------
    table_top_pdf = min(y for y, _ in table_lines)
    table_bottom_pdf = max(
        max(w["bottom"] for w in ws) for _, ws in table_lines
    )

    # ---------------------------
    # 3. Detectar coluna de preços
    # ---------------------------
    price_pattern = re.compile(r"\d[\d\.]*,\d{2}")
    price_x_candidates = []

    # 3.1 Detecção por valores monetários
    for _, ws in table_lines:
        for w in ws:
            if price_pattern.search(w["text"]):
                price_x_candidates.append(w["x0"])

    # 3.2 Fallback por cabeçalho (páginas 1 e 2)
    if not price_x_candidates:
        for _, ws in table_lines:
            for w in ws:
                t = w["text"].lower()
                if any(k in t for k in [
                    "preço", "preco", "unitário", "unitario",
                    "total", "(r$)", "r$"
                ]):
                    price_x_candidates.append(w["x0"])

    # Se tabela não contém preço → NÃO mascara
    if not price_x_candidates:
        return image

    # Coluna de preço mais à esquerda
    cut_x_pdf = min(price_x_candidates)

    # ---------------------------
    # 4. Converter coordenadas PDF → Imagem
    # ---------------------------
    cut_x_img = (cut_x_pdf / pdf_page.width) * img_width - 5

    table_top_img = (table_top_pdf / pdf_page.height) * img_height
    table_bottom_img = (table_bottom_pdf / pdf_page.height) * img_height

    # Proteções finais
    cut_x_img = max(0, min(cut_x_img, img_width))
    table_top_img = max(0, min(table_top_img, img_height))
    table_bottom_img = max(0, min(table_bottom_img, img_height))

    # ---------------------------
    # 5. Aplicar máscara (somente na tabela)
    # ---------------------------
    fill_color = (255, 0, 0, 120) if debug else "white"

    draw.rectangle(
        [
            cut_x_img,
            table_top_img,
            img_width,
            table_bottom_img
        ],
        fill=fill_color,
        outline=None
    )

    # ---------------------------
    # 6. Modo diagnóstico (opcional)
    # ---------------------------
    if debug:
        # contorno da tabela
        draw.rectangle(
            [
                0,
                table_top_img,
                img_width,
                table_bottom_img
            ],
            outline="blue",
            width=2
        )

        # linha de corte
        draw.line(
            [
                (cut_x_img, table_top_img),
                (cut_x_img, table_bottom_img)
            ],
            fill="red",
            width=2
        )

    return image.convert("RGB")
