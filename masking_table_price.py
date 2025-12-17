def apply_table_price_mask(image, pdf_page, debug=False):
    """
    Máscara cirúrgica baseada em BLOCO REAL DE TABELA.

    - Detecta tabelas mesmo com colunas mescladas
    - Não mascara texto corrido
    - Não herda estado entre páginas
    - Só aplica máscara se houver estrutura tabular + indício de preço
    """

    import re
    from PIL import ImageDraw

    # -------------------------------------------------
    # 0. Extração de palavras
    # -------------------------------------------------
    words = pdf_page.extract_words(use_text_flow=True)
    if not words:
        return image

    img_w, img_h = image.size
    draw = ImageDraw.Draw(image, "RGBA")

    # -------------------------------------------------
    # 1. Agrupar palavras em linhas visuais
    # -------------------------------------------------
    lines = {}
    for w in words:
        # Agrupamento vertical com tolerância
        y_key = round(w["top"] / 4) * 4
        lines.setdefault(y_key, []).append(w)

    # -------------------------------------------------
    # 2. Identificar linhas TABULARES reais
    #    (colunas justapostas, não texto corrido)
    # -------------------------------------------------
    tabular_lines = []

    for y, ws in sorted(lines.items()):
        xs = sorted(w["x0"] for w in ws)

        # Poucas palavras não caracterizam tabela
        if len(xs) < 3:
            continue

        # Espaços entre "colunas"
        gaps = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
        large_gaps = [g for g in gaps if g > pdf_page.width * 0.05]

        # Regra-chave: pelo menos 2 separações claras → colunas
        if len(large_gaps) >= 2:
            tabular_lines.append((y, ws))

    # -------------------------------------------------
    # 3. Exigir continuidade vertical (bloco de tabela)
    # -------------------------------------------------
    blocks = []
    current_block = []

    for (y, ws) in tabular_lines:
        if not current_block:
            current_block = [(y, ws)]
        else:
            prev_y = current_block[-1][0]
            # Linhas próximas verticalmente
            if abs(y - prev_y) <= 10:
                current_block.append((y, ws))
            else:
                if len(current_block) >= 2:
                    blocks.append(current_block)
                current_block = [(y, ws)]

    if len(current_block) >= 2:
        blocks.append(current_block)

    # Sem bloco tabular real → NÃO mascara
    if not blocks:
        return image

    # Usa o maior bloco (tabela principal da página)
    table_block = max(blocks, key=len)

    # -------------------------------------------------
    # 4. Limites verticais REAIS da tabela
    # -------------------------------------------------
    table_top_pdf = min(y for y, _ in table_block)
    table_bottom_pdf = max(
        max(w["bottom"] for w in ws) for _, ws in table_block
    )

    # -------------------------------------------------
    # 5. Detectar coluna de preço
    # -------------------------------------------------
    price_regex = re.compile(r"\d[\d\.]*,\d{2}")
    price_x_candidates = []

    # 5.1 Por valores monetários
    for _, ws in table_block:
        for w in ws:
            if price_regex.search(w["text"]):
                price_x_candidates.append(w["x0"])

    # 5.2 Fallback por cabeçalho (páginas 1 e 2)
    if not price_x_candidates:
        for _, ws in table_block:
            for w in ws:
                t = w["text"].lower()
                if any(k in t for k in ["preço", "preco", "unitário", "unitario", "total", "r$"]):
                    price_x_candidates.append(w["x0"])

    if not price_x_candidates:
        return image

    cut_x_pdf = min(price_x_candidates)

    # -------------------------------------------------
    # 6. Converter coordenadas PDF → imagem
    # -------------------------------------------------
    cut_x_img = (cut_x_pdf / pdf_page.width) * img_w - 5
    y_top_img = (table_top_pdf / pdf_page.height) * img_h
    y_bot_img = (table_bottom_pdf / pdf_page.height) * img_h

    # Proteções finais
    cut_x_img = max(0, min(cut_x_img, img_w))
    y_top_img = max(0, min(y_top_img, img_h))
    y_bot_img = max(0, min(y_bot_img, img_h))

    # -------------------------------------------------
    # 7. Aplicar máscara SOMENTE na área da tabela
    # -------------------------------------------------
    fill_color = (255, 0, 0, 120) if debug else "white"

    draw.rectangle(
        [cut_x_img, y_top_img, img_w, y_bot_img],
        fill=fill_color,
        outline=None
    )

    # -------------------------------------------------
    # 8. Modo diagnóstico (opcional)
    # -------------------------------------------------
    if debug:
        # Contorno da tabela
        draw.rectangle(
            [0, y_top_img, img_w, y_bot_img],
            outline="blue",
            width=2
        )
        # Linha de corte
        draw.line(
            [(cut_x_img, y_top_img), (cut_x_img, y_bot_img)],
            fill="red",
            width=2
        )

    return image.convert("RGB")
