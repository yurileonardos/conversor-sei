from PIL import ImageDraw
import re


# ==========================
# DETECÇÃO DE PREÇO
# ==========================

def is_price_format(text: str) -> bool:
    """
    Detecta valores monetários com duas casas decimais.
    Ex: 1.234,56 | 100,00 | 0,50
    """
    if not text:
        return False

    clean = text.strip().replace(" ", "")
    match = re.search(r'\d{1,3}(\.\d{3})*,\d{2}', clean)

    if not match:
        return False

    # elimina textos com muito ruído
    invalid = sum(1 for c in clean if c.lower() not in '0123456789.,r$()')
    return invalid <= 2


# ==========================
# SCAN VISUAL (NÚMEROS)
# ==========================

def find_x_by_visual_scan(pdf_page):
    """
    Detecta alinhamento vertical de valores monetários.
    Usado principalmente em páginas sem cabeçalho.
    """
    words = pdf_page.extract_words()
    prices = [w for w in words if is_price_format(w.get("text"))]

    if not prices:
        return None

    clusters = []
    tolerance = 8

    for w in prices:
        x = w["x0"]
        placed = False

        for c in clusters:
            if abs(c["x"] - x) <= tolerance:
                c["count"] += 1
                placed = True
                break

        if not placed:
            clusters.append({"x": x, "count": 1})

    # apenas colunas da metade direita
    page_width = pdf_page.width
    valid = [c for c in clusters if c["x"] > page_width * 0.5]

    if not valid:
        return None

    # coluna mais à esquerda entre as válidas
    return min(valid, key=lambda c: c["x"])["x"]


# ==========================
# SCAN DE CABEÇALHO (PÁG 1)
# ==========================

def find_x_by_header_scan(pdf_page):
    words = pdf_page.extract_words()

    headers = [
        "unitário", "unitario",
        "preço unitário", "preco unitario",
        "valor unitário", "valor total",
        "total (r$", "r$"
    ]

    candidates = []

    for w in words:
        txt = w["text"].lower()
        if any(h in txt for h in headers):
            if w["x0"] > pdf_page.width * 0.45:
                candidates.append(w["x0"])

    return min(candidates) if candidates else None


# ==========================
# TEXTO JURÍDICO (STOP)
# ==========================

def has_legal_text(pdf_page):
    text = (pdf_page.extract_text() or "").lower()

    blockers = [
        "prazo de entrega", "local de entrega",
        "garantia", "sanções administrativas",
        "obrigações da contratada",
        "fiscalização", "gestão do contrato",
        "vigência", "cláusula", "dotação orçamentária"
    ]

    return any(b in text for b in blockers)


# ==========================
# FUNÇÃO PRINCIPAL (V25-R)
# ==========================

def apply_masking_v25r(image, pdf_page, state, debug=False):
    """
    state = {
        "active": False,
        "cut_x": None
    }
    """

    # 1️⃣ Texto jurídico encerra máscara
    if has_legal_text(pdf_page):
        state["active"] = False
        state["cut_x"] = None
        return image, state

    # 2️⃣ Tenta detectar preços
    cut_x = find_x_by_visual_scan(pdf_page)

    if cut_x is None:
        cut_x = find_x_by_header_scan(pdf_page)

    # 3️⃣ Atualiza estado
    if cut_x:
        state["active"] = True
        state["cut_x"] = cut_x
    else:
        # sem preço → encerra
        state["active"] = False
        state["cut_x"] = None
        return image, state

    # 4️⃣ Aplica máscara
    draw = ImageDraw.Draw(image, "RGBA")
    img_width, img_height = image.size

    x_start = int(state["cut_x"] / pdf_page.width * img_width) - 4

    fill = (255, 0, 0, 90) if debug else (255, 255, 255, 255)

    draw.rectangle(
        [x_start, 0, img_width, img_height],
        fill=fill,
        outline=None
    )

    return image, state
