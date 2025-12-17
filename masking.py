import re
from PIL import ImageDraw

def is_price_format(text):
    if not text:
        return False
    clean = text.strip().replace(" ", "")
    match = re.search(r'[\d\.]*,\d{2}', clean)
    if match:
        invalids = sum(1 for c in clean if c.lower() not in '0123456789.,r$()')
        return invalids <= 2
    return False


def find_x_by_visual_scan(pdf_page):
    words = pdf_page.extract_words()
    price_words = [w for w in words if is_price_format(w['text'])]
    if not price_words:
        return None

    x_clusters = []
    tolerance = 10

    for w in price_words:
        x0 = w['x0']
        for c in x_clusters:
            if abs(c['avg_x'] - x0) < tolerance:
                c['points'].append(x0)
                c['avg_x'] = sum(c['points']) / len(c['points'])
                c['count'] += 1
                break
        else:
            x_clusters.append({'avg_x': x0, 'points': [x0], 'count': 1})

    page_width = pdf_page.width
    valid = [c for c in x_clusters if c['avg_x'] > page_width * 0.45]
    if not valid:
        return None

    return min(valid, key=lambda c: c['avg_x'])['avg_x']


def find_x_by_header_scan(pdf_page):
    targets = ["unitário", "unitario", "estimado", "total", "(r$)"]
    for w in pdf_page.extract_words():
        if any(t in w['text'].lower() for t in targets):
            if w['x0'] > pdf_page.width * 0.4:
                return w['x0']
    return None


def check_for_stoppers(pdf_page):
    text = pdf_page.extract_text().lower()
    keys = [
        "local de entrega", "prazo de entrega", "sanções administrativas",
        "obrigações da contratada", "gestão do contrato", "vigência"
    ]
    return any(k in text for k in keys)


def apply_masking(image, pdf_page, mask_state, debug=False):
    if check_for_stoppers(pdf_page):
        mask_state['active'] = False
        mask_state['cut_x_percent'] = None
        return image, mask_state

    found_x = find_x_by_visual_scan(pdf_page) or find_x_by_header_scan(pdf_page)
    if found_x:
        mask_state['active'] = True
        mask_state['cut_x_percent'] = found_x / pdf_page.width

    if mask_state['active'] and mask_state['cut_x_percent']:
        draw = ImageDraw.Draw(image, "RGBA")
        cut_x = mask_state['cut_x_percent'] * image.size[0] - 5
        fill = (255, 0, 0, 100) if debug else "white"
        draw.rectangle([cut_x, 0, image.size[0], image.size[1]], fill=fill)

    return image.convert("RGB"), mask_state
