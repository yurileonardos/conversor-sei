import pdfplumber
from io import BytesIO
import html

PRICE_HEADERS = [
    "preço", "valor", "unitário", "total", "r$", "estimado"
]


def pdf_tr_to_html(pdf_bytes):
    html_blocks = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:

            # ---------- TEXTO CORRIDO ----------
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    html_blocks.append(
                        f'<p class="Texto_Justificado">{html.escape(line)}</p>'
                    )

            # ---------- TABELAS ----------
            tables = page.extract_tables()
            for table in tables:
                html_blocks.append(render_table_html(table))

    return "\n".join(html_blocks)


# 🔴 ESTA FUNÇÃO PRECISA EXISTIR NO MESMO ARQUIVO 🔴
def render_table_html(table):
    """
    Reconstrói tabela preservando estrutura
    e ocultando colunas de preço
    """

    if not table or len(table) < 2:
        return ""

    headers = table[0]
    rows = table[1:]

    price_cols = []
    for i, h in enumerate(headers):
        if h and any(k in h.lower() for k in PRICE_HEADERS):
            price_cols.append(i)

    html_table = [
        '<table border="1" style="width:100%; border-collapse:collapse;">'
    ]

    # Cabeçalho
    html_table.append("<thead><tr>")
    for i, h in enumerate(headers):
        if i not in price_cols:
            html_table.append(
                f'<th style="background-color:rgb(238,238,238); text-align:center;">'
                f'{html.escape(h or "")}</th>'
            )
    html_table.append("</tr></thead>")

    # Corpo
    html_table.append("<tbody>")
    for row in rows:
        html_table.append("<tr>")
        for i, cell in enumerate(row):
            if i not in price_cols:
                html_table.append(
                    f'<td style="text-align:center;">{html.escape(cell or "")}</td>'
                )
        html_table.append("</tr>")
    html_table.append("</tbody></table><br>")

    return "\n".join(html_table)
