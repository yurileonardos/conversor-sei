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
