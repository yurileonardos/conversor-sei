import pdfplumber
from io import BytesIO
import html


def extract_tables_from_tr(pdf_bytes):
    """
    Extrai tabelas do TR preservando estrutura e conteúdo.
    Retorna lista de tabelas (listas de listas).
    """

    tables = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5
            })

            for table in page_tables:
                # limpa None e preserva texto
                clean_table = []
                for row in table:
                    clean_row = [
                        html.escape(cell.strip()) if cell else ""
                        for cell in row
                    ]
                    clean_table.append(clean_row)

                tables.append(clean_table)

    return tables
