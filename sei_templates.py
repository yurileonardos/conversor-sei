def render_table_sei(table):
    """
    Renderiza tabela no padrão HTML SEI,
    preservando layout e sem reinterpretação.
    """

    html = []
    html.append('<table style="width:100%; border-collapse:collapse;" border="1">')

    for i, row in enumerate(table):
        html.append("<tr>")
        for cell in row:
            if i == 0:
                html.append(
                    f'<th style="background-color:rgb(238,238,238); text-align:center;">{cell}</th>'
                )
            else:
                html.append(
                    f'<td style="text-align:center;">{cell}</td>'
                )
        html.append("</tr>")

    html.append("</table>")
    html.append("<br>")

    return "\n".join(html)
