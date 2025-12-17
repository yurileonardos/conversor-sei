from pdf_to_html_tr import pdf_tr_to_html
if tr_file:
    tr_html = pdf_tr_to_html(tr_file.read())

    st.components.v1.html(tr_html, height=700, scrolling=True)

    st.download_button(
        "📥 Baixar HTML (TR – SEI)",
        tr_html,
        file_name="TR_SEI.html",
        mime="text/html"
    )
