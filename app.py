import streamlit as st
from pdf_to_html_tr import pdf_tr_to_html

st.set_page_config(page_title="Conversor SEI – TR", layout="centered")

st.title("📄 Conversor TR → HTML (SEI)")

# 🔴 AQUI VOCÊ DEFINE tr_file
tr_file = st.file_uploader(
    "Envie o Termo de Referência (PDF)",
    type=["pdf"],
    accept_multiple_files=False
)

# 🔴 SÓ DEPOIS VOCÊ USA
if tr_file is not None:
    with st.spinner("Convertendo TR para HTML tabulado..."):
        tr_html = pdf_tr_to_html(tr_file.read())

    st.success("Conversão concluída.")

    st.markdown("### Prévia do TR (HTML)")
    st.components.v1.html(tr_html, height=700, scrolling=True)

    st.download_button(
        "📥 Baixar HTML para SEI",
        tr_html,
        file_name="TR_SEI.html",
        mime="text/html"
    )
