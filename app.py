import streamlit as st
from docx_builder import build_docx

st.set_page_config("SEI Converter ATA - SGB", "📑")

st.title("📑 SEI Converter ATA - SGB")

uploaded_files = st.file_uploader(
    "Envie o TR (1º) e as Propostas (demais):",
    type="pdf",
    accept_multiple_files=True
)

DEBUG = st.checkbox("Modo diagnóstico (máscara vermelha)")

if uploaded_files and st.button("🚀 Processar"):
    tr = uploaded_files[0].read()
    proposals = [f.read() for f in uploaded_files[1:]]

    with st.spinner("Processando documentos..."):
        docx = build_docx(tr, proposals, DEBUG)

    st.download_button(
        "📥 Baixar DOCX Final",
        docx,
        "TR_e_Propostas_SEI.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
