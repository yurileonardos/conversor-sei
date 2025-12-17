import streamlit as st
from docx_builder import build_docx

# -------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -------------------------------------------------
st.set_page_config(
    page_title="SEI – Conversor TR e Propostas",
    page_icon="📑",
    layout="centered"
)

# -------------------------------------------------
# TÍTULO
# -------------------------------------------------
st.title("📑 Conversor SEI – TR e Propostas de Preços")

st.markdown(
    """
    Este sistema converte **Termo de Referência (TR)** e **Propostas de Preços**
    em um **único arquivo DOCX**, pronto para inserção no **SEI**.

    🔒 *Os valores financeiros são ocultados **somente no TR***  
    📄 *As propostas são inseridas sem qualquer alteração*
    """
)

st.divider()

# -------------------------------------------------
# UPLOAD DE ARQUIVOS
# -------------------------------------------------
uploaded_files = st.file_uploader(
    label="Envie os arquivos PDF (1º TR, depois as Propostas):",
    type=["pdf"],
    accept_multiple_files=True
)

# -------------------------------------------------
# OPÇÕES
# -------------------------------------------------
debug_mode = st.checkbox(
    "Modo diagnóstico (mostrar máscara visual)",
    help="Ative apenas para conferência técnica. Não use em produção."
)

st.divider()

# -------------------------------------------------
# PROCESSAMENTO
# -------------------------------------------------
if uploaded_files:

    if len(uploaded_files) < 1:
        st.warning("Envie pelo menos o Termo de Referência.")
        st.stop()

    # REGRA INSTITUCIONAL
    # O PRIMEIRO ARQUIVO É SEMPRE O TR
    tr_file = uploaded_files[0]
    proposal_files = uploaded_files[1:]

    st.info(
        f"""
        📌 **Ordem reconhecida pelo sistema:**
        - Termo de Referência: **{tr_file.name}**
        - Propostas: **{len(proposal_files)} arquivo(s)**
        """
    )

    if st.button("🚀 Processar documentos"):

        with st.spinner("Processando documentos..."):

            try:
                tr_bytes = tr_file.read()
                proposals_bytes = [f.read() for f in proposal_files]

                # FUNÇÃO CENTRAL
                docx_output = build_docx(
                    tr_bytes=tr_bytes,
                    proposal_files=proposals_bytes,
                    debug=debug_mode
                )

                st.success("✅ Documento gerado com sucesso!")

                st.download_button(
                    label="📥 Baixar DOCX final",
                    data=docx_output,
                    file_name="TR_e_Propostas_SEI.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error("❌ Ocorreu um erro durante o processamento.")
                st.exception(e)

else:
    st.info("⬆️ Envie o Termo de Referência e, se houver, as Propostas de Preços.")

# -------------------------------------------------
# RODAPÉ
# -------------------------------------------------
st.divider()
st.caption(
    "Sistema desenvolvido para conversão institucional de documentos SEI "
    "• TR com ocultação de valores • Propostas preservadas"
)
