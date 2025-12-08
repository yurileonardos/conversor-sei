import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="SEI Converter ATA - SGB",
    page_icon="📑",
    layout="centered"
)

# --- ESTILO CSS ---
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            .footer {
                position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                background-color: #f1f1f1;
                color: #555;
                text-align: center;
                padding: 10px;
                font-size: 14px;
                z-index: 999;
            }
            .stFileUploader label {
                 font-size: 18px;
                 font-weight: bold;
            }
            /* Aumenta a área de upload visualmente para facilitar o "arrastar" */
            .stFileUploader {
                padding: 20px;
                border-radius: 10px;
                border: 2px dashed #cccccc;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- (LOGO OPCIONAL) ---
# st.image("logo.png", width=200) 

# --- TÍTULO PRINCIPAL ---
st.title("📑 SEI Converter ATA - SGB")

# --- INTRODUÇÃO ---
st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# Aviso sobre a extensão e links
col1, col2 = st.columns([0.1, 0.9])
with col1:
    try:
        st.image("icone_sei.png", width=40)
    except:
        st.write("🧩")
with col2:
    st.info("""
    Funcionalidade disponível na extensão [**SEI PRO**](https://sei-pro.github.io/sei-pro/), 
    utilizando a ferramenta [**INSERIR CONTEÚDO EXTERNO**](https://sei-pro.github.io/sei-pro/pages/INSERIRDOC.html).
    """)

st.write("---")

# --- PASSO 1: UPLOAD (COM TEXTO EXPLÍCITO DE ARRASTAR) ---
st.write("### Passo 1: Upload dos Arquivos")

uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui (ou clique para buscar):", 
    type="pdf", 
    accept_multiple_files=True
)

# --- FUNÇÃO DE CONVERSÃO ---
def convert_pdf_to_docx(file_bytes):
    images = convert_from_bytes(file_bytes)
    doc = Document()
    
    # Margens
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)

    for i, img in enumerate(images):
        img = img.resize((552, 781))
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
        img_byte_arr.seek(0)

        if i > 0:
            doc.add_page_break()

        doc.add_picture(img_byte_arr, width=Cm(19.0))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    docx_io = BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io

# --- PASSO 2: CONVERTER E DOWNLOAD (UNIFICADOS) ---
if uploaded_files:
    st.write("---")
    st.write("### Passo 2: Converter e Download")
    
    qtd = len(uploaded_files)
    st.caption(f"{qtd} arquivo(s) na fila para processamento.")

    # Botão Único de Ação
    if st.button(f"🚀 Processar Arquivos e Gerar Download"):
        with st.spinner('Otimizando imagens e gerando documentos...'):
            try:
                processed_files = []
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    progress_bar.progress((index + 1) / qtd)

                # Sucesso
                st.success("✅ Conversão finalizada!")
                
                # Exibição IMEDIATA do botão de download
                if len(processed_files) == 1:
                    name, data = processed_files[0]
                    st.download_button(
                        label=f"📥 Clique aqui para Baixar {name}",
                        data=data,
                        file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, data in processed_files:
                            zf.writestr(name, data.getvalue())
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 Clique aqui para Baixar Todos (.ZIP)",
                        data=zip_buffer,
                        file_name="Documentos_SEI_Convertidos.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"Erro ao processar: {e}")

# --- GUIA VISUAL ---
st.write("---")
st.subheader("📚 Guia Rápido: Como inserir no SEI")

col1, col2 = st.columns([0.15, 0.85])
with col1:
    try:
        st.image("icone_sei.png", width=50) 
    except:
        st.write("🧩")
with col2:
    st.markdown("""
    **1º Localize o ícone:** No editor do SEI, clique no botão da função **INSERIR CONTEÚDO EXTERNO** (representado pelo ícone ao lado).
    """)

st.write("")

st.markdown("""
**2º Configure a inserção:** Na janela que abrir, faça o upload do arquivo Word gerado aqui.
""")

st.warning("⚠️ **IMPORTANTE:** Certifique-se de deixar todas as caixas de seleção **DESMARCADAS** (como na imagem abaixo) para evitar que o arquivo substitua o conteúdo existente no documento.")

try:
    st.image("print_sei.png", caption="Exemplo: Deixe as opções desmarcadas.", use_container_width=True)
except:
    pass

# --- RODAPÉ ---
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v4.0</div>', unsafe_allow_html=True)
