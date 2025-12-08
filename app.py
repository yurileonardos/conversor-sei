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
                 font-weight: bold;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- (LOGO OPCIONAL) ---
# st.image("logo.png", width=200) 

# --- TÍTULO PRINCIPAL ---
st.title("📑 SEI Converter ATA - SGB")

# --- INSTRUÇÕES ---
st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: [**ATA DE REGISTRO DE PREÇOS**](https://sei-pro.github.io/sei-pro/).
""")

col1, col2 = st.columns([0.1, 0.9])
with col1:
    try:
        st.image("icone_sei.png", width=40)
    except:
        st.write("🧩")
with col2:
    st.info("""
    Vale ressaltar que essa funcionalidade de inserir o DOCX gerado só está presente na extensão **SEI PRO**, 
    utilizando a função **"Inserir Conteúdo Externo"**.
    """)

st.write("---")

# --- UPLOAD DE MÚLTIPLOS ARQUIVOS ---
st.write("### Passo 1: Upload dos Arquivos")
# AQUI ESTÁ A MUDANÇA: accept_multiple_files=True
uploaded_files = st.file_uploader(
    "Selecione ou arraste seus arquivos PDF (um ou vários):", 
    type="pdf", 
    accept_multiple_files=True
)

# FUNÇÃO DE CONVERSÃO (Para organizar o código)
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
    
    # Salvar em memória
    docx_io = BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io

# --- PROCESSAMENTO ---
if uploaded_files:
    st.write("---")
    st.write("### Passo 2: Conversão")
    
    # Mostra quantos arquivos foram carregados
    qtd = len(uploaded_files)
    st.caption(f"{qtd} arquivo(s) selecionado(s).")

    if st.button(f"🚀 Iniciar Conversão ({qtd} arquivos)"):
        with st.spinner('Processando... Isso pode levar alguns instantes.'):
            try:
                # Lista para guardar os resultados processados
                processed_files = []
                
                # Barra de progresso geral
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    # Processa cada arquivo
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    
                    # Atualiza barra
                    progress_bar.progress((index + 1) / qtd)

                st.success("✅ Todos os arquivos foram convertidos!")
                st.write("### Passo 3: Download")
                st.caption("O local de salvamento depende das configurações do seu navegador.")

                # LÓGICA DE DOWNLOAD INTELIGENTE
                if len(processed_files) == 1:
                    # Se for só 1 arquivo, baixa o DOCX direto
                    name, data = processed_files[0]
                    st.download_button(
                        label=f"📥 Baixar {name}",
                        data=data,
                        file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    # Se forem vários, cria um ZIP
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, data in processed_files:
                            zf.writestr(name, data.getvalue())
                    
                    zip_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Baixar Todos (Arquivo .ZIP)",
                        data=zip_buffer,
                        file_name="Documentos_SEI_Convertidos.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"Erro ao processar: {e}")

# --- RODAPÉ ---
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v2.0</div>', unsafe_allow_html=True)
