import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Conversor SEI",
    page_icon="📑",
    layout="centered"
)

# --- ESTILO CSS PARA O RODAPÉ ---
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
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- TÍTULO E CABEÇALHO ---
st.title("📑 Conversor PDF para Word (Padrão SEI)")
st.write("Converta documentos PDF em imagens otimizadas para o Sistema SEI, evitando erros de tamanho.")

st.info("💡 **Como funciona:** O sistema redimensiona cada página para 552x781px, centraliza no Word e reduz o peso do arquivo.")

# --- UPLOAD ---
uploaded_file = st.file_uploader("Arraste seu PDF aqui", type="pdf")

if uploaded_file is not None:
    # Botão de ação
    if st.button("🚀 Iniciar Conversão"):
        with st.spinner('Processando... Por favor, aguarde.'):
            try:
                # 1. Converter PDF em imagens
                images = convert_from_bytes(uploaded_file.read())
                
                # 2. Criar documento Word
                doc = Document()
                
                # Configurar margens A4
                section = doc.sections[0]
                section.page_height = Cm(29.7)
                section.page_width = Cm(21.0)
                section.left_margin = Cm(1.0)
                section.right_margin = Cm(1.0)
                section.top_margin = Cm(1.0)
                section.bottom_margin = Cm(1.0)

                total_pages = len(images)
                
                # Barra de progresso
                progress_bar = st.progress(0)

                for i, img in enumerate(images):
                    # Redimensionamento SEI (Otimizado)
                    img = img.resize((552, 781))
                    
                    img_byte_arr = BytesIO()
                    img.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
                    img_byte_arr.seek(0)

                    # Adicionar quebra de página se não for a primeira
                    if i > 0:
                        doc.add_page_break()

                    # Inserir imagem e capturar o parágrafo
                    doc.add_picture(img_byte_arr, width=Cm(19.0))
                    
                    # --- CENTRALIZAR IMAGEM ---
                    last_paragraph = doc.paragraphs[-1] 
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Atualizar barra
                    progress_bar.progress((i + 1) / total_pages)

                # 3. Preparar Download
                docx_io = BytesIO()
                doc.save(docx_io)
                docx_io.seek(0)

                st.success("✅ Conversão concluída com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_io,
                    file_name=f"{uploaded_file.name}_SEI_Yuri.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"Erro ao processar: {e}")

# --- RODAPÉ PERSONALIZADO ---
st.markdown('<div class="footer">Developed by Yuri 🚀 | Otimizador SEI v1.0</div>', unsafe_allow_html=True)