import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile
import pdfplumber
from PIL import ImageDraw
import re

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
            .stFileUploader {
                padding: 20px;
                border-radius: 10px;
                border: 2px dashed #cccccc;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- TÍTULO PRINCIPAL ---
st.title("📑 SEI Converter ATA - SGB")

# --- INTRODUÇÃO ---
st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

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

with st.expander("⚙️ Deseja escolher a pasta onde o arquivo será salvo? Clique aqui."):
    st.markdown("""
    Por segurança, os navegadores salvam automaticamente na pasta "Downloads". 
    Para escolher a pasta a cada download, configure seu navegador (Chrome/Edge):
    1. Vá em **Configurações** > **Downloads**.
    2. Ative: **"Perguntar onde salvar cada arquivo antes de fazer download"**.
    """)

st.write("---")

# --- PASSO 1: UPLOAD ---
st.write("### Passo 1: Upload dos Arquivos")
st.markdown("**Nota:** O sistema aplicará a máscara automática para ocultar preços em Termos de Referência.")

uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui (ou clique para buscar):", 
    type="pdf", 
    accept_multiple_files=True
)

# --- FUNÇÃO AUXILIAR DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    # Substitui quebras de linha por espaço, remove pontuação extra e poe minusculo
    text = text.replace('\n', ' ').replace('\r', ' ')
    return text.lower().strip()

# --- FUNÇÃO DE MASCARAMENTO (CALIBRADA v8.0) ---
def apply_masking(image, pdf_page):
    try:
        # Tenta estratégia de LINHAS (comum em governo)
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        
        # Se não achar nada, tenta estratégia de TEXTO (para tabelas sem borda)
        if not tables:
             tables = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

        draw = ImageDraw.Draw(image)
        im_width, im_height = image.size
        
        scale_x = im_width / pdf_page.width
        scale_y = im_height / pdf_page.height

        # Palavras-chave Agressivas (com e sem acento)
        keywords_target = [
            "preco unit", "preço unit", "valor unit", "vlr. unit", "unitario", "unitário",
            "valor max", "valor estim", "preço estim", "preco estim", "valor ref", 
            "vlr total", "valor total", "preco total", "preço total"
        ]
        
        keyword_total_row = "total"

        for table in tables:
            if not table.rows: continue

            # --- 1. LOCALIZAR CABEÇALHO ---
            # Vamos varrer as primeiras 3 linhas para achar o cabeçalho (as vezes tem titulo antes)
            header_found_idx = -1
            mask_start_x = None

            for row_idx in range(min(3, len(table.rows))):
                row_cells = table.rows[row_idx].cells
                
                # Verifica cada célula dessa linha
                for cell_idx, cell in enumerate(row_cells):
                    if not cell: continue
                    
                    try:
                        # Extrai texto da área
                        cropped = pdf_page.crop(cell)
                        text_raw = cropped.extract_text()
                        text_clean = clean_text(text_raw)
                        
                        # Verifica se bate com as palavras chave
                        if any(k in text_clean for k in keywords_target):
                            # BINGO! Achamos a coluna de preço
                            mask_start_x = cell[0] # Pega a coordenada X da esquerda
                            header_found_idx = row_idx
                            break 
                    except:
                        pass
                
                if mask_start_x is not None:
                    break # Para de procurar cabeçalho se já achou
            
            # --- 2. APLICAR MÁSCARA VERTICAL (COLUNAS) ---
            if mask_start_x is not None:
                # Desenha um retângulo branco do inicio da coluna encontrada até o fim da tabela
                table_rect = table.bbox
                rect = [
                    mask_start_x * scale_x,       # Começa na coluna de preço
                    table_rect[1] * scale_y,      # Topo da tabela
                    table_rect[2] * scale_x + 50, # Fim da tabela (soma +50px pra garantir borda)
                    table_rect[3] * scale_y       # Fundo da tabela
                ]
                draw.rectangle(rect, fill="white", outline="white")

            # --- 3. APLICAR MÁSCARA HORIZONTAL (TOTAL) ---
            # Verifica a última linha
            last_row = table.rows[-1]
            try:
                # Pega texto da linha inteira (ou primeira celula)
                first_cell = last_row.cells[0]
                if first_cell:
                    cropped_last = pdf_page.crop(first_cell)
                    last_text = clean_text(cropped_last.extract_text())
                    
                    if "total" in last_text:
                        # Pega bbox da linha (usando as células para calcular altura)
                        tops = [c[1] for c in last_row.cells if c]
                        bottoms = [c[3] for c in last_row.cells if c]
                        
                        if tops and bottoms:
                            l_top = min(tops)
                            l_bottom = max(bottoms)
                            
                            rect_total = [
                                table.bbox[0] * scale_x, # Esquerda da tabela
                                l_top * scale_y,         # Topo da linha
                                table.bbox[2] * scale_x, # Direita da tabela
                                l_bottom * scale_y       # Fundo da linha
                            ]
                            draw.rectangle(rect_total, fill="white", outline="white")
            except:
                pass

    except Exception as e:
        print(f"Erro mascaramento: {e}")
        pass
    
    return image

# --- FUNÇÃO DE CONVERSÃO ---
def convert_pdf_to_docx(file_bytes):
    # Tenta abrir com pdfplumber para ler texto
    try:
        pdf_plumb = pdfplumber.open(BytesIO(file_bytes))
        has_text_layer = True
    except:
        has_text_layer = False
        pdf_plumb = None

    images = convert_from_bytes(file_bytes)
    doc = Document()
    
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)

    for i, img in enumerate(images):
        # Tenta mascarar se tiver camada de texto
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img = apply_masking(img, pdf_plumb.pages[i])
        
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

# --- PASSO 2: CONVERTER E DOWNLOAD ---
if uploaded_files:
    st.write("---")
    st.write("### Passo 2: Converter e Download")
    
    qtd = len(uploaded_files)
    st.caption(f"{qtd} arquivo(s) pronto(s) para conversão.")

    if st.button(f"🚀 Processar Arquivos"):
        with st.spinner('Aplicando máscaras de sigilo e convertendo...'):
            try:
                processed_files = []
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    progress_bar.progress((index + 1) / qtd)

                st.success("✅ Conversão concluída!")
                
                if len(processed_files) == 1:
                    name, data = processed_files[0]
                    st.download_button(
                        label=f"📥 Salvar {name} no Computador",
                        data=data,
                        file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_download"
                    )
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, data in processed_files:
                            zf.writestr(name, data.getvalue())
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 Salvar Todos (.ZIP) no Computador",
                        data=zip_buffer,
                        file_name="Documentos_SEI_Convertidos.zip",
                        mime="application/zip",
                        key="btn_download_zip"
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

st.warning("⚠️ **IMPORTANTE:** Certifique-se de deixar todas as caixas de seleção **DESMARCADAS**.")

try:
    st.image("print_sei.png", caption="Exemplo: Deixe as opções desmarcadas.", use_container_width=True)
except:
    pass

# --- RODAPÉ ---
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v8.0</div>', unsafe_allow_html=True)
