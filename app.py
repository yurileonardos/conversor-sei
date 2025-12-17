import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile
import pdfplumber
from PIL import ImageDraw

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
st.markdown("**Nota:** O sistema detectará automaticamente tabelas no TR e ocultará as colunas de valores estimados e a linha de total.")

uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui (ou clique para buscar):", 
    type="pdf", 
    accept_multiple_files=True
)

# --- FUNÇÃO DE MASCARAMENTO (CENÁRIO A) ---
def apply_masking(image, pdf_page):
    """
    Recebe a imagem da página e o objeto da página do pdfplumber.
    Identifica colunas de valor e linha de total e desenha retângulos brancos sobre elas.
    """
    try:
        # Encontra tabelas na página
        tables = pdf_page.find_tables()
        
        draw = ImageDraw.Draw(image)
        im_width, im_height = image.size
        
        # Fator de escala (PDF Points -> Imagem Pixels)
        scale_x = im_width / pdf_page.width
        scale_y = im_height / pdf_page.height

        # Palavras-chave para identificar colunas a remover
        keywords_remove = ["valor unitário", "vlr. unit", "unitário", "valor total", "vlr. total", "total estimado"]

        for table in tables:
            # Pega os dados da tabela (linhas e colunas)
            data = table.extract()
            if not data:
                continue

            # 1. Identificar quais índices de coluna contêm os cabeçalhos de valor
            header_row = data[0]
            indices_to_mask = []
            
            # Normaliza o texto para minusculo para comparar
            for idx, cell_text in enumerate(header_row):
                if cell_text and any(k in cell_text.lower() for k in keywords_remove):
                    indices_to_mask.append(idx)
            
            # Se não achou colunas de valor, pula essa tabela
            if not indices_to_mask:
                continue

            # Obter as coordenadas das células da tabela
            # table.rows é uma lista de objetos Row
            
            # --- MASCARAR AS COLUNAS ---
            # Vamos iterar sobre todas as células das colunas identificadas
            for row in table.rows:
                for col_idx in indices_to_mask:
                    cell = row.cells[col_idx]
                    if cell:
                        # Coordenadas do pdfplumber (x0, top, x1, bottom)
                        x0, top, x1, bottom = cell
                        
                        # Converter para pixels da imagem
                        rect = [
                            x0 * scale_x,
                            top * scale_y,
                            x1 * scale_x,
                            bottom * scale_y
                        ]
                        # Desenha retangulo branco (preenchido)
                        draw.rectangle(rect, fill="white", outline="white")

            # --- MASCARAR LINHA DE TOTAL ---
            # Verifica a última linha da tabela
            last_row_index = len(data) - 1
            last_row_text = data[last_row_index]
            
            # Se a primeira célula da última linha contiver "Total"
            first_cell_text = str(last_row_text[0]).lower() if last_row_text[0] else ""
            
            if "total" in first_cell_text:
                # Mascarar a linha inteira
                row_obj = table.rows[last_row_index]
                for cell in row_obj.cells:
                    if cell:
                        x0, top, x1, bottom = cell
                        rect = [
                            x0 * scale_x,
                            top * scale_y,
                            x1 * scale_x, # Vai até o fim da célula
                            bottom * scale_y
                        ]
                        draw.rectangle(rect, fill="white", outline="white")

    except Exception as e:
        print(f"Erro no mascaramento: {e}")
        # Se der erro no mascaramento, retorna a imagem original sem travar
        pass
    
    return image

# --- FUNÇÃO DE CONVERSÃO ---
def convert_pdf_to_docx(file_bytes):
    # 1. Carregar para ler estrutura (pdfplumber)
    pdf_plumb = pdfplumber.open(BytesIO(file_bytes))
    
    # 2. Carregar para converter em imagem (pdf2image)
    images = convert_from_bytes(file_bytes)
    
    doc = Document()
    
    # Margens A4
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)

    for i, img in enumerate(images):
        # --- APLICAR MÁSCARA ANTES DO RESIZE ---
        # Verifica se a página existe no pdfplumber (segurança)
        if i < len(pdf_plumb.pages):
            page_plumb = pdf_plumb.pages[i]
            # Aplica a "tinta branca" nas colunas de valor e linha total
            img = apply_masking(img, page_plumb)
        
        # --- PROCESSAMENTO PADRÃO ---
        # Redimensionamento SEI (Otimizado 552x781)
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
        with st.spinner('Analisando tabelas, aplicando máscaras e convertendo...'):
            try:
                processed_files = []
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    progress_bar.progress((index + 1) / qtd)

                st.success("✅ Conversão concluída com sucesso!")
                
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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v6.0</div>', unsafe_allow_html=True)
