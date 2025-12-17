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

# --- FUNÇÃO DE MASCARAMENTO (CALIBRADA v7) ---
def apply_masking(image, pdf_page):
    """
    Estratégia: Localizar o X da coluna alvo e mascarar toda a área à direita.
    Localizar a última linha 'Total' e mascarar toda a área inferior.
    """
    try:
        # Configurações para pegar tabelas mesmo com linhas falhas
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        
        draw = ImageDraw.Draw(image)
        im_width, im_height = image.size
        
        # Fator de escala (PDF Points -> Imagem Pixels)
        scale_x = im_width / pdf_page.width
        scale_y = im_height / pdf_page.height

        # Palavras-chave exatas do seu documento + variações
        keywords_target = ["preço unitário", "valor unitário", "vlr. unit", "unitário"]
        keyword_total_row = "total"

        for table in tables:
            # Pega o texto da primeira linha (cabeçalho)
            # table.rows[0].cells fornece as coordenadas das células do cabeçalho
            if not table.rows:
                continue

            header_cells = table.rows[0].cells
            # Extrai texto usando o pdf_page.crop para garantir leitura correta
            header_texts = []
            for cell in header_cells:
                if cell:
                    # cell é (x0, top, x1, bottom)
                    try:
                        cropped = pdf_page.crop(cell)
                        text = cropped.extract_text()
                        header_texts.append(text.lower() if text else "")
                    except:
                        header_texts.append("")
                else:
                    header_texts.append("")
            
            # 1. IDENTIFICAR ONDE COMEÇA O CORTE VERTICAL
            mask_start_x = None
            
            for i, text in enumerate(header_texts):
                # Se encontrar "preço unitário" ou similar
                if any(k in text for k in keywords_target):
                    # Pega o X0 (início) dessa célula
                    cell_coords = header_cells[i]
                    if cell_coords:
                        mask_start_x = cell_coords[0]
                        break # Achou a primeira coluna de valor, para aqui e mascara tudo à direita
            
            # Se achou a coluna, desenha a TARJA VERTICAL
            if mask_start_x is not None:
                # Coordenadas da tabela inteira
                table_top = table.bbox[1]
                table_bottom = table.bbox[3]
                table_right = table.bbox[2]
                
                # Retângulo: Do início da coluna Unitário até o fim da tabela (direita)
                # E do topo da tabela até o fim da tabela
                rect = [
                    mask_start_x * scale_x,
                    table_top * scale_y,
                    table_right * scale_x, # Vai até a borda direita da tabela
                    table_bottom * scale_y
                ]
                # Desenha branco
                draw.rectangle(rect, fill="white", outline="white")

            # 2. IDENTIFICAR LINHA DE TOTAL (TARJA HORIZONTAL)
            # Verifica a última linha
            last_row = table.rows[-1]
            last_row_text = ""
            
            # Tenta ler o texto da primeira célula da última linha
            try:
                if last_row.cells[0]:
                    cropped_last = pdf_page.crop(last_row.cells[0])
                    last_row_text = cropped_last.extract_text().lower() if cropped_last.extract_text() else ""
            except:
                pass
            
            # Se a palavra TOTAL estiver na última linha (geralmente na primeira célula ou mesclada)
            # Ou se a linha anterior for dados e essa for resumo
            if keyword_total_row in last_row_text:
                # Coordenadas da última linha inteira
                # Nota: table.rows[-1] é um objeto Row, precisamos pegar o bbox dele ou das células
                # O pdfplumber Row não tem bbox direto, pegamos min/max das células
                l_top = min(c[1] for c in last_row.cells if c)
                l_bottom = max(c[3] for c in last_row.cells if c)
                l_left = table.bbox[0]
                l_right = table.bbox[2]
                
                rect_total = [
                    l_left * scale_x,
                    l_top * scale_y,
                    l_right * scale_x,
                    l_bottom * scale_y
                ]
                draw.rectangle(rect_total, fill="white", outline="white")

    except Exception as e:
        print(f"Erro no mascaramento: {e}")
        pass
    
    return image

# --- FUNÇÃO DE CONVERSÃO ---
def convert_pdf_to_docx(file_bytes):
    # 1. Ler estrutura
    pdf_plumb = pdfplumber.open(BytesIO(file_bytes))
    
    # 2. Imagens
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
        if i < len(pdf_plumb.pages):
            page_plumb = pdf_plumb.pages[i]
            img = apply_masking(img, page_plumb)
        
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
                        label=f"📥 Salvar {name}",
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
                        label="📥 Salvar Todos (.ZIP)",
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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v7.0</div>', unsafe_allow_html=True)
