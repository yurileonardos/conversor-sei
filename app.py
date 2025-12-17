import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile
import pdfplumber
from PIL import Image, ImageDraw
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="SEI Converter ATA - SGB",
    page_icon="📑",
    layout="centered"
)

# --- CONFIGURAÇÃO DE DIAGNÓSTICO ---
# True = Máscara Vermelha (Visualizar o que foi detectado)
# False = Máscara Branca (Versão Final)
DEBUG_MODE = True 

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

# --- TÍTULO ---
st.title("📑 SEI Converter ATA - SGB")

if DEBUG_MODE:
    st.warning("🔴 MODO DIAGNÓSTICO: As máscaras aparecerão em VERMELHO. Se o Grupo 1 for coberto, mudaremos para Branco na próxima versão.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES AUXILIARES ---

def clean_text(text):
    if not text: return ""
    # Remove espaços extras e quebras
    return str(text).strip()

def is_numeric_decimal(text):
    """
    Verifica se o texto parece um número decimal ou monetário.
    Ex aceitos: 100,00 | 1.000,50 | R$ 50,00 | 50,23
    """
    if not text: return False
    # Regex: Procura numeros que terminam com virgula e 2 digitos (padrão moeda BR)
    # Aceita R$ opcional no inicio
    pattern = r'(?:R\$\s*)?[\d\.]*,\d{2}\b' 
    match = re.search(pattern, text)
    return bool(match)

def scan_table_columns_for_prices(table, pdf_page):
    """
    Analisa coluna por coluna. Se achar uma coluna onde a maioria das células
    são números decimais, retorna a coordenada X (esquerda) dessa coluna.
    """
    if not table.rows: return None
    
    # Transpõe a leitura: vamos olhar coluna por coluna, não linha por linha
    # Assume que a tabela é uniforme. Pega o número de células da primeira linha válida.
    num_cols = 0
    for r in table.rows:
        if r.cells:
            num_cols = len(r.cells)
            break
            
    if num_cols < 3: return None # Ignora tabelas muito pequenas (texto)

    # Varre coluna por coluna
    for col_idx in range(num_cols):
        numeric_hits = 0
        total_valid_cells = 0
        first_valid_cell_x = None
        
        # Analisa até 15 linhas da coluna para ter amostragem
        sample_limit = min(15, len(table.rows))
        
        for row_idx in range(sample_limit):
            try:
                # Proteção de índice
                if col_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[col_idx]
                    if cell:
                        # Extrai texto
                        if isinstance(cell, (list, tuple)) and len(cell) == 4:
                            # Se não temos a coordenada X da primeira célula, pegamos agora
                            if first_valid_cell_x is None:
                                first_valid_cell_x = cell[0] # Borda Esquerda
                                
                            crop = pdf_page.crop(cell)
                            txt = clean_text(crop.extract_text())
                            
                            if txt:
                                total_valid_cells += 1
                                if is_numeric_decimal(txt):
                                    numeric_hits += 1
            except:
                pass
        
        # CRITÉRIO DE DECISÃO:
        # Se mais de 50% das células preenchidas nessa coluna forem números decimais
        # E tivermos pelo menos 2 números identificados (para evitar falso positivo em 1 célula)
        if total_valid_cells > 0 and numeric_hits >= 2:
            ratio = numeric_hits / total_valid_cells
            if ratio > 0.5:
                # Achamos uma coluna de preço!
                # Como queremos mascarar DAQUI para a direita, retornamos o X desta coluna
                
                # Validação Extra: Preço raramente é a primeira coluna (Item)
                # Só retorna se não for a col 0, a menos que o PDF esteja muito quebrado
                if col_idx > 0: 
                    return first_valid_cell_x
                
    return None

# --- FUNÇÃO DE MASCARAMENTO (v22.0 - CONTENT SCANNER) ---
def apply_masking_v22(image, pdf_page, mask_state):
    
    # Busca tabelas (Linhas e Texto)
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    # STOPPERS (Texto Jurídico - Proteção Páginas 7/8)
    keys_stop = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sancoes", "sançoes", "obrigacoes", "fiscalizacao", 
        "gestao", "clausula", "vigencia", "recursos", "dotacao", "objeto", "condicoes",
        "multas", "infracoes", "penalidades", "rescisao", "foro"
    ]
    
    # START MANUAL (Caso a varredura numérica falhe, usamos o cabeçalho Qtde como backup)
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "unid", "consumo", "catmat"]

    for table in all_tables:
        if not table.rows: continue
        
        t_bbox = table.bbox
        
        # --- 1. VERIFICAÇÃO DE STOPPER (Texto) ---
        found_stopper = False
        limit_rows = min(5, len(table.rows))
        for r_idx in range(limit_rows):
            for cell in table.rows[r_idx].cells:
                if cell and isinstance(cell, (list, tuple)):
                    try:
                        crop = pdf_page.crop(cell)
                        txt = str(crop.extract_text()).lower()
                        if any(k in txt for k in keys_stop):
                            found_stopper = True
                            break
                    except: pass
            if found_stopper: break
            
        if found_stopper:
            mask_state['active'] = False
            mask_state['cut_x_percent'] = None
            continue

        # --- 2. VERIFICAÇÃO DE START (Scanner de Conteúdo) ---
        found_start_x = None
        
        # A) Tenta detectar colunas numéricas (Preço)
        price_col_x = scan_table_columns_for_prices(table, pdf_page)
        
        if price_col_x:
            # Achou coluna de preço pelo conteúdo!
            found_start_x = price_col_x
            
        # B) Backup: Tenta detectar coluna Qtde pelo cabeçalho (se a numérica falhar)
        elif not mask_state['active']:
             for r_idx in range(limit_rows):
                for cell in table.rows[r_idx].cells:
                    try:
                        if cell and isinstance(cell, (list, tuple)):
                            crop = pdf_page.crop(cell)
                            txt = clean_text(crop.extract_text()).lower()
                            if any(k == txt or k in txt.split() for k in keys_qty):
                                found_start_x = cell[2] # Borda Direita da Qtde
                                break
                    except: pass
                if found_start_x: break

        # --- 3. GESTÃO DE ESTADO ---
        
        if found_start_x:
            mask_state['active'] = True
            mask_state['cut_x_percent'] = found_start_x / pdf_page.width
            
        # Proteção Estrutural (Páginas 7/8): Se virar texto corrido (1 coluna), desliga
        elif mask_state['active']:
             cols_count = max([len(r.cells) for r in table.rows])
             # Se tiver menos de 3 colunas, assume que não é mais tabela de itens
             if cols_count < 3:
                 mask_state['active'] = False
                 mask_state['cut_x_percent'] = None

        # --- 4. APLICAÇÃO VISUAL ---
        if mask_state['active'] and mask_state['cut_x_percent']:
            
            cut_x_pixel = mask_state['cut_x_percent'] * im_width
            scale_y = im_height / pdf_page.height
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação de posição (O corte deve ser lógico)
            t_x0_pixel = t_bbox[0] * (im_width / pdf_page.width)
            
            if cut_x_pixel > t_x0_pixel:
                
                # CORES
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100)
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # Máscara
                draw.rectangle(
                    [cut_x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )

                # Linha
                draw.line([(cut_x_pixel, top_pixel), (cut_x_pixel, bottom_pixel)], fill=line, width=3)
                
                if not DEBUG_MODE:
                    draw.line([(cut_x_pixel, top_pixel), (cut_x_pixel - 5, top_pixel)], fill="black", width=2)
                    draw.line([(cut_x_pixel, bottom_pixel), (cut_x_pixel - 5, bottom_pixel)], fill="black", width=2)

    return image.convert("RGB"), mask_state

# --- FUNÇÃO DE CONVERSÃO ---
def convert_pdf_to_docx(file_bytes):
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
    section.bottom_margin = Cm(0.5)

    mask_state = {'active': False, 'cut_x_percent': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v22(img, pdf_plumb.pages[i], mask_state)
        
        img = img.resize((595, 842)) 
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
        img_byte_arr.seek(0)

        doc.add_picture(img_byte_arr, width=Cm(18.0))
        
        par = doc.paragraphs[-1]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        
        if i < len(images) - 1:
            doc.add_page_break()
    
    docx_io = BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io

# --- INTERFACE ---
uploaded_files = st.file_uploader("Arraste e solte seus arquivos PDF aqui:", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.write("---")
    btn_label = "🚀 Processar (Diagnóstico - VERMELHO)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Escaneando colunas numéricas...'):
            try:
                processed_files = []
                for uploaded_file in uploaded_files:
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))

                st.success("✅ Concluído!")
                
                if len(processed_files) == 1:
                    name, data = processed_files[0]
                    st.download_button("📥 Baixar Arquivo DOCX", data, file_name=name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, data in processed_files:
                            zf.writestr(name, data.getvalue())
                    zip_buffer.seek(0)
                    st.download_button("📥 Baixar Todos (.ZIP)", zip_buffer, "Arquivos_SEI.zip", mime="application/zip")

            except Exception as e:
                st.error(f"Erro: {e}")

# --- RODAPÉ ---
st.write("---")
st.subheader("📚 Guia Rápido: Como inserir no SEI")
col1, col2 = st.columns([0.15, 0.85])
with col1:
    try:
        st.image("icone_sei.png", width=50) 
    except:
        st.write("🧩")
with col2:
    st.markdown("*1º Localize o ícone:* No editor do SEI, clique no botão da função *INSERIR CONTEÚDO EXTERNO*.")
st.write("")
st.markdown("*2º Configure a inserção:* Faça o upload do arquivo Word gerado aqui.")
st.warning("⚠️ *IMPORTANTE:* Certifique-se de deixar todas as caixas de seleção *DESMARCADAS*.")
try:
    st.image("print_sei.png", caption="Exemplo: Deixe as opções desmarcadas.", use_container_width=True)
except:
    pass

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v22.0 (Content Scanner)</div>', unsafe_allow_html=True)
