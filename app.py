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

# --- MODO DIAGNÓSTICO ---
# Mude para False para a versão final (Branca)
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
    st.warning("🔴 MODO DIAGNÓSTICO: Máscaras em VERMELHO (Respeitando altura da tabela).")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES LÓGICAS ---

def clean_text(text):
    if not text: return ""
    return str(text).strip()

def is_numeric_decimal(text):
    """
    Identifica células que são CLARAMENTE valores monetários.
    Ex: 100,00 | 1.520,50 | R$ 50,00
    Rejeita: Datas, Leis (14.133), Inteiros (50)
    """
    if not text: return False
    clean = text.replace(" ", "")
    # Regex: Numeros (com pontos opcionais) + Virgula + 2 Digitos no final
    match = re.search(r'[\d\.]*,\d{2}$', clean)
    if match:
        # Filtro anti-falso positivo (ex: Lei 8.666/93 não passa)
        # Se tiver caracteres que não sejam numeros, pontos, virgulas ou R$, rejeita.
        if any(c for c in clean if c.lower() not in '0123456789.,r$'):
            return False
        return True
    return False

def check_structure_and_stop(table, pdf_page):
    """
    Verifica se a tabela é válida para mascaramento.
    Retorna False se for Assinatura, Texto Jurídico ou tiver poucas colunas.
    """
    # 1. Checagem de Colunas (Elimina Pág 9)
    # Tabelas de itens têm Item, Descrição, Unid, Qtd, Valor... (Minimo 3 colunas visualmente)
    max_cols = 0
    if table.rows:
        max_cols = max([len(r.cells) for r in table.rows])
    
    if max_cols < 3:
        return False # Ignora tabelas de assinatura ou layout simples

    # 2. Checagem de Texto (Stopwords)
    keys_stop = [
        "local de entrega", "prazo", "assinatura", "garantia", "sanções", 
        "obrigações", "fiscalização", "gestão", "cláusula", "vigência", 
        "dotação", "assinado", "eletronicamente", "testemunhas", "foro"
    ]
    
    # Amostra de texto da tabela
    sample_txt = ""
    for r in table.rows[:5]: # Olha as primeiras 5 linhas
        for c in r.cells:
            if c:
                try:
                    crop = pdf_page.crop(c)
                    sample_txt += crop.extract_text().lower() + " "
                except: pass
    
    if any(k in sample_txt for k in keys_stop):
        return False
        
    return True

def find_cut_x_in_table(table, pdf_page):
    """
    Encontra a coordenada X onde começa a área de preço dentro de uma tabela específica.
    Usa abordagem híbrida: Conteúdo Numérico OU Cabeçalho.
    """
    found_x = None
    
    # A) ESTRATÉGIA DE CONTEÚDO (Varre colunas procurando números decimais)
    # Analisa coluna por coluna (transversal)
    max_cols = max([len(r.cells) for r in table.rows])
    
    # Itera sobre índices de coluna (0, 1, 2...)
    for col_idx in range(max_cols):
        decimal_hits = 0
        valid_cells = 0
        col_x = None
        
        # Olha as primeiras 10 linhas
        for r_idx in range(min(10, len(table.rows))):
            try:
                row_cells = table.rows[r_idx].cells
                if col_idx < len(row_cells):
                    cell = row_cells[col_idx]
                    if cell and isinstance(cell, (list, tuple)):
                        if col_x is None: col_x = cell[0] # Pega X da borda esquerda
                        
                        crop = pdf_page.crop(cell)
                        txt = clean_text(crop.extract_text())
                        if txt:
                            valid_cells += 1
                            if is_numeric_decimal(txt):
                                decimal_hits += 1
            except: pass
        
        # Se a coluna tem >50% de numeros decimais, é Preço
        if valid_cells > 0 and decimal_hits >= 1: # Flexibilizei para 1 hit se for claro
             ratio = decimal_hits / valid_cells
             if ratio >= 0.5:
                 # Validação: Deve estar na direita (>40% da página)
                 if col_x and col_x > (pdf_page.width * 0.4):
                     return col_x

    # B) ESTRATÉGIA DE CABEÇALHO (Backup para Pág 1 se não tiver dados suficientes)
    keys_header = ["unitário", "unitario", "estimado", "total", "(r$)", "valor"]
    
    for r in table.rows[:3]: # Primeiras 3 linhas
        for cell in r.cells:
            try:
                if cell and isinstance(cell, (list, tuple)):
                    crop = pdf_page.crop(cell)
                    txt = str(crop.extract_text()).lower()
                    if any(k in txt for k in keys_header):
                         # Verifica posição
                         if cell[0] > (pdf_page.width * 0.4):
                             return cell[0] # Borda Esquerda
            except: pass
            
    return None

# --- FUNÇÃO DE MASCARAMENTO (v26.0 - BOUNDED MASK) ---
def apply_masking_v26(image, pdf_page, mask_state):
    
    # Busca todas as tabelas (Linhas e Texto)
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    for table in all_tables:
        if not table.rows: continue
        
        # 1. VALIDAÇÃO DE ESTRUTURA (Resolve Pág 9 e Textos)
        is_valid = check_structure_and_stop(table, pdf_page)
        
        if not is_valid:
            # Se a tabela é inválida (assinatura, texto), e a máscara estava ativa,
            # verificamos se devemos desligar.
            # Se for uma mudança brusca de estrutura (ex: 5 cols -> 1 col), desliga.
            cols = max([len(r.cells) for r in table.rows])
            if cols < 3:
                mask_state['active'] = False
                mask_state['cut_x_percent'] = None
            continue # Pula para a próxima tabela sem desenhar nada nesta
            
        # 2. LOCALIZAÇÃO DO CORTE (Resolve Pág 1 e Continuações)
        cut_x = find_cut_x_in_table(table, pdf_page)
        
        if cut_x:
            mask_state['active'] = True
            mask_state['cut_x_percent'] = cut_x / pdf_page.width
        
        # 3. APLICAÇÃO VISUAL (Resolve Pág 1 - Limites Verticais)
        if mask_state['active'] and mask_state['cut_x_percent']:
            
            # Coordenadas Horizontais
            cut_x_pixel = mask_state['cut_x_percent'] * im_width
            safe_cut_x = cut_x_pixel - 5 
            
            # Coordenadas Verticais (LIMITADAS À TABELA)
            # Usa o bbox da tabela para definir onde começa e termina o vermelho
            t_bbox = table.bbox # (x0, top, x1, bottom)
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação geométrica: O corte deve estar dentro da tabela
            t_x0_pixel = t_bbox[0] * scale_x
            if cut_x_pixel > t_x0_pixel:
                
                # Cores
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100) # Vermelho
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # Desenha o retângulo APENAS dentro dos limites da tabela
                draw.rectangle(
                    [safe_cut_x, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )
                
                # Linha Vertical
                draw.line([(safe_cut_x, top_pixel), (safe_cut_x, bottom_pixel)], fill=line, width=3)
                
                # Acabamento (linhas horizontais no topo e base da máscara)
                if not DEBUG_MODE:
                    draw.line([(safe_cut_x, top_pixel), (safe_cut_x - 5, top_pixel)], fill="black", width=2)
                    draw.line([(safe_cut_x, bottom_pixel), (safe_cut_x - 5, bottom_pixel)], fill="black", width=2)

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
    
    # Configuração A4
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
            img, mask_state = apply_masking_v26(img, pdf_plumb.pages[i], mask_state)
        
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
    btn_label = "🚀 Processar (Diagnóstico Final)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Processando...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v26.0 (Bounded & Specific)</div>', unsafe_allow_html=True)
