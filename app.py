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
# Se True, desenha em VERMELHO. Se False, desenha em BRANCO (Final).
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
    st.warning("🔴 MODO DIAGNÓSTICO: As máscaras aparecerão em VERMELHO para facilitar a conferência.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÃO DE LIMPEZA ---
def clean_text(text):
    if not text: return ""
    text = str(text).lower().strip()
    # Remove pontuação básica para palavras-chave, mas mantém para números depois
    clean = text
    for ch in ['.', ':', '-', '/']:
        clean = clean.replace(ch, '')
    replacements = {
        'ç': 'c', 'ã': 'a', 'á': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
        'í': 'i', 'ó': 'o', 'õ': 'o', 'ú': 'u'
    }
    for k, v in replacements.items():
        clean = clean.replace(k, v)
    return clean

# --- DETECTOR DE VALORES MONETÁRIOS (REGEX) ---
def is_money_value(text):
    """Retorna True se o texto parece um valor monetário (R$ XX,XX ou XX.XXX,XX)"""
    if not text: return False
    # Padrões:
    # 1. R$ 100,00 ou R$100,00
    # 2. 1.000,00 (ponto milhar, virgula decimal)
    # 3. 100,00 (apenas virgula decimal)
    # Ignora números simples como "100" ou datas "2024"
    
    # Limpa espaços extras
    t = text.strip()
    
    # Regex para formato brasileiro de moeda
    # Procura por R$ opcional + numeros com ponto opcional + virgula obrigatória + 2 digitos
    pattern = r'(?:r\$\s*)?[\d\.]+\,\d{2}'
    
    match = re.search(pattern, t)
    return bool(match)

# --- FUNÇÃO DE MASCARAMENTO (v21.0 - MONEY PATTERN DETECTOR) ---
def apply_masking_v21(image, pdf_page, mask_state):
    
    # Combina estratégias de tabela (Linhas + Texto)
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    # PALAVRAS-CHAVE
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "unid", "consumo", "catmat"]
    
    # STOPPERS (Texto Jurídico)
    keys_stop = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sancoes", "sançoes", "obrigacoes", "fiscalizacao", 
        "gestao", "clausula", "vigencia", "recursos", "dotacao", "objeto", "condicoes",
        "multas", "infracoes", "penalidades", "rescisao", "foro"
    ]

    for table in all_tables:
        if not table.rows: continue
        
        t_bbox = table.bbox # (x0, top, x1, bottom)
        
        found_cut_x = None
        found_stopper = False
        text_content_sample = ""
        
        # --- DEEP SCAN (Varredura Profunda: até 10 linhas) ---
        # Aumentamos para 10 para pegar valores monetários dentro da tabela
        limit_rows = min(10, len(table.rows))
        
        for r_idx in range(limit_rows):
            row_cells = table.rows[r_idx].cells
            for cell in row_cells:
                if not cell: continue
                try:
                    if isinstance(cell, (list, tuple)) and len(cell) == 4:
                        crop = pdf_page.crop(cell)
                        raw_text = crop.extract_text()
                        cleaned = clean_text(raw_text)
                        text_content_sample += cleaned + " "
                        
                        # 1. VERIFICA STOPPER (Texto Jurídico)
                        if any(k in cleaned for k in keys_stop):
                            found_stopper = True
                        
                        # 2. VERIFICA START: CABEÇALHO 'QTDE' (Âncora Padrão)
                        if any(k == cleaned or k in cleaned.split() for k in keys_qty):
                            found_cut_x = cell[2] # Borda Direita
                        
                        # 3. VERIFICA START: PADRÃO MONETÁRIO (R$ ou XX,XX)
                        # Se acharmos dinheiro, cortamos à ESQUERDA dessa célula
                        elif found_cut_x is None and is_money_value(raw_text):
                            # Filtro de sanidade: Preço geralmente está na direita (>40% da página)
                            if cell[0] > (pdf_page.width * 0.4):
                                found_cut_x = cell[0] # Borda ESQUERDA da célula de dinheiro

                except:
                    pass
            if found_cut_x or found_stopper: break

        # --- GESTÃO DE ESTADO (PERSISTÊNCIA) ---
        
        # 1. STOPPER DETECTADO -> Desliga
        if found_stopper:
            mask_state['active'] = False
            mask_state['cut_x_percent'] = None
        
        # 2. START DETECTADO (Por Qtde ou por Dinheiro) -> Liga
        elif found_cut_x is not None:
            mask_state['active'] = True
            mask_state['cut_x_percent'] = found_cut_x / pdf_page.width
            
        # 3. COLAPSO ESTRUTURAL (Proteção contra texto corrido)
        elif mask_state['active']:
            cols_count = max([len(r.cells) for r in table.rows])
            # Se virou 1 ou 2 colunas e tem muito texto, é parágrafo
            if cols_count < 3 and len(text_content_sample) > 30:
                mask_state['active'] = False
                mask_state['cut_x_percent'] = None

        # --- APLICAÇÃO DA MÁSCARA ---
        if mask_state['active'] and mask_state['cut_x_percent']:
            
            # Converte % para pixels reais
            cut_x_pixel = mask_state['cut_x_percent'] * im_width
            scale_y = im_height / pdf_page.height
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação: O corte deve estar geometricamente após o início da tabela
            t_x0_pixel = t_bbox[0] * (im_width / pdf_page.width)
            
            if cut_x_pixel > t_x0_pixel:
                
                # Definição de Cores
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100) # Vermelho Transparente
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # 1. Retângulo de Ocultação
                draw.rectangle(
                    [cut_x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )

                # 2. Linha de Fechamento
                draw.line([(cut_x_pixel, top_pixel), (cut_x_pixel, bottom_pixel)], fill=line, width=3)
                
                # 3. Acabamento
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
    
    # Configuração A4
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(0.5)

    # ESTADO INICIAL GLOBAL
    mask_state = {'active': False, 'cut_x_percent': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v21(img, pdf_plumb.pages[i], mask_state)
        
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
    btn_label = "🚀 Processar (Modo Diagnóstico - Vermelho)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Processando com detector monetário...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v21.0 (Money Detector)</div>', unsafe_allow_html=True)
