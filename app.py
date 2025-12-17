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
# True = Vermelho | False = Branco
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
    st.warning("🔴 MODO DIAGNÓSTICO: Máscaras em VERMELHO. Verifique se UF e QTD estão visíveis.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES LÓGICAS ---

def clean_text(text):
    if not text: return ""
    return str(text).strip().lower()

def is_price_column(text):
    """Verifica se o conteúdo parece preço (R$ ou 0,00)"""
    if not text: return False
    # Regex para XX,XX
    return bool(re.search(r'[\d\.]*,\d{2}', text))

def get_table_mask_x(table, pdf_page, current_global_x):
    """
    Analisa UMA tabela específica e define onde começar o corte (Eixo X).
    Retorna:
    - cut_x (float): Posição do corte
    - update_global (bool): Se deve atualizar a referência global
    """
    
    # 1. VERIFICA STOPWORDS (Se for tabela de texto, ignora)
    stop_words = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sanções", "sancoes", "obrigações", "fiscalização", 
        "gestão", "cláusula", "vigência", "dotação", "objeto", "condições", "multas", 
        "infrações", "penalidades", "rescisão", "foro", "assinado", "eletronicamente"
    ]
    
    # Extrai texto das primeiras linhas para ver se é texto jurídico
    header_text = ""
    for r in table.rows[:4]:
        for c in r.cells:
            if c:
                try:
                    crop = pdf_page.crop(c)
                    header_text += clean_text(crop.extract_text()) + " "
                except: pass
    
    if any(sw in header_text for sw in stop_words):
        return None, False # Tabela proibida (Texto)

    # 2. PROCURA CABEÇALHO "UF" ou "QTDE" (Âncora à Direita)
    # Queremos cortar DEPOIS dessas colunas.
    target_cols = ["uf", "unid", "unidade", "qtde", "qtd", "quantidade", "quant"]
    
    found_anchor_right_x = None
    
    # Varre células procurando os cabeçalhos
    for r in table.rows[:3]: # Primeiras 3 linhas
        for cell in r.cells:
            if cell and isinstance(cell, (list, tuple)):
                try:
                    crop = pdf_page.crop(cell)
                    txt = clean_text(crop.extract_text())
                    # Verifica match exato ou parcial seguro
                    if txt in target_cols or any(t == txt for t in target_cols):
                        # Achamos! O corte deve ser na borda DIREITA (x1)
                        if found_anchor_right_x is None or cell[2] > found_anchor_right_x:
                            found_anchor_right_x = cell[2]
                except: pass
    
    if found_anchor_right_x:
        # Se achou UF/QTD, essa é a nova referência mestre
        return found_anchor_right_x, True

    # 3. SE NÃO TEM CABEÇALHO (Continuação)
    # Verifica se tem números de preço. Se tiver, usa a referência global anterior.
    # Se não tiver números e não tiver cabeçalho, provavelmente é texto.
    has_prices = False
    for r in table.rows[:5]:
        for cell in r.cells:
            if cell:
                try:
                    crop = pdf_page.crop(cell)
                    if is_price_column(crop.extract_text()):
                        has_prices = True
                        break
                except: pass
        if has_prices: break
        
    if has_prices and current_global_x:
        return current_global_x, False # Mantém o corte anterior
        
    return None, False

# --- FUNÇÃO DE MASCARAMENTO (v29.0 - BOUNDING BOX & RIGHT ANCHOR) ---
def apply_masking_v29(image, pdf_page, global_mask_percent):
    
    # Extrai tabelas (Linhas e Texto)
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    # Prioriza tabelas com linhas, usa texto como fallback
    all_tables = tables_lines + tables_text if tables_lines else tables_text
    
    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    # Filtra duplicatas de tabelas (bbox muito próximos)
    unique_tables = []
    for t in all_tables:
        is_dup = False
        for ut in unique_tables:
            # Se bbox for muito parecido
            if abs(t.bbox[0] - ut.bbox[0]) < 10 and abs(t.bbox[1] - ut.bbox[1]) < 10:
                is_dup = True
                break
        if not is_dup:
            unique_tables.append(t)

    for table in unique_tables:
        if not table.rows: continue
        
        # Ignora tabelas com menos de 3 colunas (quase sempre texto)
        # EXCETO se já temos um global_mask_percent ativo (pode ser continuação quebrada)
        max_cols = max([len(r.cells) for r in table.rows])
        if max_cols < 3 and global_mask_percent is None:
            continue

        # Calcula onde cortar nesta tabela
        cut_x, should_update = get_table_mask_x(table, pdf_page, global_mask_percent)
        
        if should_update and cut_x:
            # Converte para % para persistir entre páginas
            global_mask_percent = cut_x / pdf_page.width
        
        # Se temos um ponto de corte válido (local ou global)
        final_cut_x_percent = (cut_x / pdf_page.width) if cut_x else global_mask_percent
        
        if final_cut_x_percent:
            
            # --- DESENHO RESTRITO AO BBOX DA TABELA ---
            
            # X: Começa na referência (UF/QTD) + Margem
            x_pixel = (final_cut_x_percent * im_width) + 10 # +10px de folga para não cortar a letra da Qtd
            
            # Y: Usa EXATAMENTE o topo e fundo da tabela detectada
            t_bbox = table.bbox
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação: Só desenha se o corte estiver "dentro" da largura da página
            if x_pixel < im_width:
                
                # Cores
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100)
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # Retângulo (Da âncora até a borda direita da tabela ou da página)
                # Usamos im_width na direita para garantir que cubra até o fim, 
                # mas limitado verticalmente pela tabela.
                draw.rectangle(
                    [x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )
                
                draw.line([(x_pixel, top_pixel), (x_pixel, bottom_pixel)], fill=line, width=3)
                
                if not DEBUG_MODE:
                    # Acabamento
                    draw.line([(x_pixel, top_pixel), (x_pixel - 5, top_pixel)], fill="black", width=2)
                    draw.line([(x_pixel, bottom_pixel), (x_pixel - 5, bottom_pixel)], fill="black", width=2)

    return image.convert("RGB"), global_mask_percent

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

    global_mask_percent = None # Estado persistente do corte X

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, global_mask_percent = apply_masking_v29(img, pdf_plumb.pages[i], global_mask_percent)
        
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
    btn_label = "🚀 Processar (Vermelho)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Processando com proteção de UF/QTD...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v29.0 (Box & Right Anchor)</div>', unsafe_allow_html=True)
