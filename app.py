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
# Mantive True (Vermelho) para você validar que a Qtde foi salva.
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
    st.warning("🔴 MODO DIAGNÓSTICO: Máscaras em VERMELHO.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES LÓGICAS ---

def clean_text(text):
    if not text: return ""
    return str(text).strip().lower()

def is_table_blocked(table, pdf_page):
    """
    PASSO 1: O FILTRO DE BLOQUEIO.
    Verifica se a tabela contém texto jurídico ou assinaturas.
    Retorna True se deve ser ignorada.
    """
    # 1. Verifica número de colunas (Tabelas de itens têm muitas, texto tem 1 ou 2)
    max_cols = 0
    if table.rows:
        max_cols = max([len(r.cells) for r in table.rows])
    
    if max_cols < 3:
        return True # Bloqueia tabelas de layout/texto

    # 2. Verifica palavras proibidas (Stoppers)
    stop_words = [
        "local", "entrega", "prazo", "assinatura", "garantia", "sancoes", "sanções", 
        "obrigacoes", "obrigações", "fiscalizacao", "fiscalização", "gestao", "clausula", 
        "cláusula", "vigencia", "vigência", "dotacao", "objeto", "condicoes", "foro",
        "eletronicamente", "autenticidade", "código verificador", "brasília"
    ]
    
    # Amostra de texto (Topo e Fundo da tabela)
    sample_txt = ""
    rows_to_check = table.rows[:3] + table.rows[-2:] # 3 primeiras e 2 ultimas
    for r in rows_to_check:
        for c in r.cells:
            if c:
                try:
                    crop = pdf_page.crop(c)
                    sample_txt += clean_text(crop.extract_text()) + " "
                except: pass
    
    if any(sw in sample_txt for sw in stop_words):
        return True
        
    return False

def determine_cut_x(table, pdf_page):
    """
    Define a coordenada X do corte baseado em prioridades sequenciais.
    Retorna: (cut_x, 'left' ou 'right')
    """
    # PRIORIDADE 1: Cabeçalhos de Preço (Corta à ESQUERDA da coluna)
    # Procuramos explicitamente onde o preço começa.
    price_headers = ["preço unit", "preco unit", "valor unit", "vlr. unit", "estimado (r$)", "total (r$)"]
    
    for r in table.rows[:3]: # Apenas cabeçalho
        for cell in r.cells:
            if cell:
                try:
                    crop = pdf_page.crop(cell)
                    txt = clean_text(crop.extract_text())
                    if any(h in txt for h in price_headers):
                        return cell[0], 'left' # Retorna a borda ESQUERDA
                except: pass

    # PRIORIDADE 2: Cabeçalhos de Âncora (Corta à DIREITA da coluna)
    # Se não achou preço, procura Qtde/Unid e corta logo depois.
    anchor_headers = ["qtde", "qtd", "quantidade", "quant", "unid", "unidade", "catmat", "uf"]
    
    for r in table.rows[:3]:
        for cell in r.cells:
            if cell:
                try:
                    crop = pdf_page.crop(cell)
                    txt = clean_text(crop.extract_text())
                    # Match exato ou inicio de palavra
                    if txt in anchor_headers or any(txt.startswith(a) for a in anchor_headers):
                        return cell[2], 'right' # Retorna a borda DIREITA
                except: pass
                
    return None, None

# --- FUNÇÃO DE MASCARAMENTO (v30.0 - SEQUENTIAL PIPELINE) ---
def apply_masking_v30(image, pdf_page, global_cut_percent):
    
    # Extração de tabelas
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    # Prioridade para LINHAS, usa TEXTO só se linhas falhar
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    # Se não achou tabelas, mantemos o estado global (pode ser página de continuação sem linhas claras)
    # Mas se tiver texto de parada na página, resetamos.
    page_text = clean_text(pdf_page.extract_text())
    if "assinado eletronicamente" in page_text or "cláusula" in page_text:
        global_cut_percent = None # Reseta memória
        return image.convert("RGB"), None

    for table in all_tables:
        if not table.rows: continue
        
        # 1. FILTRO DE BLOQUEIO (Resolve Páginas 7-11 e 5-texto)
        if is_table_blocked(table, pdf_page):
            # Se encontrou tabela bloqueada, considera que o contexto mudou. Reseta global.
            global_cut_percent = None
            continue 

        # 2. DECISÃO DE CORTE (Resolve Página 1 e Atualizações)
        cut_x, mode = determine_cut_x(table, pdf_page)
        
        current_cut_percent = None
        
        if cut_x:
            # Encontrou novo cabeçalho! Atualiza global.
            current_cut_percent = cut_x / pdf_page.width
            
            # Ajuste Fino: Se o modo for 'right' (Qtde), adiciona margem segura
            # Se for 'left' (Preço), não precisa margem (ou pequena negativa)
            if mode == 'right':
                # Adiciona 0.5% da largura da página como margem para não colar na letra
                current_cut_percent += 0.005 
            
            global_cut_percent = current_cut_percent
            
        elif global_cut_percent:
            # Não tem cabeçalho, mas tem memória (Resolve Páginas 2, 3, 4)
            current_cut_percent = global_cut_percent
            
        # 3. APLICAÇÃO VISUAL (Resolve "Iluminar a tabela apenas")
        if current_cut_percent:
            
            x_pixel = current_cut_percent * im_width
            
            # Limites Verticais ESTRITOS da tabela
            # bbox = (x0, top, x1, bottom)
            t_bbox = table.bbox
            scale_y = im_height / pdf_page.height
            
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação: Só desenha se x estiver dentro da imagem
            if x_pixel < im_width:
                
                # Cores
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100)
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # Desenha o retângulo APENAS dentro da altura da tabela
                draw.rectangle(
                    [x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )
                
                draw.line([(x_pixel, top_pixel), (x_pixel, bottom_pixel)], fill=line, width=3)
                
                if not DEBUG_MODE:
                    draw.line([(x_pixel, top_pixel), (x_pixel - 5, top_pixel)], fill="black", width=2)
                    draw.line([(x_pixel, bottom_pixel), (x_pixel - 5, bottom_pixel)], fill="black", width=2)

    return image.convert("RGB"), global_cut_percent

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

    global_cut_percent = None # Memória persistente

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, global_cut_percent = apply_masking_v30(img, pdf_plumb.pages[i], global_cut_percent)
        
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
        with st.spinner('Processando via Pipeline Sequencial...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v30.0 (Sequential Pipeline)</div>', unsafe_allow_html=True)
