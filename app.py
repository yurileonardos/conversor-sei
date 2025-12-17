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
# Mude para False quando validar que a Pág 1 funcionou
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
    st.warning("🔴 MODO DIAGNÓSTICO ATIVADO. Máscaras em VERMELHO.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES LÓGICAS ---

def is_price_format(text):
    """Detecta números decimais (ex: 100,00)"""
    if not text: return False
    clean = text.strip().replace(" ", "")
    # Regex flexível: Aceita numeros com virgula e 2 digitos no final
    match = re.search(r'[\d\.]*,\d{2}', clean)
    if match:
        # Filtro: garante que tem pouco ruído de letras
        invalids = sum(1 for c in clean if c.lower() not in '0123456789.,r$()')
        return invalids <= 2
    return False

def find_x_by_visual_scan(pdf_page):
    """Estratégia 1: Procura pilha de números (Para Págs 2, 3, 4...)"""
    words = pdf_page.extract_words()
    price_words = [w for w in words if is_price_format(w['text'])]
    
    if not price_words: return None
        
    # Agrupa por coordenadas X (Cluster)
    x_clusters = []
    tolerance = 10 # Aumentei a tolerância para 10px
    
    for w in price_words:
        x0 = w['x0']
        found_cluster = False
        for cluster in x_clusters:
            if abs(cluster['avg_x'] - x0) < tolerance:
                cluster['points'].append(x0)
                cluster['avg_x'] = sum(cluster['points']) / len(cluster['points'])
                cluster['count'] += 1
                found_cluster = True
                break
        if not found_cluster:
            x_clusters.append({'avg_x': x0, 'points': [x0], 'count': 1})
            
    # Filtra clusters (precisa estar na direita da página)
    # Reduzi exigência de count para 1 caso seja uma linha única muito clara
    page_width = pdf_page.width
    valid_clusters = [c for c in x_clusters if c['avg_x'] > (page_width * 0.45)]
    
    if not valid_clusters: return None
        
    # Pega o cluster mais à esquerda
    best_cluster = min(valid_clusters, key=lambda c: c['avg_x'])
    return best_cluster['avg_x']

def find_x_by_header_scan(pdf_page):
    """Estratégia 2: Procura palavras de cabeçalho (Para Pág 1)"""
    words = pdf_page.extract_words()
    
    # Palavras-chave que indicam o inicio da área de preço
    target_words = ["unitário", "unitario", "estimado", "total", "(r$)", "(r$)"]
    
    found_candidates = []
    for w in words:
        txt = w['text'].lower().strip()
        if any(target in txt for target in target_words):
            # Garante que está na metade direita da página
            if w['x0'] > (pdf_page.width * 0.4):
                found_candidates.append(w['x0'])
    
    if found_candidates:
        # Retorna o X mais à esquerda encontrado (provavelmente "Unitário")
        return min(found_candidates)
    return None

def check_for_stoppers(pdf_page):
    """Verifica se há palavras de parada (Texto Jurídico)"""
    text = pdf_page.extract_text().lower()
    keys_stop = [
        "local de entrega", "prazo de entrega", "assinatura do contrato", 
        "garantia dos bens", "sanções administrativas", "obrigações da contratada", 
        "fiscalização", "gestão do contrato", "cláusula", "vigência", "dotação orçamentária"
    ]
    return any(k in text for k in keys_stop)

# --- FUNÇÃO DE MASCARAMENTO (v25.0 - HÍBRIDA) ---
def apply_masking_v25(image, pdf_page, mask_state):
    
    im_width, im_height = image.size
    
    # 1. VERIFICA STOPPER (Texto Jurídico)
    if check_for_stoppers(pdf_page):
        mask_state['active'] = False
        mask_state['cut_x_percent'] = None
    
    # 2. DETECÇÃO DE CORTE
    else:
        found_x = None
        
        # A) Tenta Scanner Visual (Prioridade: Números Reais)
        found_x = find_x_by_visual_scan(pdf_page)
        
        # B) Se falhar (Pág 1 com poucos itens), tenta Scanner de Cabeçalho
        if found_x is None:
            found_x = find_x_by_header_scan(pdf_page)
        
        # ATUALIZA ESTADO
        if found_x:
            mask_state['active'] = True
            mask_state['cut_x_percent'] = found_x / pdf_page.width
            
        # Se não achou nada nesta página, mantém o estado anterior (Herança)
        # a menos que pareça uma página vazia/texto (Stopper cuida disso)

    # 3. APLICAÇÃO VISUAL
    if mask_state['active'] and mask_state['cut_x_percent']:
        draw = ImageDraw.Draw(image, "RGBA")
        
        cut_x_pixel = mask_state['cut_x_percent'] * im_width
        
        # Recuo de segurança (-5px) para garantir que cobre o início do número/texto
        safe_cut_x = cut_x_pixel - 5 
        
        # Cores
        if DEBUG_MODE:
            fill = (255, 0, 0, 100)
            line = "red"
        else:
            fill = "white"
            line = "black"

        draw.rectangle(
            [safe_cut_x, 0, im_width, im_height],
            fill=fill, outline=None
        )
        
        draw.line([(safe_cut_x, 0), (safe_cut_x, im_height)], fill=line, width=3)

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
            img, mask_state = apply_masking_v25(img, pdf_plumb.pages[i], mask_state)
        
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
    btn_label = "🚀 Processar (Diagnóstico Vermelho)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Aplicando estratégia híbrida (Visual + Cabeçalho)...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v25.0 (Hybrid Fix)</div>', unsafe_allow_html=True)
