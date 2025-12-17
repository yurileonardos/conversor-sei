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
DEBUG_MODE = True  # True = Vermelho | False = Branco

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

def find_anchor_geometry(pdf_page):
    """
    Procura a palavra-chave 'QTDE' ou 'QUANTIDADE' ou 'UNID' para servir de âncora.
    Retorna: (x_right, y_top) -> Borda direita da palavra e Borda superior.
    """
    words = pdf_page.extract_words()
    targets = ["qtde", "qtde.", "quantidade", "quant.", "unid", "unid.", "catmat"]
    
    # Filtra candidatos na metade direita da página (para evitar falsos positivos no texto)
    candidates = []
    for w in words:
        txt = w['text'].lower().strip()
        if any(t == txt or t in txt for t in targets):
            if w['x0'] > (pdf_page.width * 0.3): # Deve estar mais para o meio/direita
                candidates.append(w)
    
    if candidates:
        # Pega o que estiver mais acima (primeiro cabeçalho que aparecer)
        best = min(candidates, key=lambda w: w['top'])
        return best['x1'], best['top'] # x1 é a borda direita
        
    return None, None

def check_page_signature(pdf_page):
    """Detecta se é página de assinatura (Pág 9)"""
    text = pdf_page.extract_text().lower()
    # Termos específicos de assinatura do SEI
    signatures = [
        "documento assinado eletronicamente", 
        "assinado eletronicamente", 
        "autenticidade deste documento", 
        "código verificador", 
        "oficial de brasília",
        "chave de acesso"
    ]
    count = sum(1 for s in signatures if s in text)
    return count >= 1

# --- FUNÇÃO DE MASCARAMENTO (v28.0 - ANCHOR & SAFETY) ---
def apply_masking_v28(image, pdf_page, mask_state):
    
    im_width, im_height = image.size
    
    # 1. SEGURANÇA: É PÁGINA DE ASSINATURA?
    if check_page_signature(pdf_page):
        mask_state['active'] = False
        mask_state['cut_x_percent'] = None
        return image.convert("RGB"), mask_state

    # 2. BUSCA ÂNCORA (QTDE)
    anchor_x, anchor_y = find_anchor_geometry(pdf_page)
    
    # Variáveis de Desenho
    draw_start_y = 0 # Padrão: Topo da página (para continuações)
    
    if anchor_x:
        # ACHOU CABEÇALHO NOVO!
        mask_state['active'] = True
        mask_state['cut_x_percent'] = anchor_x / pdf_page.width
        # O corte começa na altura do cabeçalho encontrados
        draw_start_y = (anchor_y / pdf_page.height) * im_height
        
    # Se não achou âncora, mas está ATIVO (Págs 2, 3, 4...), mantém.
    # E aplica uma margem superior pequena para não cortar cabeçalho de página (se houver)
    elif mask_state['active']:
        draw_start_y = im_height * 0.05 # 5% de margem superior segura

    # 3. APLICAÇÃO
    if mask_state['active'] and mask_state['cut_x_percent']:
        draw = ImageDraw.Draw(image, "RGBA")
        
        # X: Onde cortar (convertido para px)
        cut_x_px = mask_state['cut_x_percent'] * im_width
        safe_cut_x = cut_x_px + 5 # +5px para direita da palavra QTDE (margem)
        
        # Y: Altura
        # Proteção de Rodapé: Não desenha nos últimos 5% da página
        draw_end_y = im_height * 0.95 
        
        # Cores
        if DEBUG_MODE:
            fill = (255, 0, 0, 100)
            line = "red"
        else:
            fill = "white"
            line = "black"

        # Desenha
        draw.rectangle(
            [safe_cut_x, draw_start_y, im_width, draw_end_y],
            fill=fill, outline=None
        )
        
        draw.line([(safe_cut_x, draw_start_y), (safe_cut_x, draw_end_y)], fill=line, width=3)

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
            img, mask_state = apply_masking_v28(img, pdf_plumb.pages[i], mask_state)
        
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v28.0 (Anchor & Safety)</div>', unsafe_allow_html=True)
