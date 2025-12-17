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
from collections import Counter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="SEI Converter ATA - SGB",
    page_icon="📑",
    layout="centered"
)

# --- MODO DIAGNÓSTICO ---
# True = Vermelho (Teste) | False = Branco (Final)
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
    st.warning("🔴 MODO DIAGNÓSTICO: Máscaras em VERMELHO. Se funcionar nas págs 1-4, me avise para finalizar em branco.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES LÓGICAS ---

def is_price_format(text):
    """
    Detecta formato numérico estrito:
    - Deve ter vírgula
    - Deve ter exatamente 2 dígitos após a vírgula
    - Pode ter R$ ou pontos de milhar
    Ex: 100,00 | 1.200,50 | R$ 50,00 | 0,21
    """
    if not text: return False
    clean = text.strip().replace(" ", "")
    # Regex: (Opcional R$) + (Digitos/Pontos) + Virgula + 2 Digitos + Fim da string
    match = re.search(r'(?:R\$)?[\d\.]*,\d{2}$', clean)
    if match:
        # Rejeita se tiver letras no meio (ex: Lei nº 10.520,02 - falso positivo raro, mas possivel)
        # Conta caracteres que não são numeros nem pontuação de preço
        invalids = sum(1 for c in clean if c.lower() not in '0123456789.,r$')
        return invalids == 0
    return False

def find_price_column_x(pdf_page):
    """
    Analisa TODAS as palavras da página (sem depender de tabelas).
    Procura alinhamento vertical de números decimais.
    Retorna a coordenada X da coluna de preço mais à esquerda encontrada.
    """
    words = pdf_page.extract_words()
    
    # Filtra apenas palavras que parecem preços
    price_words = [w for w in words if is_price_format(w['text'])]
    
    if not price_words:
        return None
        
    # Agrupa por coordenadas X (com tolerância de 5pts para desalinhamentos leves)
    # Cria "buckets" de colunas
    x_clusters = []
    tolerance = 5
    
    for w in price_words:
        x0 = w['x0']
        # Tenta encaixar em um cluster existente
        found_cluster = False
        for cluster in x_clusters:
            # Se a média do cluster está perto deste x0
            if abs(cluster['avg_x'] - x0) < tolerance:
                cluster['points'].append(x0)
                cluster['avg_x'] = sum(cluster['points']) / len(cluster['points'])
                cluster['count'] += 1
                found_cluster = True
                break
        
        if not found_cluster:
            x_clusters.append({'avg_x': x0, 'points': [x0], 'count': 1})
            
    # Filtra clusters relevantes (precisa ter pelo menos 2 preços alinhados para ser uma coluna)
    # E precisa estar na metade direita da página (> 40% da largura) para evitar falsos positivos
    page_width = pdf_page.width
    valid_clusters = [c for c in x_clusters if c['count'] >= 2 and c['avg_x'] > (page_width * 0.4)]
    
    if not valid_clusters:
        return None
        
    # Pega o cluster mais à esquerda (menor X) - Provavelmente "Preço Unitário"
    best_cluster = min(valid_clusters, key=lambda c: c['avg_x'])
    return best_cluster['avg_x']

def check_for_stoppers(pdf_page):
    """Verifica se há palavras de parada na página"""
    text = pdf_page.extract_text().lower()
    keys_stop = [
        "local de entrega", "prazo de entrega", "assinatura do contrato", 
        "garantia dos bens", "sanções administrativas", "obrigações da contratada", 
        "fiscalização", "gestão do contrato", "cláusula", "vigência", "dotação orçamentária"
    ]
    return any(k in text for k in keys_stop)

# --- FUNÇÃO DE MASCARAMENTO (v24.0 - VISUAL COLUMN DETECTOR) ---
def apply_masking_v24(image, pdf_page, mask_state):
    
    im_width, im_height = image.size
    
    # 1. VERIFICA STOPPER (Texto Jurídico)
    # Se encontrar, desliga a máscara imediatamente
    if check_for_stoppers(pdf_page):
        mask_state['active'] = False
        mask_state['cut_x_percent'] = None
    
    # 2. DETECÇÃO DE COLUNA (Se não houver stopper)
    else:
        # Tenta achar uma coluna de preços na página atual
        found_x = find_price_column_x(pdf_page)
        
        if found_x:
            # ACHOU! Ativa a máscara e atualiza a posição
            mask_state['active'] = True
            mask_state['cut_x_percent'] = found_x / pdf_page.width
            
        # Se não achou nesta página, mas o estado estava ATIVO...
        # Mantemos ativo (herança), a menos que a página pareça estar vazia ou muito diferente.
        # (Nesta versão simplificada, confiamos no Stopper para desligar)

    # 3. APLICAÇÃO VISUAL
    if mask_state['active'] and mask_state['cut_x_percent']:
        draw = ImageDraw.Draw(image, "RGBA")
        
        # Converte % para pixels
        cut_x_pixel = mask_state['cut_x_percent'] * im_width
        
        # Define área de corte: Da linha detectada até o fim da página
        # Margem de segurança: Recua um pouco (ex: -10px) para garantir que cobre o número todo
        safe_cut_x = cut_x_pixel - 5 
        
        # Cores
        if DEBUG_MODE:
            fill = (255, 0, 0, 100) # Vermelho
            line = "red"
        else:
            fill = "white"
            line = "black"

        # Desenha Máscara (Na página inteira, respeitando margens verticais se necessário, 
        # mas aqui vamos simplificar para cobrir a coluna verticalmente)
        draw.rectangle(
            [safe_cut_x, 0, im_width, im_height],
            fill=fill, outline=None
        )
        
        # Linha Vertical
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
            img, mask_state = apply_masking_v24(img, pdf_plumb.pages[i], mask_state)
        
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
        with st.spinner('Escaneando alinhamento vertical de preços...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v24.0 (Visual Column Detector)</div>', unsafe_allow_html=True)
