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
# True = Máscara Vermelha (Para calibrar) | False = Máscara Branca (Final)
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
    st.warning("🔴 MODO DIAGNÓSTICO: Máscaras em VERMELHO. Se funcionar, mudamos para branco.")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÕES DE LÓGICA ---

def clean_text(text):
    if not text: return ""
    return str(text).strip()

def is_strict_decimal(text):
    """
    Verifica se é ESTRITAMENTE um valor numérico com 2 casas decimais.
    Padrão aceito: digitos + virgula + 2 digitos.
    Ex: 100,00 | 1.200,50 | 0,21 | 122,33
    Ignora: 50 (inteiro), 1.200 (sem virgula), Texto
    """
    if not text: return False
    # Regex: 
    # [\d\.]+  -> Procura dígitos ou pontos (milhar)
    # ,        -> Vírgula obrigatória
    # \d{2}    -> Exatamente 2 dígitos depois
    # \b       -> Fim da palavra (evita pegar texto longo)
    match = re.search(r'[\d\.]*,\d{2}\b', text)
    
    # Filtro extra: rejeita se tiver muitas letras (evita falsos positivos em texto jurídico)
    if match:
        # Conta letras na string original
        letters = sum(c.isalpha() for c in text if c.lower() not in ['r', 's', '$'])
        if letters > 3: return False # Se tem muita letra, não é preço puro
        return True
    return False

# --- FUNÇÃO DE MASCARAMENTO (v23.0 - STRICT NUMERIC SCAN) ---
def apply_masking_v23(image, pdf_page, mask_state):
    
    # Busca todas as tabelas
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    # Palavras de Parada (Segurança Pág 7/8)
    keys_stop = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sancoes", "sançoes", "obrigacoes", "fiscalizacao", 
        "gestao", "clausula", "vigencia", "recursos", "dotacao", "objeto", "condicoes",
        "multas", "infracoes", "penalidades", "rescisao", "foro"
    ]

    for table in all_tables:
        if not table.rows: continue
        
        t_bbox = table.bbox
        
        # Variáveis de detecção
        found_decimal_x = None
        found_stopper = False
        row_count = len(table.rows)
        
        # --- 1. DETECÇÃO DE STOPPER (Texto) ---
        # Analisa amostra de texto para ver se mudou o assunto
        limit_check = min(5, row_count)
        for r_idx in range(limit_check):
            for cell in table.rows[r_idx].cells:
                try:
                    if cell and isinstance(cell, (list, tuple)):
                        crop = pdf_page.crop(cell)
                        txt = str(crop.extract_text()).lower()
                        if any(k in txt for k in keys_stop):
                            found_stopper = True
                except: pass
            if found_stopper: break

        # --- 2. DETECÇÃO NUMÉRICA (O Coração da Lógica) ---
        # Se não for texto jurídico, varre a tabela procurando a coluna de decimais
        if not found_stopper:
            
            # Varre até 15 linhas para garantir que pega números
            limit_scan = min(15, row_count)
            min_x_found = None # Queremos a coluna mais à esquerda que tenha decimais
            
            for r_idx in range(limit_scan):
                for cell in table.rows[r_idx].cells:
                    try:
                        if cell and isinstance(cell, (list, tuple)):
                            crop = pdf_page.crop(cell)
                            raw_txt = clean_text(crop.extract_text())
                            
                            # APLICANDO A REGRA ESTRITA DE 2 CASAS DECIMAIS
                            if is_strict_decimal(raw_txt):
                                cell_x = cell[0] # Borda esquerda da célula
                                
                                # Validação de Posição:
                                # Preços geralmente estão da metade para a direita da tabela
                                # Ignora se estiver muito à esquerda (ex: Quantidade 1,00 na esquerda? raro, mas possível)
                                if cell_x > (pdf_page.width * 0.4):
                                    if min_x_found is None or cell_x < min_x_found:
                                        min_x_found = cell_x
                    except: pass
            
            if min_x_found is not None:
                found_decimal_x = min_x_found

        # --- 3. ATUALIZAÇÃO DO ESTADO ---
        
        # A) Texto Jurídico detectado -> Desliga Máscara
        if found_stopper:
            mask_state['active'] = False
            mask_state['cut_x_percent'] = None
        
        # B) Números decimais detectados -> Liga/Atualiza Máscara
        elif found_decimal_x is not None:
            mask_state['active'] = True
            mask_state['cut_x_percent'] = found_decimal_x / pdf_page.width
            
        # C) Proteção Estrutural (Pág 7/8)
        # Se a tabela tem poucas colunas (<3) e a máscara estava ativa, verifica se ainda tem números
        elif mask_state['active']:
            cols_count = max([len(r.cells) for r in table.rows])
            if cols_count < 3:
                # Se caiu o número de colunas drasticamente, assume que virou texto
                mask_state['active'] = False
                mask_state['cut_x_percent'] = None

        # --- 4. APLICAÇÃO ---
        if mask_state['active'] and mask_state['cut_x_percent']:
            
            # Define pixel de corte
            cut_x_pixel = mask_state['cut_x_percent'] * im_width
            
            # Limites verticais da tabela atual
            scale_y = im_height / pdf_page.height
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Validação: Corte deve ser após o inicio da tabela
            t_x0_pixel = t_bbox[0] * (im_width / pdf_page.width)
            
            if cut_x_pixel > t_x0_pixel:
                
                # CORES
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100)
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # DESENHA MÁSCARA
                draw.rectangle(
                    [cut_x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )
                
                # LINHA VERTICAL
                draw.line([(cut_x_pixel, top_pixel), (cut_x_pixel, bottom_pixel)], fill=line, width=3)
                
                # ACABAMENTO
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
            img, mask_state = apply_masking_v23(img, pdf_plumb.pages[i], mask_state)
        
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
    btn_label = "🚀 Processar (Vermelho - Teste)" if DEBUG_MODE else "🚀 Processar Arquivos"
    
    if st.button(btn_label):
        with st.spinner('Aplicando máscara em valores decimais...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v23.0 (Numeric Strict)</div>', unsafe_allow_html=True)
