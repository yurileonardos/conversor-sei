import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile
import pdfplumber
from PIL import Image, ImageDraw

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="SEI Converter ATA - SGB",
    page_icon="📑",
    layout="centered"
)

# --- MODO DIAGNÓSTICO ---
# True = Vermelho (Teste) | False = Branco (Produção)
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

# --- FUNÇÃO DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    text = str(text).lower().strip()
    for ch in ['.', ':', '-', '/']:
        text = text.replace(ch, '')
    replacements = {
        'ç': 'c', 'ã': 'a', 'á': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
        'í': 'i', 'ó': 'o', 'õ': 'o', 'ú': 'u'
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text

# --- FUNÇÃO DE MASCARAMENTO (v20.0 - FLUXO PERSISTENTE) ---
def apply_masking_v20(image, pdf_page, mask_state):
    """
    mask_state:
      'active': bool
      'cut_x_percent': float (Posição relativa do corte 0.0 a 1.0)
    """
    
    # Busca todas as tabelas possíveis (Linhas e Texto)
    # Mesclamos as estratégias para garantir que nada passe batido
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    
    # Prioriza 'lines', mas se não achar, usa 'text'. 
    # Porém, para o Grupo 1 que pode estar quebrado, vamos iterar sobre o que tiver disponível.
    all_tables = tables_lines if tables_lines else tables_text

    draw = ImageDraw.Draw(image, "RGBA") 
    im_width, im_height = image.size
    
    # Palavras-chave
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "unid", "consumo", "catmat"]
    keys_price = ["preco", "unitario", "estimado", "valor", "total", "maximo"]
    
    # STOPPERS (Expandido para proteger Páginas 7/8)
    keys_stop = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sancoes", "sançoes", "obrigacoes", "fiscalizacao", 
        "gestao", "clausula", "vigencia", "recursos", "dotacao", "condicoes", "multas", 
        "infracoes", "penalidades", "rescisao", "foro"
    ]

    for table in all_tables:
        if not table.rows: continue
        
        # Geometria da tabela atual
        t_bbox = table.bbox # (x0, top, x1, bottom)
        t_width = t_bbox[2] - t_bbox[0]
        
        # --- ANÁLISE DE CONTEÚDO ---
        found_start_x = None
        found_stopper = False
        text_content = ""
        
        # Varredura Profunda (Deep Scan) nas primeiras linhas
        limit_rows = min(8, len(table.rows))
        for r_idx in range(limit_rows):
            for cell in table.rows[r_idx].cells:
                if not cell: continue
                try:
                    # Extrai texto (com proteção contra erro de crop)
                    if isinstance(cell, (list, tuple)) and len(cell) == 4:
                        crop = pdf_page.crop(cell)
                        txt = clean_text(crop.extract_text())
                        text_content += txt + " "
                        
                        # 1. Verifica STOPPER
                        if any(k in txt for k in keys_stop):
                            found_stopper = True
                        
                        # 2. Verifica START (Qtde -> Direita)
                        if any(k == txt or k in txt.split() for k in keys_qty):
                            found_start_x = cell[2] # Borda Direita
                        
                        # 3. Verifica START BACKUP (Preço -> Esquerda)
                        elif found_start_x is None and any(k in txt for k in keys_price):
                            # Validação: Preço deve estar na direita da página
                            if cell[0] > (pdf_page.width * 0.4):
                                found_start_x = cell[0]
                except:
                    pass
            if found_start_x or found_stopper: break

        # --- LÓGICA DE DECISÃO (PERSISTÊNCIA) ---
        
        # CASO 1: ENCONTROU PARADA (Stopper)
        if found_stopper:
            mask_state['active'] = False
            mask_state['cut_x_percent'] = None
        
        # CASO 2: COLAPSO ESTRUTURAL (Proteção Pág 7/8)
        # Se a tabela tem 1 coluna e MUITO texto, é parágrafo, não tabela de itens.
        # Mas só desliga se a máscara estava ativa.
        elif mask_state['active']:
            cols_count = max([len(r.cells) for r in table.rows])
            # Se caiu para 1 coluna e tem texto longo (> 50 chars), é texto corrido
            if cols_count < 3 and len(text_content) > 50:
                mask_state['active'] = False
                mask_state['cut_x_percent'] = None

        # CASO 3: ENCONTROU INÍCIO (Header)
        if found_start_x is not None and not found_stopper:
            mask_state['active'] = True
            # Grava a porcentagem da largura para aplicar proporcionalmente na imagem
            mask_state['cut_x_percent'] = found_start_x / pdf_page.width

        # --- APLICAÇÃO ---
        # Se ativo (seja por novo header ou persistência do anterior)
        if mask_state['active'] and mask_state['cut_x_percent']:
            
            # Converte % de volta para pixels da imagem
            cut_x_pixel = mask_state['cut_x_percent'] * im_width
            
            scale_y = im_height / pdf_page.height
            top_pixel = t_bbox[1] * scale_y
            bottom_pixel = t_bbox[3] * scale_y
            
            # Ajuste fino: Se o corte calculado estiver ANTES do inicio da tabela, ignora (erro de cálculo)
            t_x0_pixel = t_bbox[0] * (im_width / pdf_page.width)
            if cut_x_pixel > t_x0_pixel:

                # CORES
                if DEBUG_MODE:
                    fill = (255, 0, 0, 100) # Vermelho Transparente
                    line = "red"
                else:
                    fill = "white"
                    line = "black"

                # DESENHO
                # Retângulo vai do corte até o fim da imagem (direita)
                draw.rectangle(
                    [cut_x_pixel, top_pixel, im_width, bottom_pixel],
                    fill=fill, outline=None
                )
                
                # Linha Vertical
                draw.line([(cut_x_pixel, top_pixel), (cut_x_pixel, bottom_pixel)], fill=line, width=3)
                
                # Linhas Horizontais (Fechamento)
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

    # ESTADO INICIAL GLOBAL DO ARQUIVO
    mask_state = {'active': False, 'cut_x_percent': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v20(img, pdf_plumb.pages[i], mask_state)
        
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
    btn_txt = "🚀 Processar (Modo Diagnóstico)" if DEBUG_MODE else "🚀 Processar Arquivos"
    if st.button(btn_txt):
        with st.spinner('Processando tabelas com Fluxo Persistente...'):
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

st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v20.0 (Persistent Flow)</div>', unsafe_allow_html=True)
