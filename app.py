import streamlit as st
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile
import pdfplumber
from PIL import ImageDraw

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="SEI Converter ATA - SGB",
    page_icon="📑",
    layout="centered"
)

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

# --- TÍTULO PRINCIPAL ---
st.title("📑 SEI Converter ATA - SGB")

st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

# --- FUNÇÃO AUXILIAR DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    text = text.lower().strip()
    # Remove pontuação básica
    for ch in ['.', ':', '-', '/']:
        text = text.replace(ch, '')
    # Remove acentos
    replacements = {
        'ç': 'c', 'ã': 'a', 'á': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
        'í': 'i', 'ó': 'o', 'õ': 'o', 'ú': 'u'
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text

# --- FUNÇÃO DE MASCARAMENTO (v18.0 - SHAPE & CONTEXT LOCK) ---
def apply_masking_v18(image, pdf_page, mask_state):
    """
    mask_state:
      'active': bool (Máscara ligada?)
      'mask_x': float (Posição do corte)
      'ref_cols': int (Número de colunas da tabela de itens original)
      'last_bbox': list (Geometria anterior)
    """
    
    # Busca tabelas
    tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    if not tables:
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

    draw = ImageDraw.Draw(image)
    im_width, im_height = image.size
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    # --- DICIONÁRIOS DE PALAVRAS-CHAVE ---
    # Start: Variações de Quantidade
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "unid", "consumo"]
    
    # Start Backup: Variações de Preço (caso Qtde falhe)
    keys_price = ["preco", "unitario", "estimado", "valor", "total", "maximo"]
    
    # Stop: Termos que indicam Texto Jurídico/Cláusulas (Pág 7/8 Protection)
    keys_stop = [
        "local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", 
        "validade", "pagamento", "sancoes", "sançoes", "obrigacoes", "fiscalizacao", 
        "gestao", "clausula", "vigencia", "recursos", "dotacao", "objeto", "condicoes"
    ]

    for table in tables:
        if not table.rows: continue
        
        # --- ANÁLISE DE FORMA (SHAPE ANALYSIS) ---
        # Calcula o nº máximo de colunas nesta tabela (ignora linhas mescladas como "Grupo 01")
        max_row_cols = max([len(r.cells) for r in table.rows])
        
        # 1. PROTEÇÃO CONTRA TEXTO (Páginas 7 e 8)
        # Se a máscara estava ativa numa tabela de itens (ex: 5 colunas) 
        # e agora encontramos uma "tabela" de 1 ou 2 colunas, É TEXTO. PARAR.
        if mask_state['active'] and mask_state['ref_cols'] >= 3 and max_row_cols < 3:
            mask_state = {'active': False, 'mask_x': None, 'ref_cols': 0, 'last_bbox': None}
            continue

        # --- VARREDURA DE CONTEÚDO ---
        found_new_cut_x = None
        found_stopper = False
        
        # Varre até 8 linhas (para pular "Grupo 01", "Processo", etc e achar o Item)
        scan_limit = min(8, len(table.rows))
        
        for row_idx in range(scan_limit):
            row_cells = table.rows[row_idx].cells
            for cell_idx, cell in enumerate(row_cells):
                if not cell: continue
                try:
                    cropped = pdf_page.crop(cell)
                    text = clean_text(cropped.extract_text())
                    
                    # A) VERIFICAR STOPPER
                    if any(k in text for k in keys_stop):
                        found_stopper = True
                        break

                    # B) VERIFICAR START (Quantidade - Borda Direita)
                    # Verifica match exato ou palavra contida
                    if any(k == text or k in text.split() for k in keys_qty):
                        found_new_cut_x = cell[2]
                    
                    # C) VERIFICAR START BACKUP (Preço - Borda Esquerda)
                    elif found_new_cut_x is None and any(k in text for k in keys_price):
                         # Só aceita se estiver na metade direita da página
                         if cell[0] > (pdf_page.width * 0.4):
                            found_new_cut_x = cell[0]

                except:
                    pass
            if found_new_cut_x or found_stopper: break

        # --- MÁQUINA DE ESTADOS (STATE MACHINE) ---
        
        if found_stopper:
            # Encontrou texto jurídico -> DESLIGA TUDO
            mask_state = {'active': False, 'mask_x': None, 'ref_cols': 0, 'last_bbox': None}
        
        elif found_new_cut_x is not None:
            # Encontrou cabeçalho de Itens -> LIGA (Ajuste para Grupo 1 Pág 1)
            # Só aceita se a tabela tiver estrutura de itens (>= 3 colunas)
            if max_row_cols >= 3:
                mask_state['active'] = True
                mask_state['mask_x'] = found_new_cut_x
                mask_state['ref_cols'] = max_row_cols
                mask_state['last_bbox'] = table.bbox
        
        elif mask_state['active']:
            # MODO CONTINUAÇÃO (Sem cabeçalho novo)
            # Verifica se a tabela "parece" com a anterior
            if mask_state['last_bbox']:
                prev = mask_state['last_bbox']
                curr = table.bbox
                
                # Alinhamento horizontal parecido?
                aligned = abs(curr[0] - prev[0]) < 50
                # Largura parecida? (Evita confundir tabela larga com caixa de assinatura pequena)
                width_match = abs((curr[2]-curr[0]) - (prev[2]-prev[0])) < 50
                
                if aligned and width_match:
                    mask_state['last_bbox'] = table.bbox # Atualiza bbox
                else:
                    # Geometria mudou -> Fim da tabela
                    mask_state = {'active': False, 'mask_x': None, 'ref_cols': 0, 'last_bbox': None}

        # --- APLICAÇÃO VISUAL ---
        if mask_state['active'] and mask_state['mask_x'] is not None:
            cut_x = mask_state['mask_x']
            t_bbox = table.bbox
            
            # Garante que o corte está dentro dos limites horizontais da tabela
            if t_bbox[0] < cut_x < (t_bbox[2] + 50):
                
                x_pixel = cut_x * scale_x
                top_pixel = t_bbox[1] * scale_y
                bottom_pixel = t_bbox[3] * scale_y
                right_pixel_mask = im_width 
                
                # 1. Retângulo Branco (Apaga colunas de preço)
                draw.rectangle(
                    [x_pixel, top_pixel, right_pixel_mask, bottom_pixel],
                    fill="white", outline=None
                )

                # 2. Linha Preta (Fecha a tabela)
                draw.line(
                    [(x_pixel, top_pixel), (x_pixel, bottom_pixel)],
                    fill="black", width=3
                )
                
                # 3. Acabamento (Traços horizontais)
                draw.line([(x_pixel, top_pixel), (x_pixel - 5, top_pixel)], fill="black", width=2)
                draw.line([(x_pixel, bottom_pixel), (x_pixel - 5, bottom_pixel)], fill="black", width=2)

    return image, mask_state

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

    # ESTADO INICIAL (Limpo para cada arquivo)
    mask_state = {'active': False, 'mask_x': None, 'ref_cols': 0, 'last_bbox': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v18(img, pdf_plumb.pages[i], mask_state)
        
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

# --- PASSO 1: UPLOAD ---
uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui:", 
    type="pdf", 
    accept_multiple_files=True
)

# --- PASSO 2: PROCESSAR ---
if uploaded_files:
    st.write("---")
    if st.button(f"🚀 Processar {len(uploaded_files)} Arquivo(s)"):
        with st.spinner('Processando (Smart Shape Detection)...'):
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

# --- GUIA VISUAL ---
st.write("---")
st.subheader("📚 Guia Rápido: Como inserir no SEI")

col1, col2 = st.columns([0.15, 0.85])
with col1:
    try:
        st.image("icone_sei.png", width=50) 
    except:
        st.write("🧩")
with col2:
    st.markdown("""
    *1º Localize o ícone:* No editor do SEI, clique no botão da função *INSERIR CONTEÚDO EXTERNO* (representado pelo ícone ao lado).
    """)

st.write("")

st.markdown("""
*2º Configure a inserção:* Na janela que abrir, faça o upload do arquivo Word gerado aqui.
""")

st.warning("⚠️ *IMPORTANTE:* Certifique-se de deixar todas as caixas de seleção *DESMARCADAS*.")

try:
    st.image("print_sei.png", caption="Exemplo: Deixe as opções desmarcadas.", use_container_width=True)
except:
    pass

# --- RODAPÉ ---
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v18.0 (Shape & Context Lock)</div>', unsafe_allow_html=True)
