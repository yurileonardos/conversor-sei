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
    # Remove pontuação básica para match
    for ch in ['.', ':', '-', '/']:
        text = text.replace(ch, '')
    # Remove acentos para garantir (ex: "Preço" vira "Preco")
    replacements = {
        'ç': 'c', 'ã': 'a', 'á': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
        'í': 'i', 'ó': 'o', 'õ': 'o', 'ú': 'u'
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text

# --- FUNÇÃO DE MASCARAMENTO (v15.0 - ESTRATÉGIA MISTA & TRAVA DE TEXTO) ---
def apply_masking_v15(image, pdf_page, mask_state):
    """
    mask_state guarda: {'mask_x': float, 'last_bbox': list, 'strategy': str, 'cols': int}
    """
    
    # 1. Tenta achar tabelas com LINHAS (Alta precisão)
    tables_lines = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    
    # 2. Tenta achar tabelas por TEXTO (Baixa precisão, para tabelas sem borda)
    tables_text = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    
    # Decide qual usar nesta página
    if tables_lines:
        current_tables = tables_lines
        current_strategy = 'lines'
    else:
        current_tables = tables_text
        current_strategy = 'text'

    draw = ImageDraw.Draw(image)
    im_width, im_height = image.size
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    # PALAVRAS-CHAVE
    # Grupo 1: Cortar à DIREITA da Quantidade
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "quantitativo"]
    
    # Grupo 2: Cortar à ESQUERDA do Preço (Backup se Qtde falhar ou para pegar início exato)
    keys_price = ["preco", "unitario", "estimado", "valor", "total", "maximo", "ref", "medio"]
    
    # Grupo 3: Parar máscara (Stopper)
    keys_stop = ["local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante", "validade", "pagamento"]

    for table in current_tables:
        if not table.rows: continue
        
        # Ignora tabelas com menos de 3 colunas (quase sempre é texto/layout)
        num_cols = max([len(r.cells) for r in table.rows])
        if num_cols < 3:
            mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}
            continue

        # --- ANÁLISE DE CABEÇALHO ---
        cut_x = None
        found_stopper = False
        
        # Analisa até 5 linhas para garantir
        for row_idx in range(min(5, len(table.rows))):
            row_cells = table.rows[row_idx].cells
            for cell_idx, cell in enumerate(row_cells):
                if not cell: continue
                try:
                    cropped = pdf_page.crop(cell)
                    text = clean_text(cropped.extract_text())
                    
                    # 1. STOPPER
                    if any(k in text for k in keys_stop):
                        found_stopper = True
                        break

                    # 2. PREÇO (Esquerda) - Prioridade alta para pegar o Grupo 1 se Qtde falhar
                    # Se achar "Preço Unitário", corta na esquerda dele
                    if any(k in text for k in keys_price):
                        # Validação: Preço geralmente está na metade direita da tabela
                        # Se estiver muito à esquerda, pode ser falso positivo
                        if cell[0] > table.bbox[0] + (table.bbox[2] - table.bbox[0]) * 0.4:
                            cut_x = cell[0] # Borda ESQUERDA
                            break

                    # 3. QUANTIDADE (Direita)
                    if cut_x is None and any(k == text or k in text.split() for k in keys_qty):
                        cut_x = cell[2] # Borda DIREITA
                        
                except:
                    pass
            if cut_x or found_stopper: break

        # --- ATUALIZAÇÃO DO ESTADO ---
        active_cut_x = None

        if found_stopper:
            # Encontrou tabela de texto (Local, Prazo) -> Reseta tudo
            mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}
        
        elif cut_x is not None:
            # ACHOU CABEÇALHO NOVO!
            mask_state['mask_x'] = cut_x
            mask_state['last_bbox'] = table.bbox
            mask_state['strategy'] = current_strategy
            mask_state['cols'] = num_cols
            active_cut_x = cut_x
        
        else:
            # SEM CABEÇALHO (Possível Continuação)
            if mask_state['mask_x'] is not None and mask_state['last_bbox']:
                
                # --- CHECAGEM RIGOROSA DE CONTINUIDADE ---
                prev = mask_state['last_bbox']
                curr = table.bbox
                
                # 1. Checagem de Estratégia (CRÍTICO PARA A PÁGINA 7)
                # Se a anterior era 'lines' (tabela real) e a atual é 'text' (texto solto), NÃO É CONTINUAÇÃO.
                if mask_state['strategy'] == 'lines' and current_strategy == 'text':
                    active_cut_x = None
                    # Reseta para evitar danos futuros
                    mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}
                
                # 2. Checagem de Colunas
                # Se o número de colunas mudou drasticamente, não é a mesma tabela
                elif abs(num_cols - mask_state['cols']) > 2:
                    active_cut_x = None
                    mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}
                
                # 3. Checagem Geométrica (Alinhamento)
                # Alinhamento esquerdo e largura similares
                elif abs(curr[0] - prev[0]) < 50 and abs((curr[2]-curr[0]) - (prev[2]-prev[0])) < 50:
                    active_cut_x = mask_state['mask_x']
                    mask_state['last_bbox'] = table.bbox # Atualiza bbox
                    # Mantém estratégia e cols da original
                else:
                    # Desalinhou -> Reseta
                    mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}

        # --- APLICAÇÃO DA MÁSCARA ---
        if active_cut_x is not None:
            t_bbox = table.bbox
            
            # Segurança: Corte deve estar dentro da tabela
            if t_bbox[0] < active_cut_x < t_bbox[2]:
                
                x_pixel = active_cut_x * scale_x
                top_pixel = t_bbox[1] * scale_y
                bottom_pixel = t_bbox[3] * scale_y
                # Vai até a direita da IMAGEM (para cobrir vazamentos) mas visualmente fecha na tabela
                right_pixel_mask = im_width 
                
                # 1. Retângulo Branco (Apaga os dados)
                draw.rectangle(
                    [x_pixel, top_pixel, right_pixel_mask, bottom_pixel],
                    fill="white", outline=None
                )

                # 2. Linha Preta (Fecha a tabela visualmente)
                # Linha vertical grossa
                draw.line(
                    [(x_pixel, top_pixel), (x_pixel, bottom_pixel)],
                    fill="black", width=3
                )
                
                # Linhas de acabamento superior/inferior (pequenos traços para a esquerda)
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

    # Estado Inicial
    mask_state = {'mask_x': None, 'last_bbox': None, 'strategy': None, 'cols': 0}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v15(img, pdf_plumb.pages[i], mask_state)
        
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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v15.0 (Precision Fix)</div>', unsafe_allow_html=True)
