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
    # Remove acentos básicos manualmente para garantir match
    text = text.lower().strip()
    replacements = {
        'ç': 'c', 'ã': 'a', 'á': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
        'í': 'i', 'ó': 'o', 'õ': 'o', 'ú': 'u', '.': '', ':': ''
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text

# --- FUNÇÃO DE MASCARAMENTO (v14.0 - GEOMETRIA RÍGIDA) ---
def apply_masking_v14(image, pdf_page, mask_state):
    """
    Correções:
    1. Só aplica em tabelas com >= 3 colunas (Evita mascarar texto solto).
    2. Busca Qtde (Direita) OU Preço (Esquerda) para pegar o Grupo 1.
    3. Máscara restrita à altura da tabela (bbox).
    """
    
    # Busca tabelas usando linhas (melhor para TRs)
    tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    if not tables:
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

    draw = ImageDraw.Draw(image)
    im_width, im_height = image.size
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    # PALAVRAS-CHAVE
    # Âncora Principal: Corta à DIREITA destas
    keys_qty = ["qtde", "qtd", "quantidade", "quant", "quantitativo"]
    
    # Âncora Secundária (Backup): Corta à ESQUERDA destas
    keys_price = ["preco", "unitario", "estimado", "valor", "total", "maximo"]
    
    # Parada de Segurança
    keys_stop = ["local", "entrega", "prazo", "assinatura", "garantia", "marca", "fabricante"]

    for table in tables:
        if not table.rows: continue
        
        # --- TRAVA DE SEGURANÇA 1: NÚMERO DE COLUNAS ---
        # Tabelas de itens geralmente têm: Item, Desc, Unid, Qtd, Valor... (5+ colunas).
        # Se tiver menos de 3 colunas, provavelmente é texto ou layout, não mascaramos.
        # Pegamos a linha com mais células para checar.
        max_cols = max([len(r.cells) for r in table.rows])
        if max_cols < 3:
            # Se for uma tabela "fina" (texto), desativa a memória para não riscar a página
            mask_state['mask_x'] = None
            continue

        # --- VARREDURA DE CABEÇALHO ---
        cut_x = None
        found_stopper = False
        
        # Aumentei para analisar as primeiras 5 linhas (para pegar tabelas com títulos mesclados)
        for row_idx in range(min(5, len(table.rows))):
            row_cells = table.rows[row_idx].cells
            for cell_idx, cell in enumerate(row_cells):
                if not cell: continue
                try:
                    cropped = pdf_page.crop(cell)
                    text_raw = cropped.extract_text()
                    text = clean_text(text_raw)
                    
                    # 1. Verifica STOPPER (Mudou o assunto?)
                    if any(k in text for k in keys_stop):
                        found_stopper = True
                        break

                    # 2. Verifica QUANTIDADE (Prioridade) -> Pega borda DIREITA (cell[2])
                    if any(k == text or k in text.split() for k in keys_qty):
                        cut_x = cell[2]
                        break # Achamos o ponto exato

                    # 3. Verifica PREÇO (Backup) -> Pega borda ESQUERDA (cell[0])
                    # Só usa se ainda não achou Qtde e se a palavra "preço" ou "unitário" estiver clara
                    if cut_x is None and any(k in text for k in keys_price):
                        # Validação extra: "Preço" geralmente está na coluna 4 ou 5
                        cut_x = cell[0]
                        
                except:
                    pass
            if cut_x or found_stopper: break

        # --- ATUALIZAÇÃO DE ESTADO ---
        active_cut_x = None

        if found_stopper:
            # Encontrou tabela de "Local de Entrega", etc.
            mask_state['mask_x'] = None
            mask_state['last_table_bbox'] = None
        
        elif cut_x is not None:
            # Encontrou cabeçalho novo válido
            mask_state['mask_x'] = cut_x
            mask_state['last_table_bbox'] = table.bbox
            active_cut_x = cut_x
        
        elif mask_state['mask_x'] is not None:
            # Continuação (sem cabeçalho)
            # Verifica alinhamento geométrico para não mascarar coisas erradas
            if mask_state['last_table_bbox']:
                prev = mask_state['last_table_bbox']
                curr = table.bbox
                # Se a tabela tiver largura parecida (+- 5%) e alinhamento esquerdo parecido
                if abs(curr[0] - prev[0]) < 50 and abs((curr[2]-curr[0]) - (prev[2]-prev[0])) < 50:
                    active_cut_x = mask_state['mask_x']
                    mask_state['last_table_bbox'] = table.bbox
                else:
                    # Formato mudou muito, cancela máscara
                    mask_state['mask_x'] = None

        # --- APLICAÇÃO DA MÁSCARA (RESTRIÇÃO BBOX) ---
        if active_cut_x is not None:
            t_bbox = table.bbox # (x0, top, x1, bottom)
            
            # Verificação final: O corte deve estar DENTRO da largura da tabela
            if t_bbox[0] < active_cut_x < t_bbox[2]:
                
                # Coordenadas ajustadas para escala da imagem
                x_pixel = active_cut_x * scale_x
                top_pixel = t_bbox[1] * scale_y
                bottom_pixel = t_bbox[3] * scale_y
                right_pixel = im_width # Vai até o fim da folha para garantir
                
                # 1. Desenha o Retângulo Branco
                # Note que usamos top_pixel e bottom_pixel da TABELA ATUAL
                # Isso impede que a máscara invada o cabeçalho da página ou rodapé fora da tabela
                draw.rectangle(
                    [x_pixel, top_pixel, right_pixel, bottom_pixel],
                    fill="white",
                    outline=None
                )

                # 2. Desenha a Linha de Fechamento (Preta)
                draw.line(
                    [(x_pixel, top_pixel), (x_pixel, bottom_pixel)],
                    fill="black",
                    width=3
                )
                
                # Linhas horizontais de acabamento (só um pouquinho para a esquerda)
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
    
    # Margens
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(0.5)

    mask_state = {'mask_x': None, 'last_table_bbox': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v14(img, pdf_plumb.pages[i], mask_state)
        
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
        with st.spinner('Processando tabelas com ajuste geométrico...'):
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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v14.0 (Hybrid Precision)</div>', unsafe_allow_html=True)
