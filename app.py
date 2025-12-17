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

# --- INTRODUÇÃO ---
st.markdown("""
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas, 
a fim de inseri-las no documento SEI: **ATA DE REGISTRO DE PREÇOS**.
""")

col1, col2 = st.columns([0.1, 0.9])
with col1:
    try:
        st.image("icone_sei.png", width=40)
    except:
        st.write("🧩")
with col2:
    st.info("""
    Funcionalidade disponível na extensão [**SEI PRO**](https://sei-pro.github.io/sei-pro/), 
    utilizando a ferramenta [**INSERIR CONTEÚDO EXTERNO**](https://sei-pro.github.io/sei-pro/pages/INSERIRDOC.html).
    """)

with st.expander("⚙️ Deseja escolher a pasta onde o arquivo será salvo? Clique aqui."):
    st.markdown("""
    Por segurança, os navegadores salvam automaticamente na pasta "Downloads". 
    Para escolher a pasta a cada download, configure seu navegador (Chrome/Edge):
    1. Vá em **Configurações** > **Downloads**.
    2. Ative: **"Perguntar onde salvar cada arquivo antes de fazer download"**.
    """)

st.write("---")

# --- PASSO 1: UPLOAD ---
st.write("### Passo 1: Upload dos Arquivos")
st.markdown("**Nota:** O sistema aplicará a máscara automática para ocultar preços em Termos de Referência.")

uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui (ou clique para buscar):", 
    type="pdf", 
    accept_multiple_files=True
)

# --- FUNÇÃO AUXILIAR DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    text = text.replace('\n', ' ').replace('\r', ' ')
    return text.lower().strip()

# --- FUNÇÃO DE MASCARAMENTO INTELIGENTE (v10 - COM MEMÓRIA DE CONTINUIDADE) ---
def apply_masking_v10(image, pdf_page, mask_state):
    """
    mask_state: dicionário contendo {'mask_x': float/None, 'last_table_bbox': list/None}
    Retorna: image, mask_state atualizado
    """
    try:
        # Tenta achar tabelas com linhas
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        # Se falhar, tenta achar por texto
        if not tables:
             tables = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

        draw = ImageDraw.Draw(image)
        im_width, im_height = image.size
        
        scale_x = im_width / pdf_page.width
        scale_y = im_height / pdf_page.height

        keywords_target = [
            "preco unit", "preço unit", "valor unit", "vlr. unit", "unitario", "unitário",
            "valor max", "valor estim", "preço estim", "preco estim", "valor ref", 
            "vlr total", "valor total", "preco total", "preço total"
        ]
        
        for table in tables:
            if not table.rows: continue

            # --- 1. VERIFICAR SE EXISTE CABEÇALHO NOVO ---
            found_new_header_mask_x = None
            
            # Varre as primeiras 3 linhas para tentar achar keywords
            for row_idx in range(min(3, len(table.rows))):
                row_cells = table.rows[row_idx].cells
                for cell_idx, cell in enumerate(row_cells):
                    if not cell: continue
                    try:
                        cropped = pdf_page.crop(cell)
                        text_raw = cropped.extract_text()
                        text_clean = clean_text(text_raw)
                        
                        if any(k in text_clean for k in keywords_target):
                            # Achou um cabeçalho de preço!
                            found_new_header_mask_x = cell[0] 
                            break 
                    except:
                        pass
                if found_new_header_mask_x is not None:
                    break
            
            # --- 2. LÓGICA DE DECISÃO (MÁQUINA DE ESTADOS) ---
            active_mask_x = None

            if found_new_header_mask_x is not None:
                # CASO A: Novo cabeçalho de preço detectado
                # Atualiza a memória com a nova posição
                mask_state['mask_x'] = found_new_header_mask_x
                mask_state['last_table_bbox'] = table.bbox
                active_mask_x = found_new_header_mask_x
            
            else:
                # CASO B: Nenhum cabeçalho de preço encontrado
                # Verifica se podemos usar a memória (Continuação de Tabela)
                if mask_state['mask_x'] is not None and mask_state['last_table_bbox'] is not None:
                    
                    # Checagem Geométrica: A tabela atual parece com a anterior?
                    # Tolerância de 30 pontos na largura/posição
                    prev_bbox = mask_state['last_table_bbox']
                    curr_bbox = table.bbox
                    
                    same_left_align = abs(curr_bbox[0] - prev_bbox[0]) < 30
                    same_width = abs((curr_bbox[2]-curr_bbox[0]) - (prev_bbox[2]-prev_bbox[0])) < 30
                    
                    if same_left_align: # Se estiver alinhada à esquerda, assumimos continuação
                        active_mask_x = mask_state['mask_x']
                        # Atualiza o bbox para a próxima página comparar com esta
                        mask_state['last_table_bbox'] = table.bbox
                    else:
                        # Tabela muito diferente (provavelmente nova tabela sem preço)
                        # Reseta a memória para não mascarar errado
                        mask_state['mask_x'] = None
                        mask_state['last_table_bbox'] = None
                else:
                    # Sem memória e sem cabeçalho -> Não faz nada
                    pass

            # --- 3. APLICAR MÁSCARA SE DEFINIDO ---
            if active_mask_x is not None:
                table_rect = table.bbox
                
                # Verifica se o ponto de corte está dentro da tabela (segurança)
                if active_mask_x < table_rect[2]:
                    # A) Máscara Branca
                    rect_mask = [
                        active_mask_x * scale_x,       
                        table_rect[1] * scale_y,      
                        table_rect[2] * scale_x + 50, 
                        table_rect[3] * scale_y       
                    ]
                    draw.rectangle(rect_mask, fill="white", outline=None)

                    # B) Borda Preta (Corte Visual)
                    rect_border = [
                        table_rect[0] * scale_x,
                        table_rect[1] * scale_y,
                        active_mask_x * scale_x, # Limite visual
                        table_rect[3] * scale_y
                    ]
                    draw.rectangle(rect_border, outline="black", width=2)
                
                # --- 4. LIMPEZA DE RODAPÉ (TOTAL) ---
                # Apenas se a palavra "total" aparecer na última linha
                last_row = table.rows[-1]
                try:
                    first_cell = last_row.cells[0]
                    if first_cell:
                        cropped_last = pdf_page.crop(first_cell)
                        last_text = clean_text(cropped_last.extract_text())
                        
                        if "total" in last_text:
                            tops = [c[1] for c in last_row.cells if c]
                            bottoms = [c[3] for c in last_row.cells if c]
                            if tops and bottoms:
                                l_top = min(tops)
                                l_bottom = max(bottoms)
                                rect_total = [
                                    table.bbox[0] * scale_x,
                                    l_top * scale_y,
                                    active_mask_x * scale_x, 
                                    l_bottom * scale_y
                                ]
                                draw.rectangle(rect_total, fill="white", outline="black", width=2)
                except:
                    pass

    except Exception as e:
        # print(f"Erro silencioso: {e}")
        pass
    
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

    # Estado Inicial da Máscara (Resetado por arquivo)
    mask_state = {'mask_x': None, 'last_table_bbox': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            # Passamos o estado da máscara e recebemos o atualizado
            img, mask_state = apply_masking_v10(img, pdf_plumb.pages[i], mask_state)
        
        # Otimização visual e redimensionamento
        img = img.resize((595, 842)) 
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=80, optimize=True)
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

# --- PASSO 2: PROCESSAR ---
if uploaded_files:
    st.write("---")
    st.write("### Passo 2: Converter e Download")
    
    qtd = len(uploaded_files)
    st.caption(f"{qtd} arquivo(s) pronto(s) para conversão.")

    if st.button(f"🚀 Processar Arquivos"):
        with st.spinner('Analisando tabelas, aplicando máscaras e convertendo...'):
            try:
                processed_files = []
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    progress_bar.progress((index + 1) / qtd)

                st.success("✅ Conversão concluída com sucesso!")
                
                if len(processed_files) == 1:
                    name, data = processed_files[0]
                    st.download_button(
                        label=f"📥 Salvar {name} no Computador",
                        data=data,
                        file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, data in processed_files:
                            zf.writestr(name, data.getvalue())
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 Salvar Todos (.ZIP) no Computador",
                        data=zip_buffer,
                        file_name="Documentos_SEI_Convertidos.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")

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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v10.0 (Multi-Page Fix)</div>', unsafe_allow_html=True)
