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
st.markdown("**Nota:** O sistema manterá visível até a coluna 'Quantidade' e removerá o restante.")

uploaded_files = st.file_uploader(
    "Arraste e solte seus arquivos PDF aqui (ou clique para buscar):", 
    type="pdf", 
    accept_multiple_files=True
)

# --- FUNÇÃO AUXILIAR DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    text = text.replace('\n', ' ').replace('\r', ' ')
    # Remove pontos finais isolados para facilitar match (ex: "qtde." -> "qtde")
    return text.lower().strip().rstrip('.')

# --- FUNÇÃO DE MASCARAMENTO (v13.0 - CORTE ESTRITO NA QUANTIDADE) ---
def apply_masking_v13(image, pdf_page, mask_state):
    """
    Estratégia:
    1. Acha exclusivamente a coluna 'QUANTIDADE'.
    2. Define a borda DIREITA dessa coluna como o início do corte.
    3. Apaga tudo até o final da imagem (im_width).
    """
    
    # Busca tabelas (Linhas ou Texto)
    tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    if not tables:
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

    draw = ImageDraw.Draw(image)
    im_width, im_height = image.size
    scale_x = im_width / pdf_page.width
    scale_y = im_height / pdf_page.height

    # ÂNCORAS: Apenas variações de Quantidade
    keywords_anchor = ["qtde", "qtd", "quantidade", "quant", "quantitativo"]
    
    # STOPPERS: Palavras que indicam nova tabela (não de itens)
    keywords_stopper = ["local", "endereço", "entrega", "prazo", "responsável", "fiscal", "assinatura", "sanções", "garantia"]

    for table in tables:
        if not table.rows: continue

        found_anchor_x = None
        found_stopper = False
        
        # Analisa cabeçalho (3 primeiras linhas)
        for row_idx in range(min(3, len(table.rows))):
            row_cells = table.rows[row_idx].cells
            for cell_idx, cell in enumerate(row_cells):
                if not cell: continue
                try:
                    cropped = pdf_page.crop(cell)
                    text_raw = cropped.extract_text()
                    text_clean = clean_text(text_raw)
                    
                    # 1. Verifica ÂNCORA (Quantidade)
                    # Separa por palavras para evitar falsos positivos parciais
                    words = text_clean.split()
                    if any(k in words or k == text_clean for k in keywords_anchor):
                        found_anchor_x = cell[2] # Pega a borda DIREITA da célula
                    
                    # 2. Verifica STOPPER
                    if any(k in text_clean for k in keywords_stopper):
                        found_stopper = True
                        
                except:
                    pass
            if found_anchor_x: break

        # --- DECISÃO ---
        active_mask_x = None

        if found_anchor_x:
            # Achou Quantidade: Configura novo corte
            mask_state['mask_x'] = found_anchor_x
            mask_state['last_table_bbox'] = table.bbox
            active_mask_x = found_anchor_x
        
        elif found_stopper:
            # Achou Tabela Diferente: Reseta máscara
            mask_state['mask_x'] = None
            mask_state['last_table_bbox'] = None
            active_mask_x = None

        else:
            # Sem cabeçalho (Continuação): Usa memória com verificação de alinhamento
            if mask_state['mask_x'] is not None and mask_state['last_table_bbox'] is not None:
                prev_bbox = mask_state['last_table_bbox']
                curr_bbox = table.bbox
                
                # Se a tabela estiver alinhada horizontalmente (+- 40px)
                if abs(curr_bbox[0] - prev_bbox[0]) < 40: 
                    active_mask_x = mask_state['mask_x']
                    mask_state['last_table_bbox'] = table.bbox
                else:
                    mask_state['mask_x'] = None

        # --- DESENHO ---
        if active_mask_x is not None:
            table_rect = table.bbox
            
            # Garante que o corte não seja absurdo (ex: antes do inicio da tabela)
            if active_mask_x > table_rect[0]:
                
                # 1. MÁSCARA BRANCA (Até o fim da IMAGEM)
                rect_mask = [
                    active_mask_x * scale_x,       
                    table_rect[1] * scale_y,      
                    im_width, # Garante que apaga até a borda da folha
                    table_rect[3] * scale_y       
                ]
                draw.rectangle(rect_mask, fill="white", outline=None)

                # 2. BORDA PRETA (Linha de Fechamento)
                # Desenha linha vertical grossa na posição do corte
                draw.line(
                    [
                        (active_mask_x * scale_x, table_rect[1] * scale_y),
                        (active_mask_x * scale_x, table_rect[3] * scale_y)
                    ],
                    fill="black",
                    width=3
                )
                
                # Fecha com linha horizontal superior e inferior para acabamento
                draw.line(
                     [
                        (active_mask_x * scale_x, table_rect[1] * scale_y),
                        (active_mask_x * scale_x - 5, table_rect[1] * scale_y) # Pequeno traço para esquerda
                     ],
                     fill="black", width=2
                )
                draw.line(
                     [
                        (active_mask_x * scale_x, table_rect[3] * scale_y),
                        (active_mask_x * scale_x - 5, table_rect[3] * scale_y)
                     ],
                     fill="black", width=2
                )

                # 3. Limpeza de Rodapé (Total)
                try:
                    last_row = table.rows[-1]
                    first_cell_text = clean_text(pdf_page.crop(last_row.cells[0]).extract_text()) if last_row.cells[0] else ""
                    if "total" in first_cell_text:
                        tops = [c[1] for c in last_row.cells if c]
                        bottoms = [c[3] for c in last_row.cells if c]
                        if tops:
                             rect_total_clean = [
                                table.bbox[0] * scale_x,
                                min(tops) * scale_y,
                                active_mask_x * scale_x,
                                max(bottoms) * scale_y
                            ]
                             # Opcional: Desenhar borda no total também
                             draw.rectangle(rect_total_clean, outline="black", width=2)
                except:
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

    mask_state = {'mask_x': None, 'last_table_bbox': None}

    for i, img in enumerate(images):
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img, mask_state = apply_masking_v13(img, pdf_plumb.pages[i], mask_state)
        
        # Redimensionamento ideal para A4
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

# --- PASSO 2: PROCESSAR ---
if uploaded_files:
    st.write("---")
    st.write("### Passo 2: Converter e Download")
    
    qtd = len(uploaded_files)
    st.caption(f"{qtd} arquivo(s) pronto(s) para conversão.")

    if st.button(f"🚀 Processar Arquivos"):
        with st.spinner('Ajustando tabelas (Limite na Coluna Quantidade)...'):
            try:
                processed_files = []
                progress_bar = st.progress(0)
                
                for index, uploaded_file in enumerate(uploaded_files):
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))
                    progress_bar.progress((index + 1) / qtd)

                st.success("✅ Ajuste concluído! Tabela finalizada na Quantidade.")
                
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
st.markdown('<div class="footer">Developed by Yuri 🚀 | SEI Converter ATA - SGB v13.0 (Strict Quantity Cut)</div>', unsafe_allow_html=True)
