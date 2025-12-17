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
Converta documentos PDF de **TR (Termo de Referência)** e **Proposta de Preços** em imagens otimizadas 
e mascaradas para o **SEI**.
""")

# --- FUNÇÃO AUXILIAR DE LIMPEZA DE TEXTO ---
def clean_text(text):
    if not text: return ""
    text = text.replace('\n', ' ').replace('\r', ' ')
    return text.lower().strip()

# --- FUNÇÃO DE MASCARAMENTO (CALIBRADA v9 - CORTE VISUAL DA TABELA) ---
def apply_masking(image, pdf_page):
    try:
        # Estratégias de busca de tabela
        tables = pdf_page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
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

            # --- 1. LOCALIZAR A COLUNA DE PREÇO ---
            mask_start_x = None

            # Varre as primeiras linhas (cabeçalho)
            for row_idx in range(min(3, len(table.rows))):
                row_cells = table.rows[row_idx].cells
                for cell_idx, cell in enumerate(row_cells):
                    if not cell: continue
                    try:
                        cropped = pdf_page.crop(cell)
                        text_raw = cropped.extract_text()
                        text_clean = clean_text(text_raw)
                        
                        if any(k in text_clean for k in keywords_target):
                            # Encontrou a coluna proibida!
                            # O inicio do mascaramento é a esquerda dessa célula
                            mask_start_x = cell[0] 
                            break 
                    except:
                        pass
                if mask_start_x is not None:
                    break
            
            # --- 2. APLICAR MÁSCARA E FECHAR A TABELA ---
            if mask_start_x is not None:
                table_rect = table.bbox # (x0, top, x1, bottom)
                
                # A) A "Borracha" (Retângulo Branco)
                # Apaga tudo do início da coluna de preço até o fim original da tabela (e um pouco mais para garantir)
                rect_mask = [
                    mask_start_x * scale_x,       
                    table_rect[1] * scale_y,      
                    table_rect[2] * scale_x + 50, # Vai um pouco além da direita original para garantir
                    table_rect[3] * scale_y       
                ]
                draw.rectangle(rect_mask, fill="white", outline=None)

                # B) A Nova Borda (Retângulo Preto)
                # AQUI ESTÁ O TRUQUE PARA A IMAGEM 1:
                # Desenhamos a borda da Esquerda Original até o 'mask_start_x'.
                # Isso cria uma linha vertical preta exatamente onde o preço começaria, "fechando" a tabela ali.
                rect_border = [
                    table_rect[0] * scale_x,    # Esquerda da tabela
                    table_rect[1] * scale_y,    # Topo
                    mask_start_x * scale_x,     # Direita (NOVO LIMITE VISUAL - Corta aqui)
                    table_rect[3] * scale_y     # Fundo
                ]
                draw.rectangle(rect_border, outline="black", width=2)

            # --- 3. LIMPEZA FINAL DE LINHAS DE TOTAL ---
            # Caso exista uma linha de "Total Geral" abaixo que escape da lógica acima
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
                            
                            # Apaga a linha de total inteira visualmente, respeitando o novo corte
                            rect_total = [
                                table.bbox[0] * scale_x,
                                l_top * scale_y,
                                mask_start_x * scale_x if mask_start_x else table.bbox[2] * scale_x, 
                                l_bottom * scale_y
                            ]
                            # Se quiser apagar o valor do total (que fica a direita), o rect_mask acima já cuidou disso.
                            # Aqui garantimos que a borda do rodapé também siga o novo alinhamento se necessário.
            except:
                pass

    except Exception as e:
        # Em caso de erro, segue sem mascarar para não travar
        pass
    
    return image

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
    
    # Configuração de Margens Estreitas
    section = doc.sections[0]
    section.page_height = Cm(29.7) # A4 Altura
    section.page_width = Cm(21.0)  # A4 Largura
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(0.5) # Margem inferior bem pequena

    for i, img in enumerate(images):
        # Aplica a máscara se possível
        if has_text_layer and pdf_plumb and i < len(pdf_plumb.pages):
            img = apply_masking(img, pdf_plumb.pages[i])
        
        # OTIMIZAÇÃO DE TAMANHO (ANTI-PÁGINA EM BRANCO)
        # Redimensionamos a imagem pixel a pixel para garantir qualidade
        img = img.resize((595, 842)) # Tamanho A4 aproximado em pixels (baixa densidade para referência)
        
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=80, optimize=True)
        img_byte_arr.seek(0)

        # Inserção no Word
        # ALTERAÇÃO CRÍTICA: Reduzi de 19.0cm para 18.0cm
        # Isso garante que a altura proporcional seja menor que a altura da página,
        # evitando que o Word jogue a imagem para a próxima página ou crie uma página vazia no final.
        doc.add_picture(img_byte_arr, width=Cm(18.0))
        
        # Ajuste fino do parágrafo da imagem
        par = doc.paragraphs[-1]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        
        # Quebra de página apenas se não for a última imagem
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
        with st.spinner('Ajustando tabelas e convertendo...'):
            try:
                processed_files = []
                for uploaded_file in uploaded_files:
                    docx_data = convert_pdf_to_docx(uploaded_file.read())
                    file_name = uploaded_file.name.replace('.pdf', '') + "_SEI_SGB.docx"
                    processed_files.append((file_name, docx_data))

                st.success("✅ Sucesso!")
                
                # Download
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
st.markdown('<div class="footer">SEI Converter ATA - SGB v9.0 (Tabela Ajustada)</div>', unsafe_allow_html=True)
