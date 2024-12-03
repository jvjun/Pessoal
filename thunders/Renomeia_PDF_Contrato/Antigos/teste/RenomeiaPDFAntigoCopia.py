import os
import re
import fitz  # PyMuPDF
from PIL import Image, ImageDraw

def extract_text_from_rectangle(pdf_path, page_number, rect):
    """
    Extrai texto de um retângulo específico em uma página de um PDF.
    """
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            return None
        page = doc.load_page(page_number - 1)  # as pages are zero-indexed
        text = page.get_text("text", clip=rect)
        text = text.strip()
        return text
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None
    finally:
        doc.close()

def get_text_after_keyword(text, keyword):
    if text and keyword in text:
        keyword_index = text.find(keyword)
        if keyword_index != -1:
            after_keyword = text[keyword_index + len(keyword):].strip()
            after_keyword = after_keyword.split()[0]  # Stop at the first space
            return after_keyword
    return None

def extract_date_after_assinado_em(text):
    """
    Extrai a data após a frase "Assinado em" e retorna a data no formato desejado.
    """
    match = re.search(r'Assinado em\s*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    if match:
        date_str = match.group(1)
        return date_str.replace('/', '.')
    return None

def clean_date(text):
    # Convert date format from dd/mm/yyyy to dd.mm.yyyy
    return text.replace('/', '.')

def clean_energy(text):
    # Extract energy value from "Energia Total Contratada: 3.600,000[MWh]"
    match = re.search(r'Energia Total Contratada:\s*([\d,.]+)\[', text)
    if match:
        return match.group(1).replace('.', '').replace(',', '.')  # Replace comma with dot for decimal
    return None

def extract_dates(text):
    # Extract start and end dates from text
    start_date = re.search(r'Início:\s*([\d/]+)', text)
    end_date = re.search(r'Final:\s*([\d/]+)', text)
    
    if start_date:
        start_date = clean_date(start_date.group(1))
    if end_date:
        end_date = clean_date(end_date.group(1))
    
    return start_date, end_date

def sanitize_filename(filename):
    # Remove characters that are not allowed in filenames and replace "/" with "-"
    return "".join(c if c.isalnum() or c in (' ', '.', '_') else '-' for c in filename).rstrip()

def truncate_filename(filename, max_length=255):
    # Truncate the filename if it exceeds the maximum length
    if len(filename) <= max_length:
        return filename
    base, ext = os.path.splitext(filename)
    return base[:max_length - len(ext)] + ext

def extract_info_from_pdf(pdf_path):
    # Define os retângulos para extrair energia e datas
    rect_energia = fitz.Rect(220, 710, 300, 720)
    rect_datas = fitz.Rect(110, 280, 300, 320)

    # Extrair energia e datas
    energia_text = extract_text_from_rectangle(pdf_path, 30, rect_energia)
    datas_text = extract_text_from_rectangle(pdf_path, 31, rect_datas)
    
    print(f"Texto da Energia: {energia_text}")  # Adicionado para depuração
    print(f"Texto das Datas: {datas_text}")  # Adicionado para depuração
    
    # Extrair o valor de energia
    energia_match = re.search(r'[\d,.]+', energia_text) if energia_text else None
    energia_total = energia_match.group(0).replace('.', '').replace(',', '.') if energia_match else "EnergiaDesconhecida"
    
    # Extrair datas de início e final
    datas = datas_text.split('\n') if datas_text else []
    inicio_data = datas[0].strip().replace('/', '.') if len(datas) > 0 else "DataInicioDesconhecida"
    final_data = datas[1].strip().replace('/', '.') if len(datas) > 1 else "DataFinalDesconhecida"
    
    return energia_total, inicio_data, final_data

def save_page_with_rectangle(pdf_path, page_number, rect, output_path):
    """
    Salva uma página inteira do PDF como uma imagem PNG e desenha um retângulo vermelho ao redor da área especificada.
    """
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            raise ValueError(f"Page {page_number} does not exist in {pdf_path}")
        page = doc.load_page(page_number - 1)  # as pages are zero-indexed
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Draw rectangle
        draw = ImageDraw.Draw(img)
        draw.rectangle([(rect.x0, rect.y0), (rect.x1, rect.y1)], outline="red", width=5)
        
        # Save image
        img.save(output_path)
        print(f"Saved page with rectangle as PNG: {output_path}")
    except Exception as e:
        print(f"Error saving page with rectangle as PNG: {e}")
    finally:
        doc.close()

def extract_info_and_rename(pdf_path, folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas):
    # Extrair cliente e contratado dos retângulos
    cliente = extract_text_from_rectangle(pdf_path, 27, rect_cliente)
    contratado = extract_text_from_rectangle(pdf_path, 27, rect_contratado)
    
    if not cliente or not contratado:
        print(f"Failed to extract client or contracted info from {pdf_path}")
    else:
        print(f"Cliente: {cliente}")
        print(f"Contratado: {contratado}")
    
    # Extrair texto da página 27 para obter a data de assinatura
    text_page_27 = extract_text_from_rectangle(pdf_path, 27, fitz.Rect(0, 0, 595, 842))  # Página inteira
    data_assinatura = extract_date_after_assinado_em(text_page_27)
    
    # Extrair energia e datas
    energia_total, inicio_data, final_data = extract_info_from_pdf(pdf_path)

    # Sanitize and format extracted data
    sanitized_cliente = sanitize_filename(cliente) if cliente else ""
    sanitized_contratado = sanitize_filename(contratado) if contratado else ""
    sanitized_frase = "ContratoCastrolanda"
    sanitized_inicio_data = sanitize_filename(inicio_data) if inicio_data else ""
    sanitized_final_data = sanitize_filename(final_data) if final_data else ""
    sanitized_data_assinatura = sanitize_filename(data_assinatura) if data_assinatura else ""
    sanitized_energia_total = sanitize_filename(energia_total) if energia_total else ""
    
    # Construct new filename
    new_filename = f"{sanitized_cliente}_{sanitized_contratado}_{sanitized_frase}_{sanitized_inicio_data}-{sanitized_final_data}_assi{sanitized_data_assinatura}_montante{sanitized_energia_total}.pdf"
    new_filename = truncate_filename(new_filename)
    new_path = os.path.join(folder_path, new_filename)
    
    print(f"Attempting to rename {pdf_path} to {new_path}")
    
    try:
        os.rename(pdf_path, new_path)
        print(f"Renamed '{pdf_path}' to '{new_filename}'")
    except Exception as e:
        print(f"Error renaming {pdf_path} to '{new_filename}': {e}")

    # Criar pasta para armazenar as imagens PNG
    output_folder = os.path.join(folder_path, f"{sanitized_cliente}_{sanitized_contratado}_{sanitized_energia_total}")
    os.makedirs(output_folder, exist_ok=True)
    
    # Excluir arquivos PNG existentes na pasta
    for png_file in os.listdir(output_folder):
        if png_file.endswith(".png"):
            os.remove(os.path.join(output_folder, png_file))
    
    # Save the pages with rectangles as PNGs
    save_page_with_rectangle(pdf_path, 27, rect_cliente, os.path.join(output_folder, "cliente.png"))
    save_page_with_rectangle(pdf_path, 27, rect_contratado, os.path.join(output_folder, "contratado.png"))
    save_page_with_rectangle(pdf_path, 30, rect_energia, os.path.join(output_folder, "energia.png"))
    save_page_with_rectangle(pdf_path, 31, rect_datas, os.path.join(output_folder, "datas.png"))

def rename_all_pdfs_in_folder(folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}")
            extract_info_and_rename(pdf_path, folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas)

# Define folder path containing PDFs
folder_path = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Antigos\teste"

# Defina os retângulos para extrair o texto do cliente, do contratado, da energia e das datas
rect_cliente = fitz.Rect(50, 175, 185, 190)  # Ajuste conforme necessário
rect_contratado = fitz.Rect(280, 175, 380, 190)  # Ajuste conforme necessário
rect_energia = fitz.Rect(220, 710, 300, 720)
rect_datas = fitz.Rect(110, 280, 300, 320)

# Rename all PDFs in the folder
rename_all_pdfs_in_folder(folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas)
