import os
import re
from PyPDF2 import PdfReader
import fitz  # PyMuPDF

def extract_text_from_page(pdf_path, page_numbers):
    text = {}
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            for page_number in page_numbers:
                if len(reader.pages) >= page_number:
                    page = reader.pages[page_number - 1]  # pages are zero-indexed
                    text[page_number] = page.extract_text()
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return text

def get_text_after_keyword(text, keyword):
    if text and keyword in text:
        keyword_index = text.find(keyword)
        if keyword_index != -1:
            after_keyword = text[keyword_index + len(keyword):].strip()
            return after_keyword.split()[0]  # Stop at the first space
    return None

def get_date_after_keyword(text, keyword):
    if text and keyword in text:
        keyword_index = text.find(keyword)
        if keyword_index != -1:
            after_keyword = text[keyword_index + len(keyword):].strip()
            match = re.match(r'^[0-9/]+', after_keyword)
            if match:
                return match.group(0)
    return None

def extract_price(text):
    """
    Extrai o preço do texto na frente da frase 'Condições de Pagamento e Preço '.
    """
    keyword = "Condições de Pagamento e Preço "
    if text and keyword in text:
        keyword_index = text.find(keyword)
        if keyword_index != -1:
            after_keyword = text[keyword_index + len(keyword):].strip()
            # Extrair o valor na frente
            match = re.search(r'R\$[\s]*([\d.,]+)', after_keyword)
            if match:
                return match.group(1).replace('.', '').replace(',', '.')  # Formatar valor
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

def clean_date(text):
    # Convert date format from dd/mm/yyyy to dd.mm.yyyy
    return text.replace('/', '.')

def clean_energy(text):
    # Extract energy value from "Energia Total Contratada: 3.600,000[MWh]"
    match = re.search(r'Energia Total Contratada:\s*([\d,.]+)\[', text)
    if match:
        return match.group(1).replace('.', '').replace(',', '.')  # Replace comma with dot for decimal
    return None

def sanitize_filename(filename):
    # Remove characters that are not allowed in filenames and replace "/" with "-"
    return "".join(c if c.isalnum() or c in (' ', '.', '_') else '-' for c in filename).rstrip()

def extract_info_and_rename(pdf_path, folder_path):
    text = extract_text_from_page(pdf_path, [32, 34, 35, 36])
    
    cliente = get_text_after_keyword(text.get(32, ""), "Parte Compradora")
    data_assinatura = get_date_after_keyword(text.get(32, ""), "Assinado em ")
    energia_total = clean_energy(text.get(34, ""))  # Directly using clean_energy function
    inicio_data, final_data = extract_dates(text.get(35, "") + text.get(36, ""))
    preco = extract_price(text.get(34, "") + text.get(36, ""))  # Buscando o preço nas páginas
    
    # Sanitize and format extracted data
    sanitized_cliente = sanitize_filename(cliente) if cliente else ""
    sanitized_frase = "ContratoCastrolanda"  # Assuming this is fixed
    sanitized_inicio_data = sanitize_filename(inicio_data) if inicio_data else ""
    sanitized_final_data = sanitize_filename(final_data) if final_data else ""
    sanitized_data_assinatura = sanitize_filename(clean_date(data_assinatura)) if data_assinatura else ""
    sanitized_energia_total = sanitize_filename(energia_total) if energia_total else ""
    sanitized_preco = sanitize_filename(preco) if preco else "SemPreco"
    
    # Construct new filename
    new_filename = f"{sanitized_cliente}_{sanitized_frase}_{sanitized_inicio_data}-{sanitized_final_data}_{sanitized_data_assinatura}_{sanitized_energia_total}_{sanitized_preco}.pdf"
    new_path = os.path.join(folder_path, new_filename)
    
    print(f"Attempting to rename {pdf_path} to {new_path}")
    
    try:
        os.rename(pdf_path, new_path)
        print(f"Renamed '{pdf_path}' to '{new_filename}'")
    except Exception as e:
        print(f"Error renaming {pdf_path} to '{new_filename}': {e}")

def rename_all_pdfs_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}")
            extract_info_and_rename(pdf_path, folder_path)

# Define folder path containing PDFs
folder_path = r"C:\Users\joaoj\Desktop\Pessoal\thunders\Renomeia_PDF_Contrato"

# Rename all PDFs in the folder
rename_all_pdfs_in_folder(folder_path)
