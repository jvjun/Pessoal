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

def get_date_before_keyword(text, keyword):
    if text and keyword in text:
        keyword_index = text.find(keyword)
        if keyword_index != -1:
            before_keyword = text[:keyword_index].strip()
            date_match = re.search(r'\d{2}/\d{2}/\d{4}', before_keyword[::-1])
            if date_match:
                return date_match.group(0)[::-1]  # Reverse it back to normal
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

def extract_info_and_rename(pdf_path, folder_path):
    text = extract_text_from_page(pdf_path, [32, 34, 35, 36])
    
    cliente = get_text_after_keyword(text.get(32, ""), "Parte Compradora")
    data_assinatura = get_date_after_keyword(text.get(32, ""), "Assinado em ")
    energia_total = clean_energy(text.get(34, ""))  # Directly using clean_energy function
    inicio_data, final_data = extract_dates(text.get(35, "") + text.get(36, ""))
    
    # Sanitize and format extracted data
    sanitized_cliente = sanitize_filename(cliente) if cliente else ""
    sanitized_frase = "ContratoCastrolanda"  # Assuming this is fixed
    sanitized_inicio_data = sanitize_filename(inicio_data) if inicio_data else ""
    sanitized_final_data = sanitize_filename(final_data) if final_data else ""
    sanitized_data_assinatura = sanitize_filename(clean_date(data_assinatura)) if data_assinatura else ""
    sanitized_energia_total = sanitize_filename(energia_total) if energia_total else ""
    
    # Construct new filename
    new_filename = f"{sanitized_cliente}_{sanitized_frase}_{sanitized_inicio_data}-{sanitized_final_data}_{sanitized_data_assinatura}_{sanitized_energia_total}.pdf"
    new_path = os.path.join(folder_path, new_filename)
    
    print(f"Attempting to rename {pdf_path} to {new_path}")
    
    try:
        os.rename(pdf_path, new_path)
        print(f"Renamed '{pdf_path}' to '{new_filename}'")
    except Exception as e:
        print(f"Error renaming {pdf_path} to '{new_filename}': {e}")

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
        return text
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None

def search_word_in_pdfs(folder_path, page_number, rect):
    """
    Procura palavras em um retângulo específico em uma página de PDFs em uma pasta.
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_rectangle(pdf_path, page_number, rect)
            if text:
                words = text.split()
                print(f"Words in rectangle on page {page_number} of '{filename}': {words}")
            else:
                print(f"No text found in rectangle on page {page_number} of '{filename}'")

def rename_all_pdfs_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}")
            extract_info_and_rename(pdf_path, folder_path)

# Define folder path containing PDFs
folder_path = r"C:\Users\joaoj\Desktop\Pessoal\thunders\Renomeia_PDF_Contrato"

page_number = 32  # Página que você deseja verificar
rect = fitz.Rect(50, 460, 200, 470) 

print(rect)
search_word_in_pdfs(folder_path, page_number, rect)

# Rename all PDFs in the folder
rename_all_pdfs_in_folder(folder_path)
