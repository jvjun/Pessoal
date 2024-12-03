import fitz  # PyMuPDF
import os
import re

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

def extract_dates(text):
    """
    Extrai as datas de início e final do texto usando expressões regulares.
    """
    # Expressões regulares para datas no formato DD/MM/AAAA ou DD-MM-AAAA
    date_pattern = r'\b(\d{2}[-/]\d{2}[-/]\d{4})\b'
    dates = re.findall(date_pattern, text)
    
    if len(dates) >= 2:
        start_date = dates[0]  # Supondo que a primeira data é a de início
        end_date = dates[1]    # Supondo que a segunda data é a final
        return start_date, end_date
    return None, None

def search_dates_in_pdfs(folder_path, page_number, rect):
    """
    Procura e armazena datas em um retângulo específico em uma página de PDFs em uma pasta.
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_rectangle(pdf_path, page_number, rect)
            if text:
                start_date, end_date = extract_dates(text)
                if start_date and end_date:
                    print(f"Start date: {start_date}, End date: {end_date} in rectangle on page {page_number} of '{filename}'")
                else:
                    print(f"No dates found in rectangle on page {page_number} of '{filename}'")
            else:
                print(f"No text found in rectangle on page {page_number} of '{filename}'")

def main():
    folder_path = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Antigos"
    page_number = 31  # Página que você deseja verificar
    # Define o retângulo baseado nas coordenadas fornecidas, ampliado para cobrir variações
    rect = fitz.Rect(100, 260, 340, 340)  # Ajuste as dimensões conforme necessário

    search_dates_in_pdfs(folder_path, page_number, rect)

if __name__ == "__main__":
    main()
