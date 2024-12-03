import os
import re
import fitz  # PyMuPDF

def extract_text_from_rectangle(pdf_path, page_number, rect):
    """
    Extrai o texto de uma área específica de uma página de um PDF.

    :param pdf_path: Caminho do arquivo PDF.
    :param page_number: Número da página (1-based).
    :param rect: Retângulo definindo a área a ser extraída.
    :return: Texto extraído ou None se houver erro.
    """
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            return None
        page = doc.load_page(page_number - 1)  # Páginas são indexadas a partir de 0
        text = page.get_text("text", clip=rect)
        return text.strip() if text else None
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    finally:
        doc.close()
    return None

def extract_date_after_assinado_em(text):
    """
    Extrai a data após o texto 'Assinado em' de um texto.

    :param text: Texto contendo a data.
    :return: Data no formato dd.mm.aaaa ou None se não encontrada.
    """
    if text:
        match = re.search(r'Assinado em\s*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
        if match:
            date_str = match.group(1)
            return date_str.replace('/', '.')
    return None

def clean_date(text):
    """
    Converte datas no formato dd/mm/aaaa para dd.mm.aaaa.

    :param text: Texto da data.
    :return: Data formatada.
    """
    return text.replace('/', '.')

def clean_energy(text):
    """
    Limpa e formata o texto da energia extraída.

    :param text: Texto contendo a energia.
    :return: Energia formatada ou None se não encontrada.
    """
    if text:
        match = re.search(r'(\d+,\d+)', text)
        if match:
            return match.group(1).replace('.', '').replace(',', '.')  # Substitui vírgula por ponto para decimal
    return None

def clean_price(text):
    """
    Extrai e formata o preço do texto, removendo "R$" e substituindo vírgula por ponto.

    :param text: Texto contendo o preço.
    :return: Preço formatado ou None se não encontrado.
    """
    if text:
        match = re.search(r'R\$\s*(\d+,\d+)', text)
        if match:
            return match.group(1).replace(',', '.')  # Remove "R$" e substitui vírgula por ponto
    return None

def extract_dates(text):
    """
    Extrai e organiza as datas de início e fim de um texto.

    :param text: Texto contendo as datas.
    :return: Tupla (data de início, data final) ou (None, None) se não encontradas.
    """
    start_date, end_date = None, None
    if text:
        dates = re.findall(r'(\d{2}/\d{2}/\d{4})', text)
        if dates:
            sorted_dates = sorted([clean_date(date) for date in dates], key=lambda date: list(map(int, date.split('.'))))
            start_date = sorted_dates[0]
            if len(sorted_dates) > 1:
                end_date = sorted_dates[1]
    return start_date, end_date

def sanitize_filename(filename):
    """
    Remove caracteres inválidos do nome do arquivo.

    :param filename: Nome do arquivo.
    :return: Nome do arquivo sanitizado.
    """
    return "".join(c if c.isalnum() or c in (' ', '.', '_') else '-' for c in filename).rstrip()

def truncate_filename(filename, max_length=255):
    """
    Trunca o nome do arquivo se exceder o comprimento máximo.

    :param filename: Nome do arquivo.
    :param max_length: Comprimento máximo permitido.
    :return: Nome do arquivo truncado.
    """
    if len(filename) <= max_length:
        return filename
    base, ext = os.path.splitext(filename)
    return base[:max_length - len(ext)] + ext

def extract_info_from_pdf(pdf_path, rect_energia, rect_datas, rect_preco):
    """
    Extrai informações de energia, datas e preço de um PDF.

    :param pdf_path: Caminho do arquivo PDF.
    :param rect_energia: Retângulo para extrair energia.
    :param rect_datas: Retângulo para extrair datas.
    :param rect_preco: Retângulo para extrair preço.
    :return: Tupla contendo energia, data de início, data final e preço.
    """
    energia_text = extract_text_from_rectangle(pdf_path, 30, rect_energia)
    datas_text = extract_text_from_rectangle(pdf_path, 31, rect_datas)
    preco_text = extract_text_from_rectangle(pdf_path, 31, rect_preco)
    
    print(f"Texto da Energia: {energia_text}")  # Adicionado para depuração
    print(f"Texto das Datas: {datas_text}")  # Adicionado para depuração
    print(f"Texto do Preço: {preco_text}")  # Adicionado para depuração
    
    energia_total = clean_energy(energia_text) or "EnergiaNONE"
    inicio_data, final_data = extract_dates(datas_text)
    inicio_data = inicio_data or "DINONE"
    final_data = final_data or "DFNONE"
    preco = clean_price(preco_text) or "PrecoNeg"
    
    return energia_total, inicio_data, final_data, preco

def extract_info_and_rename(pdf_path, folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas, rect_preco):
    """
    Extrai informações de um PDF e renomeia o arquivo com base nessas informações.

    :param pdf_path: Caminho do arquivo PDF.
    :param folder_path: Caminho da pasta para salvar o arquivo renomeado.
    :param rect_cliente: Retângulo para extrair o nome do cliente.
    :param rect_contratado: Retângulo para extrair o nome do contratado.
    :param rect_energia: Retângulo para extrair a energia.
    :param rect_datas: Retângulo para extrair as datas.
    :param rect_preco: Retângulo para extrair o preço.
    """
    cliente = extract_text_from_rectangle(pdf_path, 27, rect_cliente)
    contratado = extract_text_from_rectangle(pdf_path, 27, rect_contratado)
    
    if not cliente or not contratado:
        print(f"Failed to extract client or contracted info from {pdf_path}")
        return
    
    print(f"Cliente: {cliente}")
    print(f"Contratado: {contratado}")
    
    text_page_27 = extract_text_from_rectangle(pdf_path, 27, fitz.Rect(0, 0, 595, 842))
    data_assinatura = extract_date_after_assinado_em(text_page_27)
    
    energia_total, inicio_data, final_data, preco = extract_info_from_pdf(pdf_path, rect_energia, rect_datas, rect_preco)

    sanitized_cliente = sanitize_filename(cliente) if cliente else ""
    sanitized_contratado = sanitize_filename(contratado) if contratado else ""
    sanitized_inicio_data = sanitize_filename(inicio_data) if inicio_data else ""
    sanitized_final_data = sanitize_filename(final_data) if final_data else ""
    sanitized_data_assinatura = sanitize_filename(data_assinatura) if data_assinatura else ""
    sanitized_energia_total = sanitize_filename(energia_total) if energia_total else ""
    sanitized_preco = sanitize_filename(preco) if preco else ""
    
    new_filename = f"VEN-{sanitized_cliente}_COMP-{sanitized_contratado}_Ini-{sanitized_inicio_data}-Fim {sanitized_final_data}_assi{sanitized_data_assinatura}_{sanitized_energia_total}_RS{sanitized_preco}.pdf"
    new_filename = truncate_filename(new_filename)
    new_path = os.path.join(folder_path, new_filename)
    
    print(f"Attempting to rename {pdf_path} to {new_path}")
    
    try:
        os.rename(pdf_path, new_path)
        print(f"Renamed '{pdf_path}' to '{new_filename}'")
    except Exception as e:
        print(f"Error renaming {pdf_path} to '{new_filename}': {e}")

def rename_all_pdfs_in_folder(folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas, rect_preco):
    """
    Renomeia todos os arquivos PDF em uma pasta com base nas informações extraídas.

    :param folder_path: Caminho da pasta contendo os arquivos PDF.
    :param rect_cliente: Retângulo para extrair o nome do cliente.
    :param rect_contratado: Retângulo para extrair o nome do contratado.
    :param rect_energia: Retângulo para extrair a energia.
    :param rect_datas: Retângulo para extrair as datas.
    :param rect_preco: Retângulo para extrair o preço.
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}")
            extract_info_and_rename(pdf_path, folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas, rect_preco)

# Defina o caminho da pasta contendo os PDFs
folder_path = r"C:\Users\joaoj\Desktop\Pessoal\thunders\Renomeia_PDF_Contrato\Antigos"

# Defina os retângulos para extrair o texto do cliente, do contratado, da energia, das datas e do preço
rect_cliente = fitz.Rect(50, 175, 185, 190)  # Ajuste conforme necessário
rect_contratado = fitz.Rect(280, 175, 380, 190)  # Ajuste conforme necessário
rect_energia = fitz.Rect(220, 710, 300, 720)  # Coordenadas da energia
rect_datas = fitz.Rect(110, 280, 300, 320)  # Coordenadas das datas
rect_preco = fitz.Rect(75, 350, 160, 370)  # Coordenadas do preço

# Renomear todos os PDFs na pasta
rename_all_pdfs_in_folder(folder_path, rect_cliente, rect_contratado, rect_energia, rect_datas, rect_preco)
