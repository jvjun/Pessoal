import os
import fitz  # PyMuPDF
from PIL import Image, ImageDraw

def extract_text_from_rectangle(pdf_path, page_number, rect):
    doc = None
    text = None
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            return None
        page = doc.load_page(page_number - 1)  # as pages are zero-indexed
        text = page.get_text("text", clip=rect)
        text = text.strip()
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    finally:
        if doc:
            doc.close()
    return text

def save_page_with_rectangle(pdf_path, page_number, rect, output_path):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            return None
        page = doc.load_page(page_number - 1)  # as pages are zero-indexed
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        draw = ImageDraw.Draw(img)
        draw.rectangle([(rect.x0, rect.y0), (rect.x1, rect.y1)], outline="red", width=5)
        
        img.save(output_path)
        print(f"Saved page with rectangle as PNG: {output_path}")
    except Exception as e:
        print(f"Error saving page with rectangle as PNG: {e}")
    finally:
        if doc:
            doc.close()

def process_specific_pdf(pdf_path, output_folder, rect_cliente, rect_contratado, rect_energia, rect_datas):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    save_page_with_rectangle(pdf_path, 27, rect_cliente, os.path.join(output_folder, "cliente.png"))
    save_page_with_rectangle(pdf_path, 27, rect_contratado, os.path.join(output_folder, "contratado.png"))
    save_page_with_rectangle(pdf_path, 30, rect_energia, os.path.join(output_folder, "energia.png"))
    save_page_with_rectangle(pdf_path, 31, rect_datas, os.path.join(output_folder, "datas.png"))

# Caminho do PDF específico
pdf_path = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Antigos\teste\Boven Comercializadora_Castrolanda - Comer_ContratoCastrolanda_DataInicio-DataFinal_assi15.01.2020_montanteEnergia.pdf"

# Defina os retângulos para extrair o texto do cliente, do contratado, da energia e das datas
rect_cliente = fitz.Rect(50, 175, 185, 190)  # Ajuste conforme necessário
rect_contratado = fitz.Rect(280, 175, 380, 190)  # Ajuste conforme necessário
rect_energia = fitz.Rect(220, 710, 300, 720)
rect_datas = fitz.Rect(110, 280, 300, 320)

# Pasta de saída
output_folder = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Antigos\teste\Localização pdf"

# Processar o PDF específico
process_specific_pdf(pdf_path, output_folder, rect_cliente, rect_contratado, rect_energia, rect_datas)
