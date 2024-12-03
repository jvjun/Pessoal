import fitz  # PyMuPDF

def draw_rectangle_on_pdf(pdf_path, output_pdf_path, page_number, rect):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) < page_number:
            print("O número da página é maior do que o número de páginas no documento.")
            return
        
        page = doc.load_page(page_number - 1)  # as pages are zero-indexed
        
        # Desenhar o retângulo
        shape = page.new_shape()
        shape.draw_rect(rect)
        shape.finish(color=(1, 0, 0), fill=None, width=1)  # Desenhar o retângulo em vermelho
        
        # Adicionar o retângulo à página
        shape.commit()
        
        # Salvar uma cópia do PDF com o retângulo desenhado
        doc.save(output_pdf_path)
        print(f"PDF salvo em: {output_pdf_path}")
    except Exception as e:
        print(f"Erro ao processar {pdf_path}: {e}")

# Caminho do PDF de entrada e saída
input_pdf_path = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Boven_ContratoCastrolanda_01.05.2024-31.05.2024_21.05.2024_2976.000.pdf"
output_pdf_path = r"C:\Users\Násser Saleh\OneDrive - NEOGIER ENERGIA\Área de Trabalho\Geral\Códigos\Renomeia_PDF_Contrato\Antigos\seu_pdf_com_retangulo.pdf"

# Número da página e coordenadas do retângulo
page_number = 30
rect = fitz.Rect(220, 710, 300, 720)

# Desenhar o retângulo no PDF
draw_rectangle_on_pdf(input_pdf_path, output_pdf_path, page_number, rect)
