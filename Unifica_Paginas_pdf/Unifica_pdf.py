import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch

def merge_pdf_pages_to_one(input_pdf_path, output_pdf_path):
    # Abrir o arquivo PDF de entrada
    with open(input_pdf_path, "rb") as input_pdf_file:
        # Carregar o PDF
        reader = PyPDF2.PdfReader(input_pdf_file)
        num_pages = len(reader.pages)

        # Configurar o PDF de saída com uma única página
        output_canvas = canvas.Canvas(output_pdf_path, pagesize=A4)

        # Tamanho da página A4
        width, height = A4

        # Definir quantas miniaturas de páginas você quer em cada linha e coluna
        rows = 3  # número de linhas
        cols = 2  # número de colunas
        thumbnail_width = width / cols
        thumbnail_height = height / rows

        # Iterar sobre as páginas e colocá-las na nova página
        for i, page in enumerate(reader.pages):
            # Renderizar a página como uma string de PDF
            page_data = page.extract_text()

            # Definir a posição para cada página miniaturizada
            x_pos = (i % cols) * thumbnail_width
            y_pos = height - ((i // cols) + 1) * thumbnail_height

            # Adicionar texto básico no canvas (apenas a título de exemplo, já que desenhar PDFs completos requer um tratamento avançado)
            output_canvas.drawString(x_pos + 10, y_pos + thumbnail_height - 10, f"Page {i+1}")
            output_canvas.drawString(x_pos + 10, y_pos + thumbnail_height - 30, page_data[:100])  # Exemplo de conteúdo

            # Se todas as miniaturas estiverem desenhadas, criar uma nova página
            if (i + 1) % (rows * cols) == 0 and i < num_pages - 1:
                output_canvas.showPage()

        # Salvar o arquivo PDF de saída
        output_canvas.save()

# Exemplo de uso
input_pdf_path = r"C:\Users\joaoj\Desktop\Grade BCD.pdf"  # Caminho para o PDF de entrada
output_pdf_path = r"C:\Users\joaoj\Desktop\merged_output.pdf"  # Caminho para o PDF de saída
merge_pdf_pages_to_one(input_pdf_path, output_pdf_path)
