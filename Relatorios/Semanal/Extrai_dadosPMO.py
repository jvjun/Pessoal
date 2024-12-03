import pdfplumber

def extract_table_by_title_clean(pdf_path, table_title):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if table_title in text:
                print(f"Título encontrado na página {page.page_number}")
                # Extrair tabelas da página
                tables = page.extract_tables()
                if tables:
                    print("Tabela encontrada! Processando...")
                    # Assumindo que a tabela desejada é a primeira tabela extraída
                    raw_table = tables[0]
                    # Limpar a tabela removendo valores None e células vazias
                    clean_table = []
                    for row in raw_table:
                        clean_row = [cell if cell is not None else "" for cell in row]
                        clean_row = [cell.strip() for cell in clean_row if cell.strip()]  # Remover células vazias
                        if clean_row:
                            clean_table.append(clean_row)
                    return clean_table
                else:
                    print("Nenhuma tabela encontrada nesta página.")
                    return None
        print("Título não encontrado no PDF.")
        return None

# Parâmetros do código
pdf_path = r"C:\Users\joaoj\OneDrive - NEOGIER ENERGIA\Neogier Y\Estudos\Semanais\2024\2024 - 11 - Novembro\Informe Semanal - Semana 04\RELATORIO-PMO-16_11 a 22_11.pdf"  # Substitua pelo caminho do PDF
table_title = "Tabela 2 – Previsão de ENAs da Revisão 3 de Novembro/2024"

# Extrair a tabela e processar
extracted_table = extract_table_by_title_clean(pdf_path, table_title)

# Exibir a tabela processada
if extracted_table:
    print("\nTabela extraída e processada com sucesso:")
    for row in extracted_table:
        print(row)
else:
    print("Nenhuma tabela foi extraída.")
