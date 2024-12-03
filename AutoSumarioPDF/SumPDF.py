from PyPDF2 import PdfReader, PdfWriter
import os
import pandas as pd

def mesclar_paginas_com_imagem(pdf_entrada, imagem_entrada, pasta_saida):
    # Lê o PDF de entrada
    leitor_pdf = PdfReader(open(pdf_entrada, "rb"))
    num_paginas = len(leitor_pdf.pages)

    # Lê o PDF da imagem (assumindo que ele tem apenas uma página)
    leitor_imagem = PdfReader(open(imagem_entrada, "rb"))
    pagina_imagem = leitor_imagem.pages[0]

    # Lê os dados do arquivo Excel
    df = pd.read_excel(r"C:\Users\joaoj\Desktop\Pessoal\AutoSumarioPDF\relacnpj.xlsx")

    # Cria um PDF separado para cada página
    for num_pagina in range(1, num_paginas):
        escritor_pdf = PdfWriter()
        pagina = leitor_pdf.pages[num_pagina]

        # Mescla a página da imagem como plano de fundo
        pagina.merge_page(pagina_imagem)
        escritor_pdf.add_page(pagina)

        # Obtém o texto da segunda linha da página
        linhas_texto = pagina.extract_text().split("\n")
        texto_segunda_linha = linhas_texto[1].strip()

        # Encontra o nome do cliente até encontrar a palavra "Mês"
        if "Mês" in texto_segunda_linha:
            nome_cliente = texto_segunda_linha.split("Mês")[0].strip()
        else:
            nome_cliente = texto_segunda_linha.strip()

        # Obtém o texto da terceira linha da página e substitui / por -
        texto_terceira_linha = linhas_texto[2].strip()
        texto_terceira_linha = texto_terceira_linha.replace("/", "-")

        # Filtra o DataFrame com base no nome do cliente
        filtered_df = df.loc[df['AGENTE'] == nome_cliente]
        if not filtered_df.empty:
            cod_cnpj = df.loc[df['AGENTE'] == nome_cliente, 'cod'].iloc[0]
        else:
            cod_cnpj = "0_"

        # Salva o nome do arquivo
        nome_arquivo_saida = f"{cod_cnpj}{nome_cliente}_Sumario_{texto_terceira_linha}.pdf"

        # Salva a página mesclada como um arquivo PDF separado
        caminho_saida = os.path.join(pasta_saida, nome_arquivo_saida)
        with open(caminho_saida, "wb") as arquivo_saida:
            escritor_pdf.write(arquivo_saida)

        print(f"Página {num_pagina + 1} mesclada com a imagem e salva como {caminho_saida}")

# Caminhos para os arquivos de entrada e saída
pdf_entrada = r"C:\Users\joaoj\Desktop\Pessoal\AutoSumarioPDF\SUM_agentetot.pdf"
imagem_entrada = r"C:\Users\joaoj\Desktop\Pessoal\AutoSumarioPDF\Fundo_Neogier.pdf"
pasta_saida = r"C:\Users\joaoj\Desktop\Pessoal\AutoSumarioPDF\Relatórios_Contabilização_Clientes"

# Cria a pasta de saída se não existir
if not os.path.exists(pasta_saida):
    os.makedirs(pasta_saida)

# Chame a função para mesclar as páginas com a imagem
mesclar_paginas_com_imagem(pdf_entrada, imagem_entrada, pasta_saida)

# Exibe uma mensagem de sucesso
print(f"Páginas mescladas com a imagem e salvas em {pasta_saida}")
