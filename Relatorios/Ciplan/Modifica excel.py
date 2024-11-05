import os
import pandas as pd
import xlwings as xw
from datetime import datetime

# Define a pasta de downloads e o padrão do arquivo CSV
downloads_folder = os.path.expanduser("~/Downloads")
padrão_arquivo = "exportacao_joaojun01_"

# Função para localizar o arquivo CSV mais recente
def encontrar_arquivo_mais_recente():
    arquivos = [
        os.path.join(downloads_folder, f)
        for f in os.listdir(downloads_folder)
        if f.startswith(padrão_arquivo) and f.endswith(".csv")
    ]
    if arquivos:
        return max(arquivos, key=os.path.getctime)
    else:
        raise FileNotFoundError("Nenhum arquivo encontrado com o padrão 'exportacao_joaojun01_'.")

# Função para carregar, filtrar e formatar os dados
def filtrar_dados():
    # Localiza o arquivo CSV mais recente
    arquivo_mais_recente = encontrar_arquivo_mais_recente()
    print(f"Arquivo encontrado: {arquivo_mais_recente}")
    
    # Carrega o CSV e pula as primeiras três linhas, detectando o delimitador automaticamente
    df = pd.read_csv(arquivo_mais_recente, encoding='latin1', skiprows=3, sep=None, engine='python')
    
    # Filtra as linhas que atendem aos critérios de "Agente" e "Ponto / Grupo"
    df_filtrado = df[(df['Agente'] == 'CIPLAN') & (df['Ponto / Grupo'] == 'DFCIPLENTR101 (L)')]
    
    # Seleciona as colunas desejadas e formata os dados
    df_filtrado = df_filtrado[['Data', 'Hora', 'Ativa C (kWh)']].copy()
    df_filtrado['Data'] = pd.to_datetime(df_filtrado['Data'], dayfirst=True).dt.strftime("%d/%m/%Y")  # Formato dd/mm/aaaa
    
    # Substitui vírgulas por pontos para conversão correta e formata para três casas decimais
    df_filtrado['Ativa C (kWh)'] = df_filtrado['Ativa C (kWh)'].str.replace('.', '', regex=False).str.replace(',', '.').astype(float)

    # Salva o resultado filtrado em um novo arquivo CSV
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo_filtrado = os.path.join(downloads_folder, f"exportacao_filtrada_{data_atual}.csv")
    df_filtrado.to_csv(nome_arquivo_filtrado, index=False, sep=';')
    print(f"Arquivo filtrado salvo em: {nome_arquivo_filtrado}")
    
    return nome_arquivo_filtrado

# Função para abrir o Excel, limpar as células e colar os dados como valores
def inserir_dados_no_relatorio(caminho_arquivo_filtrado):
    # Carrega o arquivo CSV filtrado
    dados = pd.read_csv(caminho_arquivo_filtrado, sep=';')
    
    # Abre o Excel e a planilha de destino
    caminho_relatorio = r"C:\Users\joaoj\OneDrive - NEOGIER ENERGIA\Neogier Y\#Clientes\#Gestao\Ciplan\#Operacoes Mensais\2024\10-out\Ciplan - Relatórios - Copia.xlsm"
    app = xw.App(visible=False)
    wb = app.books.open(caminho_relatorio)
    sheet = wb.sheets["Medições"]
    
    # Limpa as colunas B, C e D a partir da linha 3
    sheet.range("B3:D1000").clear_contents()
    
    # Cola os dados a partir da linha 3
    sheet.range("B3").options(index=False, header=False).value = dados

    # Salva e fecha o Excel
    wb.save()
    wb.close()
    app.quit()
    print(f"Dados inseridos no relatório: {caminho_relatorio}")

# Executa o processo completo
caminho_excel_filtrado = filtrar_dados()
inserir_dados_no_relatorio(caminho_excel_filtrado)

