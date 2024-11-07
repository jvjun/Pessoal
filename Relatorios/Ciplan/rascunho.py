import os
import pandas as pd

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
    
    # Carrega o CSV e pula as primeiras três linhas
    df = pd.read_csv(arquivo_mais_recente, encoding='latin1', skiprows=3, sep=None, engine='python')
    
    # Filtra as linhas que atendem aos critérios de "Agente" e "Ponto / Grupo"
    df_filtrado = df[(df['Agente'] == 'CIPLAN') & (df['Ponto / Grupo'] == 'DFCIPLENTR101 (L)')]
    
    # Seleciona e formata as colunas desejadas
    df_filtrado['Data'] = pd.to_datetime(df_filtrado['Data'], dayfirst=True).dt.strftime('%d/%m/%Y')
    df_filtrado['Ativa C (kWh)'] = df_filtrado['Ativa C (kWh)'].replace({',': '', '.': ','}, regex=True)
    
    # Retorna os dados filtrados com as colunas específicas
    return df_filtrado[['Data', 'Hora', 'Ativa C (kWh)']]

# Função para salvar os dados filtrados em um novo arquivo Excel
def salvar_dados_em_excel(dados):
    caminho_excel = r"C:\Users\joaoj\Downloads\teste.xlsx"
    dados.to_excel(caminho_excel, index=False, sheet_name="Dados Filtrados")
    print(f"Dados salvos em {caminho_excel} com sucesso.")

# Executa o processo completo
if __name__ == "__main__":
    dados_filtrados = filtrar_dados()
    salvar_dados_em_excel(dados_filtrados)
