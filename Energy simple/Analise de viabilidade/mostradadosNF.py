from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time
import pyperclip  # Para manipular a área de transferência

# Função para ler o login e senha do arquivo
def ler_credenciais(caminho_arquivo):
    with open(caminho_arquivo, 'r') as file:
        linha = file.readline().strip()
        login, senha = linha.split(';')
    return login, senha

# Função para ler os dados do cliente e valores de consumo da planilha
def ler_dados_cliente(caminho_planilha):
    df = pd.read_excel(caminho_planilha, usecols="A:I")  # Ajustando para incluir 9 colunas
    return df

# Caminho dos arquivos com as credenciais e dados
caminho_credenciais = r"C:\Users\joaoj\Desktop\Pessoal\Energy simple\loginES.txt"
caminho_planilha = r"C:\Users\joaoj\Desktop\Pessoal\Energy simple\Analise de viabilidade\teste.xlsx"
login, senha = ler_credenciais(caminho_credenciais)
df = ler_dados_cliente(caminho_planilha)

# Configuração do Edge com modo InPrivate
options = Options()
options.add_argument("inprivate")
driver = webdriver.Edge(options=options)

# URL da página de login
login_url = "https://energysimple.com.br/simple/sistema_migracao/gerenciar_tabela.php?params=O7fWYTKBopwAqPkpGJQ1R3gjVySVuEcDj1jhAxZo3H2FrSAMZ1R8mn66QexETSzTkM2hhHbZk43XVvG_wiNJcg,,#"
driver.get(login_url)

# Realiza login na página
try:
    username_field = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[2]/form/div[2]/span/input[1]"))
    )
    username_field.send_keys(login)

    password_field = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[2]/form/div[3]/span/input[1]"))
    )
    password_field.send_keys(senha)

    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/form/div[4]/a"))
    )
    login_button.click()
except Exception as e:
    print(f"Erro no login: {e}")

# Redireciona para o URL de login novamente
time.sleep(3)
driver.get(login_url)

# Aguarda e clica no botão para incluir ano
try:
    incluir_ano_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[3]/a[3]"))
    )
    incluir_ano_button.click()
    print("Botão de 'incluir ano' clicado com sucesso.")
except Exception as e:
    print(f"Erro ao clicar no botão de 'incluir ano': {e}")

# Aguarda alguns segundos e insere o nome do cliente no campo específico
time.sleep(5)
try:
    nome_cliente_field = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[14]/div[2]/div/form/table/tbody/tr[2]/td[3]/span/input[1]"))
    )
    nome_cliente_field.send_keys(df.iloc[0, 0])  # Insere o nome do cliente da primeira linha
    time.sleep(0.5)
    nome_cliente_field.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.5)
    nome_cliente_field.send_keys(Keys.ENTER)
    time.sleep(0.5)

    cnpj_field = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[14]/div[2]/div/form/table/tbody/tr[4]/td[3]/span/input[1]"))
    )
    cnpj_field.send_keys(str(df.iloc[0, 1]))  # Insere o CNPJ do cliente da primeira linha
    time.sleep(0.5)
    cnpj_field.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.5)
    cnpj_field.send_keys(Keys.ENTER)
    time.sleep(0.5)

except Exception as e:
    print(f"Erro ao inserir os dados do cliente: {e}")

# Função para inserir dados na linha correspondente
def inserir_dados_na_linha(linha_dados, linha_xpath):
    valores = {
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[1]/div": linha_dados[2],  # Ano
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[2]/div": linha_dados[3],  # Mês
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[3]/div": linha_dados[4],  # Consumo Ponta
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[4]/div": linha_dados[5],  # Consumo Fora Ponta
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[5]/div": linha_dados[6],  # Reservado
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[6]/div": linha_dados[7],  # Demanda Ponta
        f"/html/body/div[14]/div[2]/div/form/div/div[2]/div/div[2]/div[2]/table/tbody/tr[{linha_xpath}]/td[7]/div": linha_dados[8],  # Demanda Fora Ponta
    }

    # Insere cada valor no campo correspondente
    for xpath, value in valores.items():
        if pd.isna(value):  # Pula campos vazios
            continue
        
        cell = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )

        # Clique único para ativar o campo de edição
        cell.click()
        # Limpa o campo antes de inserir o valor
        action_chains = ActionChains(driver)
        action_chains.double_click(cell).send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)
        
        # Condição especial para o campo "Mês"
        if xpath.endswith(f"/td[2]/div"):
            pyperclip.copy(str(value))
            action_chains.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            time.sleep(0.5)
            action_chains.send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.5)
            action_chains.send_keys(Keys.ENTER).perform()
        else:
            pyperclip.copy(str(value))
            action_chains.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

        print(f"Valor '{value}' inserido com sucesso no campo com XPath '{xpath}'.")
        time.sleep(0.2)

# Loop para inserir os dados linha por linha
for index, row in df.iterrows():
    inserir_dados_na_linha(row, index + 1)  # Adiciona 1 porque XPath é baseado em 1, não em 0

# Clique no botão de salvar após a inserção de todos os valores
try:
    salvar_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[14]/div[3]/a[1]"))
    )
    salvar_button.click()
    print("Botão de salvar clicado com sucesso.")
except Exception as e:
    print(f"Erro ao clicar no botão de salvar: {e}")

# Aguarda para ver o resultado antes de encerrar
time.sleep(10)
driver.quit()
