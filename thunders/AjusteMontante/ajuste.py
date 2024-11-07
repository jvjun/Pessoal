from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Caminho para o arquivo de login
login_file_path = r"C:\\Users\\joaoj\\Desktop\\Pessoal\\thunders\\AjusteMontante\\loginthunders.txt"
planilha_path = r"C:\\Users\\joaoj\\OneDrive - NEOGIER ENERGIA\\Neogier Y\\Backoffice\\Curto Prazo\\#Thunders 2.0.xlsx"

# Ler login e senha do arquivo
with open(login_file_path, "r") as file:
    credentials = file.readline().strip()
login, senha = credentials.split(";")

# Configurações do Chrome para baixar arquivos PDF automaticamente e abrir em janela maximizada
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": r"C:\\seu\\diretorio\\de\\download",  # Ajuste o diretório de download conforme necessário
    "plugins.always_open_pdf_externally": True,        # Baixar PDF em vez de abrir no Chrome
    "download.prompt_for_download": False,             # Não perguntar onde salvar cada arquivo
    "directory_upgrade": True                          # Faz o download automático para o diretório configurado
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--start-maximized")  # Abrir em janela maximizada

# Iniciar o navegador com o ChromeDriver gerenciado automaticamente
service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Lista para registrar operações que não foram encontradas
erros = []

try:
    # Acessar a página inicial
    driver.get("https://www.thunders.com.br/")
    time.sleep(2)  # Espera a página carregar

    # Clicar no botão para acessar o login
    login_button = driver.find_element(By.XPATH, "/html/body/div[4]/div[1]/a[2]/div")
    login_button.click()
    time.sleep(5)  # Aguardar a página de login abrir

    # Preencher o campo de login
    login_input = driver.find_element(By.XPATH, "/html/body/div/div/div[1]/form/fieldset/div/input")
    login_input.send_keys(login)

    # Clicar para continuar
    continue_button = driver.find_element(By.XPATH, "/html/body/div/div/div[1]/form/button")
    continue_button.click()
    time.sleep(2)  # Espera a página de senha carregar

    # Preencher o campo de senha
    password_input = driver.find_element(By.XPATH, "/html/body/div/div/div[1]/form/fieldset/div/input")
    password_input.send_keys(senha)

    # Clicar para fazer login
    login_button = driver.find_element(By.XPATH, "/html/body/div/div/div[1]/form/button")
    login_button.click()
    
    # Aguardar um momento para garantir que o login seja processado
    time.sleep(5)

    # Acessar a URL diretamente após o login
    driver.get("https://app.thunders.com.br/#/v2/operacoes/acl/vendas")
    time.sleep(5)  # Espera a página carregar completamente

    # Carregar a planilha com os dados de operação
    df = pd.read_excel(planilha_path, sheet_name='ThundersCodigo')

    # Iterar sobre cada linha da planilha
    for index, row in df.iterrows():
        # Obter o número de operação e o montante em MWh, formatando o montante para três casas decimais
        operacao_numero = row['OPERAÇÃO']
        montante_mwh = "{:.3f}".format(float(row['Montante MWh']))  # Formatação para três casas decimais com ponto decimal

        # Localizar o campo de input e inserir o número de operação
        operacao_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/ng-sidebar-container/div/div/div/app-acl/div/div/div/div/div/app-vendas/div[3]/ag-grid-angular/div/div[2]/div[1]/div[1]/div[2]/div/div[2]/div[1]/div/div/input"))
        )
        operacao_input.clear()  # Limpa o campo antes de inserir um novo valor
        operacao_input.send_keys(str(operacao_numero))
        time.sleep(2)  # Aguarda após o input para garantir que o valor seja inserido corretamente

        # Tentar localizar a linha da operação e clicar na engrenagem
        try:
            row_xpath = f"//div[contains(@class, 'ag-row') and .//a[text()='{operacao_numero}']]"
            row_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, row_xpath)))

            # Encontra o botão de engrenagem na linha da operação
            engrenagem_xpath = f"{row_xpath}//span[contains(@class, 'fa-cog')]"
            engrenagem_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, engrenagem_xpath)))
            engrenagem_icon.click()  # Clica na engrenagem
            time.sleep(2)  # Aguardar após o clique para garantir que o processo continue

            # Esperar a janela modal aparecer e inserir o montante
            montante_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/ngb-modal-window/div/div/div[2]/app-ajuste-acl/ngb-tabset/div/div[1]/table/tbody/tr/td[3]/input"))
            )
            montante_input.clear()
            montante_input.send_keys(montante_mwh)
            time.sleep(1)

            # Clicar no botão de confirmação
            confirmar_button = driver.find_element(By.XPATH, "/html/body/ngb-modal-window/div/div/div[2]/app-ajuste-acl/div/button[2]")
            confirmar_button.click()
            time.sleep(2)

            # Fechar a janela modal
            fechar_modal_button = driver.find_element(By.XPATH, "/html/body/ngb-modal-window/div/div/div[1]/button/span")
            fechar_modal_button.click()
            time.sleep(2)

        except Exception as e:
            # Se não conseguir encontrar a operação, adicionar à lista de erros
            print(f"Erro na operação {operacao_numero}: {str(e)}")
            erros.append(operacao_numero)
            continue

    # Imprimir log de erros
    if erros:
        print("Operações que não foram encontradas:")
        for erro in erros:
            print(f"- Operação {erro}")

    print("Automação concluída com sucesso!")
    input("Pressione Enter para fechar o navegador...")

finally:
    driver.quit()
