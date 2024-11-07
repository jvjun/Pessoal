import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# Configurações globais para o Chrome e o driver
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Função para acessar a página inicial
def acessar_site():
    driver.get("https://operacao.ccee.org.br/ui/scde/analise/grafico")
    time.sleep(5)  # Tempo para garantir o carregamento da página

# Função para login
def login(usuario, senha):
    username_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[1]/input")
    username_field.send_keys(usuario)
    password_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[2]/input")
    password_field.send_keys(senha)
    login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[4]/button")
    login_button.click()

# Função para aguardar a verificação do usuário
def aguardar_verificacao():
    input("Por favor, insira o código de verificação manualmente no site e pressione Enter para continuar...")

# Função para mudar para o iframe
def mudar_para_iframe():
    try:
        iframe = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[@src='https://scde.ccee.org.br/relatorios']"))
        )
        driver.switch_to.frame(iframe)
        print("Mudança para o iframe bem-sucedida.")
    except Exception as e:
        print("Erro ao localizar o iframe:", e)
        driver.quit()
        exit()

# Função para selecionar opções no dropdown e agendar
def selecionar_opcoes_e_agendar():
    try:
        dropdown_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[3]/div[1]/table[2]/tbody/tr/td[1]/select"))
        )
        dropdown_button.click()

        first_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[3]/div[1]/table[2]/tbody/tr/td[1]/select/option[1]"))
        )
        first_option.click()

        confirm_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[3]/div[2]/div/div[1]/button"))
        )
        confirm_button.click()
    except Exception as e:
        print("Erro ao selecionar as opções:", e)
        driver.quit()
        exit()

    # Inserir a data atual e clicar no botão de agendamento
    data_atual = datetime.now().strftime("%d-%m-%y")

    # Espera até que o campo de data esteja visível e interativo
    try:
        data_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/input"))
        )
        data_field.click()
        data_field.clear()  # Limpa o campo antes de inserir dados
        data_field.send_keys(data_atual)
        
        # Clicar no botão de agendamento
        agendar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[3]/button[1]"))
        )
        agendar_button.click()
    except Exception as e:
        print("Erro ao interagir com o campo de data ou agendar:", e)

# Função para exportar os dados
def exportar_dados(data_atual):
    try:
        # Localiza e clica no botão de exportação
        exportar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/ul/li[4]/a"))
        )
        exportar.click()
        
        # Insere a data no campo de busca
        busca_data = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/form/div/div/div[1]/div[2]/div/label/input"))
        )
        busca_data.clear()  # Limpa o campo antes de inserir a data
        busca_data.send_keys(data_atual)
        
        time.sleep(20)

        # Aguarda e clica no botão de resultado
        botao_resultado = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/form/div/div/table/tbody/tr/td[8]/div/a/i"))
        )
        botao_resultado.click()
        print("Botão de resultado clicado com sucesso.")
    except Exception as e:
        print("Erro ao exportar os dados:", e)
        driver.quit()
        exit()

# Função principal para rodar o código completo
def main():
    acessar_site()
    login("joaojun01", "Senhafoda!4")
    aguardar_verificacao()
    mudar_para_iframe()
    selecionar_opcoes_e_agendar()
    data_atual = datetime.now().strftime("%d-%m-%y")
    exportar_dados(data_atual)

# Roda a função principal apenas se o script for executado diretamente
if __name__ == "__main__":
    main()

    # Mantenha o navegador aberto para observar o comportamento
    input("Pressione Enter para fechar o navegador...")

    # Feche o navegador após o uso
    driver.quit()
