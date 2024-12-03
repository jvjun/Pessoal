import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import win32com.client as win32

# Configurações globais para o Chrome e o driver
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Função para acessar a página inicial
def acessar_site():
    driver.get("https://operacao.ccee.org.br/ui/scde/analise/grafico")
    time.sleep(5)

# Função para login
def login(usuario, senha):
    username_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[1]/input")
    username_field.send_keys(usuario)
    password_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[2]/input")
    password_field.send_keys(senha)
    login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[4]/button")
    login_button.click()

# Função para solicitar o código de verificação
def solicitar_codigo_verificacao():
    try:
        solicitar_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div/form/div[1]/div[2]/ul/li[1]/button"))
        )
        solicitar_button.click()
        print("Botão de solicitação de código clicado.")
        time.sleep(10)  # Aguarda o e-mail chegar após a solicitação
    except Exception as e:
        print("Erro ao solicitar código de verificação:", e)

# Função para buscar o código de verificação mais recente no Outlook após a solicitação
def buscar_codigo_verificacao_outlook():
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 é a pasta de entrada padrão
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Ordena do mais recente para o mais antigo
        
        for message in messages:
            if "CCEE: Código de verificação" in message.Subject:
                body = message.Body
                codigo = re.search(r'\b\d{6}\b', body)
                if codigo:
                    print(f"Código de verificação encontrado: {codigo.group()}")
                    return codigo.group()
        print("Código de verificação não encontrado.")
        return None
    except Exception as e:
        print("Erro ao buscar o código de verificação no Outlook:", e)
        return None

# Função para verificar se o código foi aceito ou é inválido
def codigo_invalido():
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'código inválido')]"))
        )
        print("Código inválido detectado.")
        return True
    except:
        return False

# Função para inserir o código de verificação e confirmar
def inserir_codigo_verificacao(codigo):
    try:
        codigo_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div/form/div[1]/div[2]/input"))
        )
        codigo_field.clear()
        codigo_field.send_keys(codigo)

        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div/form/div[2]/div[2]/input"))
        )
        ok_button.click()
        print("Código de verificação inserido e botão OK clicado.")
        time.sleep(3)  # Pausa breve para validação do código
    except Exception as e:
        print("Erro ao inserir o código de verificação:", e)

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

# Função principal para rodar o código completo
def main():
    acessar_site()
    login("joaojun01", "Senhafoda!4")
    solicitar_codigo_verificacao()
    
    # Primeira tentativa: busca o código mais recente
    codigo_verificacao = buscar_codigo_verificacao_outlook()
    if codigo_verificacao:
        inserir_codigo_verificacao(codigo_verificacao)
        
        # Caso o código seja inválido, busque novamente o próximo código
        while codigo_invalido():
            print("Tentando o próximo código...")
            codigo_verificacao = buscar_codigo_verificacao_outlook()
            if codigo_verificacao:
                inserir_codigo_verificacao(codigo_verificacao)
            else:
                print("Não foi encontrado um novo código de verificação.")
                break

    mudar_para_iframe()

# Roda a função principal apenas se o script for executado diretamente
if __name__ == "__main__":
    try:
        main()
        input("Pressione Enter para fechar o navegador...")
    finally:
        driver.quit()
