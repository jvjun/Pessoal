import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import win32com.client as win32
from datetime import datetime
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

# Configurações globais para o Chrome e o driver
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Configurações do Outlook
outlook_app = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def acessar_site():
    """Acessa a página inicial."""
    driver.get("https://operacao.ccee.org.br/ui/scde/analise/grafico")
    time.sleep(5)


def login(usuario, senha):
    """Realiza login na página."""
    username_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[1]/input")
    username_field.send_keys(usuario)
    password_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[2]/input")
    password_field.send_keys(senha)
    login_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div[1]/div/form/div[4]/button")
    login_button.click()


def solicitar_codigo_verificacao():
    """Clica no botão para solicitar o código de verificação."""
    try:
        solicitar_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div/form/div[1]/div[2]/ul/li[1]/button"))
        )
        solicitar_button.click()
        print("Botão de solicitação de código clicado.")
        time.sleep(15)  # Aguarda o e-mail chegar após a solicitação
    except Exception as e:
        print("Erro ao solicitar código de verificação:", e)


def limpar_emails_codigo_verificacao():
    """Remove todos os e-mails com o assunto específico."""
    try:
        inbox = outlook_app.GetDefaultFolder(6)  # 6 se refere à pasta "Caixa de Entrada"
        messages = inbox.Items
        for message in messages:
            if "CCEE: Código de verificação" in message.Subject:
                print(f"Removendo e-mail: {message.Subject}")
                message.Delete()
        print("Todos os e-mails de código de verificação foram removidos.")
    except Exception as e:
        print("Erro ao limpar e-mails:", e)


def buscar_codigo_acesso():
    """
    Aguarda e busca o código de verificação mais recente no Outlook com o assunto específico.
    """
    inbox = outlook_app.GetDefaultFolder(6)  # 6 se refere à pasta "Caixa de Entrada"
    
    # Tempo máximo de espera em segundos e intervalo entre tentativas
    tempo_maximo_espera = 60  # Aguarda até 1 minuto
    intervalo_entre_tentativas = 5  # Verifica a cada 5 segundos

    # Loop de tentativa
    for tentativa in range(0, tempo_maximo_espera, intervalo_entre_tentativas):
        # Ordena os e-mails por data do mais recente para o mais antigo
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        # Itera sobre os e-mails para encontrar o correto
        for message in messages:
            if "CCEE: Código de verificação" in message.Subject:
                print(f"E-mail encontrado: {message.Subject}")

                # Obtém o corpo do e-mail
                corpo = message.Body

                # Procura o padrão do código de acesso usando regex
                match = re.search(r"CCEE: o seu codigo de acesso e (\d{6})", corpo)
                if match:
                    codigo = match.group(1)
                    print(f"Código encontrado: {codigo}")
                    return codigo
        
        # Se não encontrar o código, espera antes de tentar novamente
        print("Código não encontrado, aguardando...")
        time.sleep(intervalo_entre_tentativas)

    print("Nenhum código de verificação foi encontrado dentro do tempo limite.")
    return None

def inserir_codigo_verificacao(codigo):
    """Insere o código de verificação e confirma."""
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
        print("Código de verificação inserido com sucesso.")
        time.sleep(3)
    except Exception as e:
        print("Erro ao inserir o código de verificação:", e)
    time.sleep(3)

def executar_fluxo_pos_login():
    """
    Seleciona o iframe.
    """
    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/body/cpit-app/fx-app-shell/div/div/mat-sidenav-container/mat-sidenav-content/main/cpit-iframe/fx-iframe/iframe"))
    # Seleciona o dropdown
    time.sleep(5)
    dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="base"]'))
    )
    dropdown.click()
    print("Dropdown encontrado e clicado com sucesso.")

    # Seleciona a opção 'hora'
    opcao_hora = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[3]/div[1]/table[2]/tbody/tr/td[1]/select/option[1]"))
    )
    opcao_hora.click()
    print("Opção 'hora' selecionada.")

    # Clica no botão 'Todos'
    botao_todos = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[3]/div[2]/div/div[1]/button"))
    )
    botao_todos.click()
    print("Botão 'Todos' clicado.")

    # Insere o valor
    input_but = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/input"))
    )
    input_but.click()
    print("Botão 'Todos' clicado.")

    # Gerar a data atual no formato DD-MM
    data_atual = datetime.now().strftime("%d-%m")   

    campo_data = driver.find_element(By.XPATH, '//*[@id="modal-descricao-job"]')

    # Limpar o campo e inserir a data atual
    campo_data.clear()
    campo_data.send_keys(data_atual)

    print(f"Data atual '{data_atual}' inserida com sucesso.")

    confirma = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="modal-gerar-job"]'))
    )
    confirma.click()
    print("Botão 'Todos' clicado.")
    time.sleep(5)
    
    driver.get("https://scde.ccee.org.br/relatorios/Exportacao")

def verificar_status_download():
    """
    Verifica o status do download (Sucesso ou Executando) e retorna True se estiver pronto.
    """
    try:
        # Localiza o elemento de status
        status_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="dataTableList"]/tbody/tr[1]/td[5]'))
        )
        
        # Captura o texto do status
        status_texto = status_element.text.strip()
        print(f"Status atual: {status_texto}")
        
        # Verifica se o status é 'Sucesso'
        if status_texto == "Sucesso":
            print("Download está pronto!")
            return True
        elif status_texto == "Executando":
            print("Download ainda em execução.")
            return False
        else:
            print(f"Status desconhecido: {status_texto}")
            return False
    except Exception as e:
        print(f"Erro ao verificar status do download: {e}")
        return False

def aguardar_download():
    """
    Aguarda até o download estar pronto, verificando o status a cada intervalo.
    """
    max_tentativas = 10  # Número máximo de verificações
    intervalo = 5  # Intervalo de tempo entre verificações (em segundos)

    for tentativa in range(max_tentativas):
        print(f"Tentativa {tentativa + 1} de {max_tentativas}")
        if verificar_status_download():
            download = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="dataTableList"]/tbody/tr[1]/td[8]/div/a'))
        )
            download.click()
            print("Download pronto! Prosseguindo com a próxima etapa.")
            return True  # Sai da função
        time.sleep(intervalo)  # Aguarda antes de verificar novamente

    print("Tempo esgotado. O download não foi concluído.")
    return False


def main():
    limpar_emails_codigo_verificacao()
    acessar_site()
    login("joaojun01", "Senhafoda!4")
    solicitar_codigo_verificacao()

    # Busca e insere o código de verificação
    codigo_verificacao = buscar_codigo_acesso()
    if codigo_verificacao:
        inserir_codigo_verificacao(codigo_verificacao)
    else:
        print("Não foi possível encontrar o código de verificação. Encerrando.")
        return

    # Executa ações após login: seleciona 'hora' e clica no botão
    executar_fluxo_pos_login()

    #verifica se da pra baixar ou não
    aguardar_download()

if __name__ == "__main__":
    try:
        main()
        input("Pressione Enter para sair...")
    finally:
        driver.quit()