from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import shutil

# Dicionário de abreviações de meses em português
abreviacoes_meses_br = {
    "01": "jan",
    "02": "fev",
    "03": "mar",
    "04": "abr",
    "05": "mai",
    "06": "jun",
    "07": "jul",
    "08": "ago",
    "09": "set",
    "10": "out",
    "11": "nov",
    "12": "dez"
}

# Variável para definir manualmente o mês/ano
mes_ano = "09/2024"  # Altere manualmente para o mês/ano desejado

# Local de download dos arquivos PDF
diretorio_download = r"C:\Users\joaoj\OneDrive - NEOGIER ENERGIA\Holambra GD\Automação\Download de Fatura - Ceripa\Faturas\teste"

# Certifica-se de que o diretório existe
if not os.path.exists(diretorio_download):
    os.makedirs(diretorio_download)

# Local para salvar o log de erros
caminho_log_erros = r"C:\Users\joaoj\OneDrive - NEOGIER ENERGIA\Holambra GD\Automação\Download de Fatura - Ceripa\Faturas\teste\log_erros.txt"

# Função para escrever UCs com erro no log
def logar_erro_uc(uc, mensagem):
    with open(caminho_log_erros, "a") as log:
        log.write(f"UC {uc} - {mensagem}\n")

def ler_logins_senhas(caminho_arquivo):
    logins_senhas = []
    with open(caminho_arquivo, 'r') as arquivo:
        linhas = arquivo.readlines()
        for linha in linhas:
            partes = linha.strip().split(';')
            if len(partes) >= 4:
                login = partes[0]
                senha = partes[1]
                nome_cliente = partes[2]
                ucs = partes[3].split(',')  # Separa as UCs por vírgula
                logins_senhas.append((login, senha, nome_cliente, ucs))
    return logins_senhas

def abre_navegador_e_faz_login(login, senha):
    try:
        # Configurações do Chrome para baixar arquivos PDF automaticamente e abrir em janela maximizada
        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": diretorio_download,  # Define o diretório de download
            "plugins.always_open_pdf_externally": True,        # Baixar PDF em vez de abrir no Chrome
            "download.prompt_for_download": False,             # Não perguntar onde salvar cada arquivo
            "directory_upgrade": True                          # Faz o download automático para o diretório configurado
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--start-maximized")  # Abrir em janela maximizada

        # Iniciar o navegador com as configurações definidas
        servico = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico, options=chrome_options)

        # Acessar o site
        driver.get("https://ceripa.useallcloud.com.br/agenciavirtual3")

        # Aguarda alguns segundos para garantir que a página foi carregada
        time.sleep(3)

        # Encontra os campos de login, senha e o botão de login e faz o login
        campo_de_uc = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/form/article[2]/div[1]/div/input')
        campo_de_uc.send_keys(login)
        
        campo_de_senha = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/form/article[2]/div[2]/div/input')
        campo_de_senha.send_keys(senha)
        
        botao_login = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/form/div/button')
        botao_login.click()

        # Aguarda alguns segundos para o carregamento da próxima página após o login
        time.sleep(5)

        return driver
    except Exception as e:
        print(f"Erro ao logar com o login {login}. Pulando para o próximo. Erro: {e}")
        return None  # Retorna None em caso de erro

def acessar_historico(driver):
    # Clica no botão de menu que abre o histórico de faturas
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/header/div[2]/div[5]/div'))).click()

    # Aguarda o menu dropdown aparecer e clica no item dentro do menu
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/ul/li[1]'))).click()

    # Aguarda alguns segundos para o carregamento da página de faturas
    time.sleep(5)

def buscar_e_alterar_uc(driver, uc):
    # Localiza o campo de busca da UC e insere a UC atual
    campo_busca_uc = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section/section[2]/div[1]/div[1]/div/input'))
    )
    campo_busca_uc.clear()
    campo_busca_uc.send_keys(uc)

    # Espera alguns segundos para a UC aparecer na caixa abaixo
    time.sleep(2)

    try:
        # Verifica se a UC exibida na caixa abaixo corresponde à UC buscada
        uc_exibida = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section/section[2]/div[2]/div/div[2]/div/div[1]/div/a/button/div/h3'))
        ).text

        if uc_exibida.strip() == uc:
            # Se a UC exibida corresponde, clica na UC para acessar os detalhes
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'slick-slide')]//a"))
            ).click()
            time.sleep(5)
            return True
        else:
            print(f"UC {uc} não encontrada ou não corresponde ao código exibido.")
            logar_erro_uc(uc, "UC não encontrada ou não corresponde ao código exibido")
            return False

    except Exception as e:
        print(f"Erro ao buscar UC {uc}. Pulando para a próxima. Erro: {e}")
        logar_erro_uc(uc, "Erro ao buscar UC")
        return False

def buscar_fatura_por_mes(driver, mes_ano):
    # Divide o input Mês/Ano
    mes, ano = mes_ano.split("/")

    # Converte o número do mês para sua abreviação em português
    mes_abreviado = abreviacoes_meses_br[mes]
    data_buscada = f"{mes_abreviado}/{ano}".strip().lower()  # Removendo espaços extras e convertendo para minúsculas
    
    # Localiza o tbody onde estão as faturas, ou pula se não houver faturas
    try:
        tabela_faturas = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="root"]/div/section/section[5]/div/div/div[2]/table/tbody'))
        )
    except Exception as e:
        print(f"Erro ao buscar faturas para {mes_ano}. Pulando para a próxima UC.")
        return False
    
    # Busca todas as linhas de faturas dentro do tbody
    linhas_faturas = tabela_faturas.find_elements(By.TAG_NAME, 'tr')

    # Itera sobre as linhas da tabela
    for linha in linhas_faturas:
        data_fatura = linha.find_element(By.CSS_SELECTOR, 'td.campoData p').text.strip().lower()  # Removendo espaços extras e convertendo para minúsculas
        print(f"Comparando: {data_fatura} com {data_buscada}")  # Diagnóstico

        if data_fatura == data_buscada:
            print(f"Fatura encontrada para {data_buscada}")
            
            # Mover para o botão usando ActionChains
            botao_download = linha.find_element(By.CSS_SELECTOR, 'td.campoAcao svg.download')
            ActionChains(driver).move_to_element(botao_download).perform()
            time.sleep(1)  # Aguardar um segundo após mover o mouse

            # Tentar clicar no botão de download
            try:
                ActionChains(driver).click(botao_download).perform()
                print(f"Fatura de {data_buscada} baixada com sucesso via ActionChains.")
            except Exception as e:
                print(f"Erro ao clicar com ActionChains: {e}")

            return True
    
    print(f"Fatura para {data_buscada} não encontrada.")
    return False

def renomear_arquivo_pdf(nome_cliente, uc, mes_ano):
    # Aguarda que o arquivo PDF seja baixado (pode usar um loop para monitorar o diretório)
    time.sleep(10)  # Espera para garantir que o download seja concluído
    
    # Localiza o arquivo PDF mais recente baixado
    lista_arquivos = os.listdir(diretorio_download)
    lista_arquivos_pdf = [f for f in lista_arquivos if f.endswith(".pdf")]
    if not lista_arquivos_pdf:
        print("Nenhum arquivo PDF foi baixado.")
        return

    # Nome do PDF mais recente
    arquivo_pdf = max([os.path.join(diretorio_download, f) for f in lista_arquivos_pdf], key=os.path.getctime)

    # Novo nome do arquivo
    novo_nome_pdf = f"{mes_ano.replace('/', '-')}_Fatura_{nome_cliente}_{uc}.pdf"
    novo_caminho_pdf = os.path.join(diretorio_download, novo_nome_pdf)

    # Renomeia o arquivo
    try:
        shutil.move(arquivo_pdf, novo_caminho_pdf)
        print(f"Arquivo PDF renomeado para: {novo_nome_pdf}")
    except Exception as e:
        print(f"Erro ao renomear o arquivo PDF: {e}")

# Função principal
def main():
    caminho_arquivo_login = r"C:\Users\joaoj\OneDrive - NEOGIER ENERGIA\Holambra GD\Automação\Download de Fatura - Ceripa\Faturas\teste\theo.txt"
    
    logins_senhas = ler_logins_senhas(caminho_arquivo_login)

    for login, senha, nome_cliente, ucs in logins_senhas:
        driver = abre_navegador_e_faz_login(login, senha)
        if driver is None:
            continue  # Pula para o próximo login se houver erro no login
        
        # Acessa o histórico de faturas
        try:
            acessar_historico(driver)
        except Exception as e:
            print(f"Erro ao acessar o histórico para {login}. Pulando para o próximo. Erro: {e}")
            driver.quit()
            continue  # Pula para o próximo login
        
        for uc in ucs:
            print(f"Processando UC: {uc}")
            try:
                if buscar_e_alterar_uc(driver, uc):  # Muda a UC e pesquisa
                    # Busca e baixa a fatura para o mês/ano especificado na variável 'mes_ano'
                    if buscar_fatura_por_mes(driver, mes_ano):
                        # Renomeia o PDF após o download
                        renomear_arquivo_pdf(nome_cliente, uc, mes_ano)
            except Exception as e:
                print(f"Erro ao processar a UC {uc} para o cliente {nome_cliente}. Pulando para a próxima UC. Erro: {e}")
                logar_erro_uc(uc, "Erro no processamento da UC")
                continue  # Pula para a próxima UC em caso de erro

        # Fechar o navegador para o próximo login
        driver.quit()

    input("Pressione Enter para encerrar o programa...")

# Executa o código
main()
