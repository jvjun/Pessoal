import os
import shutil
import xml.etree.ElementTree as ET
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Definir os caminhos das pastas
folder_path = r'C:\Users\BackofficeNeogier\OneDrive - NEOGIER ENERGIA\Neogier Y\Backoffice\Programas\AutoNomeiaFatura\entrada'
destination_path = r'C:\Users\BackofficeNeogier\OneDrive - NEOGIER ENERGIA\Neogier Y\Backoffice\Programas\AutoNomeiaFatura'

# Listar todos os arquivos na pasta
files = os.listdir(folder_path)

# Filtrar os arquivos PDF e XML, ignorando maiúsculas e minúsculas
pdf_files = [file for file in files if file.lower().endswith('.pdf')]
xml_files = [file for file in files if file.lower().endswith('.xml')]

# Verificar se existe exatamente um arquivo PDF e um arquivo XML
if len(pdf_files) == 1 and len(xml_files) == 1:
    # Carregar o arquivo XML
    xml_file = xml_files[0]
    tree = ET.parse(os.path.join(folder_path, xml_file))
    root = tree.getroot()

    # Definir o namespace
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    # Extrair a data de emissão e subtrair um mês
    dhEmi = root.find('.//nfe:dhEmi', ns).text
    data_emissao = datetime.strptime(dhEmi[:10], '%Y-%m-%d')
    data_subtraida = data_emissao - relativedelta(months=1)
    mes_ano = data_subtraida.strftime('%m-%Y')  # Formato: "mes-ano"

    # Extrair as demais informações do XML
    emitente = root.find('.//nfe:emit/nfe:xFant', ns).text.split()[0]  # Emitente
    destinatario = root.find('.//nfe:dest/nfe:xNome', ns).text.split()[:1]  # Nome do destinatário (primeira palavra)
    destinatario = " ".join(destinatario)  # Unir as duas primeiras palavras com espaço
    municipio = root.find('.//nfe:dest/nfe:enderDest/nfe:xMun', ns).text  # Município do destinatário
    print(emitente)
    quantidade = root.find('.//nfe:det/nfe:prod/nfe:qCom', ns).text  # Quantidade

    # Construir o novo nome de arquivo
    new_name = f"{mes_ano}_NF_{destinatario}_{municipio}_{emitente}"

    # Renomear e mover os arquivos PDF e XML
    for file in pdf_files + xml_files:
        file_extension = os.path.splitext(file)[1]
        old_file = os.path.join(folder_path, file)
        new_file = os.path.join(folder_path, new_name + file_extension)
        
        # Renomear o arquivo
        os.rename(old_file, new_file)
        print(f"Renomeado: {file} para {new_name + file_extension}")

        # Mover o arquivo renomeado para o destino
        shutil.move(new_file, os.path.join(destination_path, new_name + file_extension))
        print(f"Movido para: {destination_path}")

else:
    print("A pasta deve conter exatamente um arquivo PDF e um arquivo XML.")

# Aviso de finalização
print("Processo de renomeação e movimentação concluído.")
