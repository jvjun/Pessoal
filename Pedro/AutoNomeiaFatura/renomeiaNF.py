import os
import xml.etree.ElementTree as ET
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

# Caminhos dos arquivos e pastas
folder_path = r'C:\Users\joaoj\Desktop\Pessoal\Pedro\AutoNomeiaFatura\entrada'
excel_path = r'C:\Users\joaoj\Desktop\Pessoal\Pedro\AutoNomeiaFatura\Relacao.xlsx'

# Função para encontrar o primeiro arquivo XML na pasta
def get_xml_file(folder):
    for file in os.listdir(folder):
        if file.lower().endswith('.xml'):  # Verifica se é um arquivo XML
            return os.path.join(folder, file)
    return None

# Função para extrair a data de emissão do XML e calcular o mês anterior
def extract_and_adjust_date(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    dhEmi = root.find('.//nfe:ide/nfe:dhEmi', ns)
    if dhEmi is not None:
        date_emission = datetime.strptime(dhEmi.text[:10], '%Y-%m-%d')
        adjusted_date = date_emission - relativedelta(months=1)
        return adjusted_date.strftime('%m-%Y')
    return None

# Função para extrair o CNPJ do XML
def extract_cnpj_from_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    cnpj_dest = root.find('.//nfe:dest/nfe:CNPJ', ns)
    return cnpj_dest.text if cnpj_dest is not None else None

# Função para extrair o nome do emitente do XML
def extract_emit_name(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    emit_name = root.find('.//nfe:emit/nfe:xNome', ns)
    if emit_name is not None:
        return emit_name.text.split()[0]  # Retorna apenas a primeira palavra
    return None

# Função para buscar informações na planilha Excel
def find_info_in_excel(cnpj, excel_file):
    excel_data = pd.read_excel(excel_file, usecols=[0, 1, 2], header=None, names=["CNPJ", "Agente", "Unidade"])
    row = excel_data.loc[excel_data["CNPJ"].astype(str) == cnpj]
    if not row.empty:
        agente = row["Agente"].values[0]
        unidade = row["Unidade"].values[0]
        return agente, unidade
    return None, None

# Função para renomear os arquivos
def rename_files(folder, new_name):
    files = os.listdir(folder)
    for file in files:
        old_file = os.path.join(folder, file)
        if os.path.isfile(old_file):
            file_extension = os.path.splitext(file)[1]
            new_file = os.path.join(folder, new_name + file_extension)
            os.rename(old_file, new_file)
            print(f"Arquivo renomeado de {file} para {new_name + file_extension}")

# Localizar o arquivo XML na pasta
xml_path = get_xml_file(folder_path)

if xml_path:
    adjusted_date = extract_and_adjust_date(xml_path)
    cnpj_dest = extract_cnpj_from_xml(xml_path)
    emit_name = extract_emit_name(xml_path)

    if adjusted_date:
        print(f"A data ajustada (mês anterior) é: {adjusted_date}")

    if cnpj_dest:
        agente, unidade = find_info_in_excel(cnpj_dest, excel_path)
        if agente and unidade:
            print(f"CNPJ encontrado. Agente: {agente}, Unidade: {unidade}")
            # Formatar o novo nome
            new_name = f"{adjusted_date}_{agente}_{unidade}_{emit_name}"
            # Renomear os arquivos na pasta
            rename_files(folder_path, new_name)
        else:
            print(f"O CNPJ {cnpj_dest} NÃO foi encontrado na planilha.")
    else:
        print("Não foi possível encontrar o CNPJ no XML.")
else:
    print("Nenhum arquivo XML foi encontrado na pasta.")
