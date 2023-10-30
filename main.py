import win32com.client
import pandas as pd
import os
import re
from bs4 import BeautifulSoup
import zipfile
import xml.etree.ElementTree as ET

# Função para extrair texto de um anexo .zip
def extract_text_from_zip(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zf:
        for file_info in zf.infolist():
            if file_info.filename.endswith('.xml'):
                with zf.open(file_info) as xml_file:
                    xml_content = xml_file.read().decode('utf-8')
                return xml_content

# Acesse a pasta "xml" no Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
folder = namespace.GetDefaultFolder(6)  # Acessa a Caixa de Entrada
subfolder = folder.Folders.Item("XML")  # Acessa a subpasta "xml"

# Expressão regular para encontrar o número após "CT-e:"
cte_pattern = r"CT-e:\s+([0-9-]+)"

# Lista para armazenar os dados
data_list = []
nome_list = []  # Inicialize a lista para os nomes
comp_data = {}  # Inicialize o dicionário para os componentes

# Diretório temporário para salvar os anexos
temp_dir = os.path.join(os.getcwd(), 'temp')
os.makedirs(temp_dir, exist_ok=True)

# Itera pelos e-mails na subpasta "xml"
for email in subfolder.Items:
    for attachment in email.Attachments:
        file_extension = attachment.FileName.split('.')[-1]
        if file_extension == 'xml':
            # Salva o anexo XML em um arquivo temporário
            temp_file_path = os.path.join(temp_dir, attachment.FileName)
            attachment.SaveAsFile(temp_file_path)
            with open(temp_file_path, 'r', encoding='utf-8') as file:
                xml_content = file.read()

        elif file_extension == 'zip':
            # Extrai o texto do anexo .zip
            temp_file_path = os.path.join(temp_dir, attachment.FileName)
            attachment.SaveAsFile(temp_file_path)
            text_content = extract_text_from_zip(temp_file_path, temp_dir)
            xml_content = text_content  # Supondo que o texto do .zip seja XML
        else:
            continue

        soup = BeautifulSoup(xml_content, 'xml')

        # Localize o elemento <receb>
        receb_element = soup.find('receb')

        # Verifique se o elemento <receb> existe
        if receb_element:
            # Localize o elemento <xNome> dentro de <receb>
            xNome_element = receb_element.find('xNome')

            # Verifique se o elemento <xNome> existe
            if xNome_element:
                # Extraia o texto de <xNome>
                nome_receb = xNome_element.text

        else:
            print("Elemento <receb> não encontrado.")


        rem_element = soup.find('rem')

        # Verifique se o elemento <rem> existe
        if rem_element:
            # Localize o elemento <xNome> dentro de <rem>
            xNome_element = rem_element.find('xNome')

            # Verifique se o elemento <xNome> existe
            if xNome_element:
                # Extraia o texto de <xNome>
                nome_rem = xNome_element.text

        else:
            print("Elemento <rem> não encontrado.")

        # Localiza a razão social do emitente
        nome = soup.find('emit').find('xNome').get_text() if soup.find('emit') else None

        # Localiza o cte original
        cte_number = re.search(cte_pattern, xml_content).group(1) if re.search(cte_pattern, xml_content) else None

        # Localiza o valor da prestação do serviço
        vTPrest = soup.find("vTPrest").text if soup.find("vTPrest") else None

        infQ_elements = soup.find_all("infQ")
        infQ_data = {infQ.find("tpMed").text: infQ.find("qCarga").text for infQ in infQ_elements if
                     infQ.find("tpMed") and infQ.find("qCarga")}

        vCargaAverb = soup.find("vCarga").text if soup.find("vCarga") else None

        comp_elements = soup.find_all("xNome")
        comp_data = {comp.text: comp.find_next("vComp").text for comp in comp_elements if comp.find_next("vComp")}

        match = re.search(r'<nCT>(\d+)</nCT>', xml_content)
        nCT = match.group(1) if match else None


        # localiza data da emissão
        dhEmi_element = soup.find("dhEmi")
        dhEmi = dhEmi_element.text.split("T")[0] if dhEmi_element else None

        root = ET.fromstring(xml_content)

        # Organiza os dados da planilha excel
        data_list.append({
            'Nome': nome,
            'Data de Emissão': dhEmi,
            'CTE Parceiro': nCT,
            'CTE': cte_number,
            'Remetente': nome_rem,
            'Destinatario': nome_receb,
            'Vlr frete': vTPrest,
            'Frete Peso': comp_data.get('FRETE PESO', ''),
            'GRIS': comp_data.get('GRIS', ''),
            'Pedágio': comp_data.get('PEDAGIO', ''),
            'M3': infQ_data.get('M3', ''),
            'PESO REAL': infQ_data.get('PESO REAL', ''),
            'PESO BASE DE CALCULO': infQ_data.get('PESO BASE DE CALCULO', ''),
            'Vlr mercadoria': vCargaAverb
        })

# Crie um DataFrame a partir dos dados extraídos
df = pd.DataFrame(data_list)


# Converte os valores em float
colunas_para_converter = ['Vlr frete', 'Frete Peso', 'GRIS', 'Pedágio', 'M3', 'PESO REAL','PESO BASE DE CALCULO','Vlr mercadoria']
df['Data de Emissão'] = pd.to_datetime(df['Data de Emissão'], format='%Y-%m-%d', errors='coerce')

# Loop para converter as colunas para float
for coluna in colunas_para_converter:
    df[coluna] = pd.to_numeric(df[coluna], errors='coerce')


# Salva o DataFrame em um arquivo Excel
df.to_excel('cte_data.xlsx', index=False)


# Remova os arquivos temporários
for file in os.listdir(temp_dir):
    file_path = os.path.join(temp_dir, file)
    os.remove(file_path)

# Remova o diretório temporário
os.rmdir(temp_dir)