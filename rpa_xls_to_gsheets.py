# -*- coding: utf-8 -*-

#%%

from time import sleep
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException
from datetime import datetime, timedelta
import re
import os
import ctypes
import win32com.client as win32
import gspread
import csv
import requests
import pdfplumber
import io
from bs4 import BeautifulSoup

#%%

options = Options()
options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-notifications")
options.add_argument("--disable-gcm-registration")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_argument("--log-level=3")  # reduz nível de log do Chrome

options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\RPA\se_suite_xls",
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

service = Service(log_path="NUL")

driver = Chrome(service=service, options=options)

#Dados Google Sheets
gc = gspread.service_account(filename=os.path.join(os.path.dirname(os.getcwd()), 'crested-century-386316-01c90985d6e4.json'))

spreadsheet = gc.open("Acompanhamento_Aquisições_RPA")
worksheet = spreadsheet.worksheet("Dados")

spreadsheet_proj_fin = gc.open("proj_fin")
worksheet_proj_fin = spreadsheet_proj_fin.worksheet("Auxiliar")

dados_proj_fin = worksheet_proj_fin.get_all_records()
auxiliar_proj_fin = pd.DataFrame(dados_proj_fin)

def obter_apelido_projeto(codigo_extraido):
    if not codigo_extraido:
        return ""

    try:
        codigo_extraido_int = int(float(codigo_extraido))
    except:
        return ""

    linha = auxiliar_proj_fin.loc[
        auxiliar_proj_fin["cd_projeto"] == codigo_extraido_int
    ]

    if linha.empty:
        return ""

    #nome_projeto_planilha = linha.iloc[0].get("nm_projeto", "")
    apelido = linha.iloc[0].get("nm_apelido_projeto", "")

    if apelido in ["Annelida2 - ISI SE", "Annelida2 - ISI SM"]:
        return "Annelida 2"

    return apelido


def login_sesuite():
    davpass = open(os.path.join(os.path.dirname(os.getcwd()), '.cpass'), 'r').read()    
    print("Acessando SE Suite...")
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
    #driver.maximize_window()

    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="userNameInput"]'))
    )
    username_input.send_keys('davi.hulse@sc.senai.br')

    password_input = driver.find_element(By.XPATH, '//*[@id="passwordInput"]')
    password_input.send_keys(davpass + Keys.ENTER)
    
    print("Login no SE Suite realizado.")


#%%

login_sesuite()

janela_principal = driver.window_handles[0]

def baixar_xls():
    
    #WebDriverWait(driver, 100).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    sleep(1)
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=tracking,104,2')
    
    #WebDriverWait(driver, 100).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    sleep(1)
    
    WebDriverWait(driver, 100).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe"))
    )
    sleep(1)
    
    # botão seta
    botao_seta = WebDriverWait(driver, 100).until(
        EC.element_to_be_clickable((By.ID, "se_admin_btnreport-menuButton"))
    )
    botao_seta.click()
    
    # "Exportar para Excel"
    botao_exportar = WebDriverWait(driver, 100).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Exportar para Excel")]'))
    )
    botao_exportar.click()
    
    #print("Baixando arquivo XLS...")
    
    sleep(1)
    
    caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"
    inicio = time.time()
    timeout = 600
    
    while time.time() - inicio < timeout:
        if os.path.exists(caminho) and not os.path.exists(caminho + ".crdownload"):
            print("Baixando arquivo XLS...")
            #print("Convertendo arquivo para XLSX...")
            break
        time.sleep(2)
    else:
        raise TimeoutError("Download não terminou dentro do tempo esperado.")
    
    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()
    
    # Volta para janela principal
    driver.switch_to.window(janela_principal)

#%%

def desbloquear_arquivo_excel(caminho_arquivo):
    # Remove a marca de "baixado da internet"
    try:
        os.remove(caminho_arquivo + ":Zone.Identifier")
    except FileNotFoundError:
        pass  # já estava desbloqueado

    # Garante permissões normais
    ctypes.windll.kernel32.SetFileAttributesW(caminho_arquivo, 0x80)  # FILE_ATTRIBUTE_NORMAL


#%%

def converter_xls_para_xlsx(caminho_xls, caminho_xlsx):
    print("Convertendo arquivo para XLSX...")
    excel = win32.DispatchEx('Excel.Application')  # cria nova instância
    excel.Visible = False  # Excel rodando "invisível"
    excel.DisplayAlerts = False  # Evita pop-ups e confirmações

    wb = excel.Workbooks.Open(caminho_xls)
    wb.SaveAs(caminho_xlsx, FileFormat=51)  # 51 = .xlsx
    wb.Close(False)
    excel.Quit()

#%%

caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"

### Comentar as 3 linhas abaixo para pular o download do XLS.
baixar_xls()
desbloquear_arquivo_excel(caminho)
converter_xls_para_xlsx(caminho,r"C:\RPA\se_suite_xls\relatorio_convertido.xlsx")


if os.path.exists(caminho):
    os.remove(caminho)
    print("Arquivo original excluído.")
else:
    print("Arquivo original não encontrado para exclusão.")

#%%

#Lê arquivo baixado do SE Suite
df = pd.read_excel(r"C:\RPA\se_suite_xls\relatorio_convertido.xlsx")

# Acessa as abas "Manuais", "Ignorar" e "ANS" da mesma planilha
worksheet_manuais = spreadsheet.worksheet("Manuais")
worksheet_ignorar = spreadsheet.worksheet("Ignorar")
worksheet_ans = spreadsheet.worksheet("ANS")

# Lê os valores das abas coluna A (sem cabeçalho)
valores_manuais = worksheet_manuais.col_values(1)
valores_ignorar = worksheet_ignorar.col_values(1)

# Lê as colunas A e B (ANS)
valores_ans = worksheet_ans.get_all_values()

# Aplica validação e conversão com zfill(6)
lista_manuais = [
    str(int(float(x))).zfill(6)
    for x in valores_manuais
    if str(x).replace('.', '', 1).isdigit()
]

lista_manuais = sorted(set(lista_manuais))

lista_ignorar = set(
    str(int(float(x))).zfill(6)
    for x in valores_ignorar
    if str(x).replace('.', '', 1).isdigit()
)

# Cria dicionário {Identificador: ANS} para ANS
mapa_ans_custom = {
    str(int(float(linha[0]))).zfill(6): linha[1]
    for linha in valores_ans
    if len(linha) >= 2 and linha[0].replace('.', '', 1).isdigit()
}

def remover_chamado_manuais(ws, numero_formatado):
    
    todas_linhas = ws.get_all_values()

    for i, linha in enumerate(todas_linhas):
        if not linha or not linha[0].strip():  # ignora linhas vazias
            continue

        valor_bruto = linha[0].strip()

        # só tenta converter se for numérico (com ponto ou não)
        if not valor_bruto.replace('.', '', 1).isdigit():
            continue

        # trata tanto "3070" quanto "003070" como "003070"
        valor_tratado = str(int(float(valor_bruto))).zfill(6)

        if valor_tratado == numero_formatado:
            ws.delete_rows(i + 1)  # gspread é 1-based
            print(f"🧹 Chamado {numero_formatado} removido da lista manual.")
            return

    print(f"⚠️ Erro ao tentar remover o chamado {numero_formatado}: não encontrado.")


#%%

df = df.iloc[1:].reset_index(drop=True)
df = df.drop(columns=['P', 'S', 'SW', 'SLA', 'PR', 'D', 'A', 'Executor', 'Processo', 'Tipo de workflow'])

#Desconsiderar atividades específicas:
#df = df[~df["Atividade habilitada"].str.startswith("Confirmar recebimento  do item solicitado", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Analisar pertinência da solicitação", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Solicitar aquisição", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Tomar ciência da negativa da solicitação", na=False)]

#Desconsidera identificadores que começam com E-PROC
df = df[~df["Identificador"].astype(str).str.startswith("E-PROC", na=False)]


df["AtividadeHabilitadaFiltrada"] = df["Atividade habilitada"].str.split("(", n=1).str[0].str.strip()
print(df["AtividadeHabilitadaFiltrada"].value_counts())
print()
print('Total de chamados: {}' .format(df["Identificador"].count()))


#%%

num_chamados = df["Identificador"].apply(lambda x: str(int(float(x))).zfill(6) if pd.notnull(x) else "").tolist()
objetos_compra = df["Título"].tolist()
atividadehabilitada = df["AtividadeHabilitadaFiltrada"].tolist()


#%%

def tratar_alerta(driver):
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print("⚠️ Alerta detectado:", alert.text)
        alert.accept()
        print("✅ Alerta aceito.")
        sleep(1)
        return True
    except (NoAlertPresentException, TimeoutException):
        return False
    except UnexpectedAlertPresentException:
        # Se o alerta aparece no meio de uma ação
        try:
            driver.switch_to.alert.accept()
            print("✅ Alerta inesperado aceito.")
            return True
        except:
            return False

#%%

def data_hoje_ontem(data_txt):
    data_txt = data_txt.strip()
    data_lower = data_txt.lower()
    if data_lower.endswith("hoje"):
        return datetime.today()
    elif data_lower.endswith("ntem"):
        return datetime.today() - timedelta(days=1)
    else:
        # Extrai dd/mm/aaaa de strings como "NOME SOBRENOME 15/09/2025 - 08:37"
        import re
        match = re.search(r'\d{2}/\d{2}/\d{4}', data_txt)
        if match:
            return datetime.strptime(match.group(), "%d/%m/%Y")
        return None
    
#%%    
#import re

def extrair_dados_oc(texto_pdf):
    numero_oc = ""
    data_emissao = ""
    nome_fornecedor_pdf = ""
    cnpj_fornecedor_pdf = ""
    prazo_entrega_pdf = ""
    
    primeiras_linhas = "\n".join(texto_pdf.splitlines()[:5])
    #print(f"🔍 Primeiras linhas:\n{primeiras_linhas}")

    if "Número AF:" in primeiras_linhas:
        # Modelo 1
        texto_limpo = re.sub(r'\(cid:\d+\)', ' ', texto_pdf)
        match_num = re.search(r'Número AF:\s*(\d+)', texto_limpo)
        match_data = re.search(r'Data:\s*(\d{2}/\d{2}/\d{4})', texto_limpo)
        match_fornecedor = re.search(r'Razão social:\s*(.+)', texto_limpo)
        match_cnpj = re.search(r'DADOS DO FORNECEDOR.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_limpo, re.DOTALL)
        match_prazo = re.search(r'Prazos de entrega.*?/\s*(\d{2}/\d{2}/\d{4})', texto_limpo, re.DOTALL)


        numero_oc = match_num.group(1) if match_num else ""
        data_emissao = match_data.group(1) if match_data else ""
        nome_fornecedor_pdf = match_fornecedor.group(1).strip() if match_fornecedor else ""
        cnpj_fornecedor_pdf = match_cnpj.group(1).strip() if match_cnpj else ""
        prazo_entrega_pdf = match_prazo.group(1).strip() if match_prazo else ""

    elif "Ordem de compra" in primeiras_linhas:
        # Modelo 2
        match_num = re.search(r'Nº\s+(\d+)\s+Valor Total:', texto_pdf)
        match_data = re.search(r'DATA EMISSÃO\s+(\d{2}/\d{2}/\d{4})', texto_pdf)
        match_fornecedor = re.search(r'Empresa Fornecedora:\s*(.+?)\s*CNPJ:', texto_pdf)
        match_cnpj = re.search(r'Empresa Fornecedora:.*?CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_pdf, re.DOTALL)
        #match_prazo = re.search(r'\d{2}/\d{2}/\d{4}\s*$', texto_pdf.split('\n')[texto_pdf.split('\n').index(next(l for l in texto_pdf.split('\n') if 'Até o dia' in l or re.search(r'\d{8}\s+\S', l)), '')] if any('Até o dia' in l or re.search(r'\d{8}\s+\S', l) for l in texto_pdf.split('\n')) else '', re.MULTILINE)
        match_prazo = next((re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l) for l in texto_pdf.split('\n') if re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l)), None)
        
        numero_oc = match_num.group(1) if match_num else ""
        data_emissao = match_data.group(1) if match_data else ""
        nome_fornecedor_pdf = match_fornecedor.group(1).strip() if match_fornecedor else ""
        cnpj_fornecedor_pdf = match_cnpj.group(1).strip() if match_cnpj else ""
        prazo_entrega_pdf = match_prazo.group(1).strip() if match_prazo else ""
        
        # if match_num:
        #     numero_oc = match_num.group(1)
        # if match_data:
        #     data_emissao = match_data.group(1)

    #return numero_oc, data_emissao
    return numero_oc, data_emissao, nome_fornecedor_pdf, cnpj_fornecedor_pdf, prazo_entrega_pdf


#%%

def extrai_dados (numchamado):
    sleep(1)
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
    
    janela_principal = driver.window_handles[0]
 
    xpaths_input = [
        '//*[@id="st-container"]/div/div/div/div[1]/ul[3]/div/div/div[1]/input',
        '//*[@id="st-container"]/div/div[1]/div/div[1]/ul[3]/div/div/div[1]/input',
        '//*[@id="st-container"]/div/div/div/div[1]/ul[3]/div/div/div[2]/input'
    ]
    
    inserir_compra = None

    for xpath_input in xpaths_input:
        try:
            inserir_compra = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath_input))
            )
            break  # encontrou, sai do loop
        except:
            continue
    
    if not inserir_compra:
        print(f"❌ Não foi possível localizar o campo de busca do chamado {numchamado}. Pulando.")
        return None
    
    inserir_compra.clear()
    sleep(1)
    inserir_compra.send_keys(str(numchamado))
    sleep(1)
    inserir_compra.send_keys(Keys.ENTER)
    
    print("Aguardando SE Suite...")
        
    try:
        primeiro_item = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="st-container"]/div/div/div/div[4]/div/div[2]/div/div/div[2]/div/div[2]/div[1]/span'))
        )
        print("Chamado localizado. Extraindo dados...")
    except TimeoutException:
        print("❌ Nenhum item encontrado para o chamado. Pulando.")
        return None
    
    for tentativa in range(5):
        handles_antes = set(driver.window_handles)
        try:
            primeiro_item.click()
            WebDriverWait(driver, 10).until(lambda d: len(set(d.window_handles) - handles_antes) > 0)
            nova_janela = list(set(driver.window_handles) - handles_antes)[0]
            driver.switch_to.window(nova_janela)
            break
        except:
            print("❌ Erro ao abrir nova janela para o chamado. Tentando novamente...")
            sleep(2)
    else:
        print("❌ Todas as tentativas falharam. Pulando chamado.")
        return None
    
    dados_dos_chamados = {}
    
    ### Tratar alerta está funcionando quando existe alerta, mas quando não tem
    ### alerta, está causando problemas. Robô não está clicando no "Solicitação
    ### de Aquisição". Investigar melhor.
    
    ### Update 22/02/2026 - após atualização da função, parece estar funcionando bem.
    
    sleep(1)   
    tratar_alerta(driver)
    
    try:
        titulo_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="headerTitle"]'))
        )
        titulo_completo = titulo_element.text.strip()
        titulo_limpo = titulo_completo.split(" - ", 1)[1] if " - " in titulo_completo else ""
    except TimeoutException:
        print("❌ Timeout ao tentar localizar o título do chamado. Pulando.")
        driver.close()
        driver.switch_to.window(janela_principal)
        return None
    
    # Status do chamado
    status_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="statusTextSpan"]'))
    )
    status_texto = status_element.text.strip()
            
    ### Troca para o frame
    try:
        WebDriverWait(driver, 50).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "ribbonFrame"))
        )
    except TimeoutException:
        print(f"❌ Timeout ao carregar frame 'ribbonFrame' no chamado {numchamado}. Pulando chamado.")
        driver.close()
        driver.switch_to.window(janela_principal)
        return None
    
    # Clica no botão "Solicitação de aquisição ISI"
    for tentativa in range(3):
        try:
            botao = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[text()="Solicitação de aquisição ISI"]/ancestor::a'))
            )
            botao.click()
            break  # Sucesso
        except:
            #print(f"Tentativa {tentativa+1}: botão não encontrado, tentando novamente...")
            sleep(2)    
        
    # Espera e entra no iframe
    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "frame_form_8a3449076f9f6db3016ff76aba7472f3"))
        )
    except TimeoutException:
        print("❌ Frame não carregou. Pulando chamado.")
        return None
    
    janela_chamado = driver.current_window_handle
    
    #Unidade
    unidade_map = {
    "INSTITUTO SENAI DE INOVAÇÃO EM MANUFATURA E LASER - 03774688005548 - 03774688000155": "ISI SM PL",
    "INSTITUTO SENAI DE INOVAÇÃO EM SISTEMAS EMBARCADOS - 03774688005467 - 03774688000155": "ISI SE",
    "INSTITUTO SENAI DE INOVAÃ‡ÃƒO EM MANUFATURA E LASER - 03774688005548 - 03774688000155": "ISI SM PL",
    "SENAI/SC - DIREÇÃO REGIONAL - 03774688000155 - 03774688000155": "DR SC",
    "INSTITUTO SENAI DE TECNOLOGIA EM LOGÍSTICA DE PRODUÇÃO - 03774688007320 - 03774688000155": "IST",
    "INSTITUTO SENAI DE TECNOLOGIA TÊXTIL, VESTUÁRIO E DESIGN - 03774688006439 - 03774688000155": "IST",
    "INSTITUTO SENAI DE TECNOLOGIA AMBIENTAL - 03774688006510 - 03774688000155": "IST",
    "INSTITUTO SENAI DE TECNOLOGIA EM EXCELÊNCIA OPERACIONAL - 03774688007320 - 03774688000155": "IST",
    "IST EM MOBILIDADE ELÉTRICA E ENERGIAS RENOVÁVEIS - 03774688008059 - 03774688000155": "IST",
    "IST ALIMENTOS E BEBIDAS E ISI SISTEMAS EMBARCADOS - 03774688007672 - 03774688000155": "IST",
    "SENAI/SC - BRUSQUE II - 03774688007400 - 03774688000155": "IST",
    }
    
    #Modalidade de Aquisição
    modalidade_map = {
    "bffd0ab8a3d83f081dfa79349ad4aa61": "ANS 13 Dias",
    "2e10d54dc4f9894e2b9a5917c4d0cd9c": "ANS 20 Dias",    
    "8f39e461ca98e4a2624e55564c613609": "ANS 30 Dias",
    "10418b61418c571e2c49ef89c5dcaf64": "ANS 4 Dias",
    "6b0f275b249de45d41617e1044e7b725": "ANS 8 Dias",
    "d2801b01f3eafc41709cbb42567ab8c0": "AQUISIÇÃO DIRETA",
    "548b6278c989e3fa6efa6c46dc292848": "AVALIAÇÃO COMPETITIVA (EMBRAPII)",
    "6c9c19595306f579a3bf2eb4d2bd9972": "COMPRA SIMPLIFICADA",
    "00f807948514d8310e6a84226f3f2e74": "CONTRATAÇÃO DIRETA (EMBRAPII)",
    "e7c5ed9c5b4e61aed21c74220a4442f4": "CREDENCIAMENTO",
    "1653d026b250b711bf6ee4edcdcf874f": "DISPENSA DE LICITAÇÃO",
    "a3782c54787727b5f76fdb1d5a660a8c": "INEXIGIBILIDADE",
    "9b30ed6f0e20484f466488710ac94370": "PREGÃO ELETRÔNICO",
    "5f1d346474fa5d3f224a230f610d1bf0": "PREGÃO PRESENCIAL",
    "aeed21038f19da74ba13c0c84bf72757": "SELEÇÃO PUBLICA (FUNDEP)",
    "e77f1a812ccb40258280b3b07db1d824": "SIMPLES COTAÇÃO (EMBRAPII)"
    }
        
    necessita_apoio_map = {
    "6841b637e9b4a208c3cd9a96a502fff3": "Não",
    "69257ea53984fcd08c85f7006b1c574b": "Sim"
    }
    
    necessita_contrato_map = {
    "6841b637e9b4a208c3cd9a96a502fff3": "Não",
    "69257ea53984fcd08c85f7006b1c574b": "Sim"
    }
    
    tipo_item_map = {
    "7014edacea7a45716f33e8085ee9a0ce": "Produto Internacional",
    "8119b542312bfdc90492e0f67b9d59a0": "Produto Nacional",
    "4ded5ec192e1ac6cde6e8ffbfa7e38f9": "Serviço Internacional",
    "165a0b9a7dc38b84e3ea3b220c316626": "Serviço Nacional"
    }
        
    #Campos a extrair
    campos = [
        ("Unidade", '//*[@id="nmwebservice_125f53af450b635b0544d2eb4d9ae6b8"]'),
        ("Data Aprovação GP", '//*[@id="field_8a3449076f9f6db3016fc927250c1163"]'),
        ("Identificador", '//*[@id="field_8a3449076f9f6db3016fc90ecee50d0f"]'),
        ("Nome Projeto", '//*[@id="nmwebservice_919e8ee72f4a21d3146166632058baff"]'),
        ("Fonte", '//*[@id="field_8a3449076f9f6db3016fc92a6763124c"]'),
        ("CR", '//*[@id="field_8a3449076f9f6db3016fd77250e735e0"]'),
        ("Projeto", '//*[@id="field_8a3449076f9f6db3016fd772bc7635f4"]'),
        ("Conta", '//*[@id="field_8a3449076f9f6db3016fd774d1863632"]'),
        ("Rubrica", '//*[@id="field_8a3449076f9f6db3016fc934596a145b"]'),
        ("Valor Inicial", '//*[@id="field_8a3449076f9f6db3016fc922d7cd109b"]'),
        ("Valor Final", '//*[@id="field_8a3449076f9f6db3016fc96466b81ca7"]'),
        ("Justificativa", '//*[@id="field_8a3449076f9f6db3016fc921c3a2107d"]'),
        ("Justificativa GP", '//*[@id="field_8a3449076f9f6db3016fc936726114cd"]'),
        ("Data Análise Célula", '//*[@id="field_8a3449076f9f6db3016fc93bb7e515bc"]'),
        ("Analista Inicial", '//*[@id="field_8a3449076f9f6db3016fc93b715515ae"]'),
        ("Analista Final", '//*[@id="field_8a3449076f9f6db3016fc953332119fd"]'),
        ("Modalidade", '//*[@id="oidzoom_8a3449076f9f6db3016ff872820c0ff2"]'),
        ("Apoio Consultivo", '//*[@id="oidzoom_8a3449076f9f6db3016ff871b2430fdf"]'),
        ("Necessita Contrato", '//*[@id="oidzoom_8a3449076f9f6db3016ff8720b910fe7"]'),
        ("Tipo Item", '//*[@id="oidzoom_8a3449076f9f6db3016ffb297b0f5c9b"]'),
        ("Processo Compra Finalizado", '//*[@id="field_8a3449076f9f6db3016fc95433971a26"]'),
        ("Data Aprovação Técnica", '//*[@id="field_8a3449076f9f6db3016fc9666f801d12"]'),
        ("Data Prevista Recebimento", '//*[@id="field_8a34490772473ce70172c30fab5e3842"]'),
        ("Data do Recebimento", '//*[@id="field_8a3449076f9f6db3016fd75554bd334c"]')
    ]
            
    for nome, xpath in campos:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        dados_dos_chamados[nome] = element.get_attribute("value")
    
    dados_dos_chamados["Descrição"] = titulo_limpo
    
    valor_final = dados_dos_chamados.get("Valor Final")
    valor_inicial = dados_dos_chamados.get("Valor Inicial")
    dados_dos_chamados["Valor R$"] = valor_final if valor_final else valor_inicial
        
    analista_final = dados_dos_chamados.get("Analista Final")
    analista_inicial = dados_dos_chamados.get("Analista Inicial")
    dados_dos_chamados["Analista"] = analista_final if analista_final else analista_inicial
        
    codigo_modalidade = dados_dos_chamados.get("Modalidade")
    dados_dos_chamados["Modalidade"] = modalidade_map.get(codigo_modalidade, codigo_modalidade)
       
    codigo_apoio_consultivo = dados_dos_chamados.get("Apoio Consultivo")
    dados_dos_chamados["Apoio Consultivo"] = necessita_apoio_map.get(codigo_apoio_consultivo, codigo_apoio_consultivo)
    
    codigo_contrato = dados_dos_chamados.get("Necessita Contrato")
    dados_dos_chamados["Necessita Contrato"] = necessita_contrato_map.get(codigo_contrato, codigo_contrato)
    
    codigo_tipo_item = dados_dos_chamados.get("Tipo Item")
    dados_dos_chamados["Tipo Item"] = tipo_item_map.get(codigo_tipo_item, codigo_tipo_item)
    
    codigo_unidade = dados_dos_chamados.get("Unidade")
    dados_dos_chamados["Código Unidade"] = unidade_map.get(codigo_unidade, codigo_unidade)


   # Extrair dados do PDF da 'Ordem de Compra'
    try:
        ordem_compra = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="link_filename8a3449076f9f6db3016fc987d8462468"]'))
        )
        texto = ordem_compra.text.strip()
        if texto.lower().endswith(".pdf"):
            texto = texto[:-4]
        dados_dos_chamados["Ordem de Compra"] = texto

        selenium_cookies = driver.get_cookies()
        session = requests.Session()
        for cookie in selenium_cookies:
            session.cookies.set(cookie['name'], cookie['value'])

        janela_chamado = driver.current_window_handle
        handles_antes = set(driver.window_handles)
        
        ordem_compra.click()
        WebDriverWait(driver, 10).until(lambda d: len(set(d.window_handles) - handles_antes) > 0)

        aba_pdf = list(set(driver.window_handles) - handles_antes)[0]
        driver.switch_to.window(janela_principal)
        driver.switch_to.window(aba_pdf)
        url_pdf = driver.current_url
        response = session.get(url_pdf)
        with pdfplumber.open(io.BytesIO(response.content)) as pdf:
            texto_pdf = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

        #print(f"📄 Texto extraído do PDF:\n{texto_pdf}")
        #numero_oc, data_emissao_oc_pdf = extrair_dados_oc(texto_pdf)
        numero_oc, data_emissao_oc_pdf, nome_fornecedor_pdf, cnpj_fornecedor_pdf, prazo_entrega_pdf = extrair_dados_oc(texto_pdf)
        #print(f"🔢 Número OC: {numero_oc} | 📅 Data Emissão: {data_emissao_oc_pdf}")

        dados_dos_chamados["numero_oc_pdf"] = numero_oc
        dados_dos_chamados["data_emissao_oc_pdf"] = data_emissao_oc_pdf
        dados_dos_chamados["nome_fornecedor_pdf"] = nome_fornecedor_pdf
        dados_dos_chamados["cnpj_fornecedor_pdf"] = cnpj_fornecedor_pdf
        dados_dos_chamados["prazo_entrega_pdf"] = prazo_entrega_pdf
        
        driver.close()

    except:
        dados_dos_chamados["Ordem de Compra"] = ""
   
    # Código Antigo extraindo apenas o texto do nome do arquivo:
    #    
    # try:
    #     ordem_compra = WebDriverWait(driver, 3).until(
    #         EC.presence_of_element_located((By.XPATH, '//*[@id="link_filename8a3449076f9f6db3016fc987d8462468"]'))
    #     )
    #     texto = ordem_compra.text.strip()
    #     if texto.lower().endswith(".pdf"):
    #         texto = texto[:-4]
    #     dados_dos_chamados["Ordem de Compra"] = texto
    # except:
    #     dados_dos_chamados["Ordem de Compra"] = "" 

    #sleep(1)

    ### HISTÓRICO
    # Volta para Ribbonframe para acessar histórico (ver se precisa dessa parte - Sim, precisa!)
    driver.switch_to.window(janela_chamado)
    driver.switch_to.default_content()
    
    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "ribbonFrame"))
        )
    except TimeoutException:
        return None
    
    # Clica no botão "Histórico"
    for tentativa in range(3):
        try:
            botao = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[text()="Histórico"]/ancestor::a'))
            )
            botao.click()
            break
        except:
            sleep(1) 
            
    #sleep(1)    

   
    try:
        WebDriverWait(driver, 15).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframe_history"))
        )
        #print("🎯 Entrou no iframe_history")
    except TimeoutException:
        print("❌ Não conseguiu acessar o iframe 'iframe_history'.")
        return None
    
    # Clica no botão "Exibir histórico completo"
    try:
        botao = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//*[starts-with(@id, "history")]/div/span/div/div/div/span[contains(text(), "Exibir histórico")]'))
        )
        botao.click()
        #print("✅ Botão 'Exibir histórico completo' clicado.")
    except TimeoutException:
        print("⚠️ Botão 'Exibir histórico completo' não clicável. Dados do histórico serão ignorados.")
        dados_dos_chamados["Data Emissão OC"] = ""
        dados_dos_chamados["Dias Suspenso"] = ""
        dados_dos_chamados["Status"] = status_texto
        worksheet_manuais.append_row([int(numchamado)])
        driver.close()
        driver.switch_to.window(janela_principal)
        return dados_dos_chamados
    
    # O código abaixo cancelava a extração quando não era possível achar o histórico completo.
    # No entanto, alguns modelos de histórico são diferentes, e até que se consiga extrair ambos,
    # optou-se por ignorar apenas os dados de emissão de OC e tempo suspenso para esses chamados.
    
    # except TimeoutException:
    #     print("❌ Botão 'Exibir histórico completo' não clicável. Pulando chamado...")
    #     return None
    
    # Aguarda conteúdo do histórico aparecer (3x)
    for tentativa in range(3):
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "timelineItem")))
            break
        except TimeoutException:
            print(f"⏳ Tentativa {tentativa+1}/3: 'timelineItem' não encontrado. Recarregando a página...")
            driver.refresh()
            sleep(2)
    else:
        print(f"❌ Falha ao localizar 'timelineItem' após 3 tentativas. Pulando chamado {numchamado}.")
        return None


    # Coleta HTML do iframe e analisa com BeautifulSoup
    html_history = driver.page_source
    soup = BeautifulSoup(html_history, 'html.parser')
    
    #atividade_habilitada = None
    data_emissao_oc = None
    data_atividade_prioritaria = None
    #atividade_prioritaria = ""
    periodos_suspensao = []
    data_inicio_suspensao = None
    dias_suspensos = 0
    data_oc_encontrada = False
    
    # Palavras-chave para atividades que devem sobrepor a regra padrão de data de emissão de OC
    gatilhos_prioritarios = [
        "encaminhar boleto pagamento dos serviços de courrier habilitada",
        "informar dados de pagamento habilitada",
        "abrir e aprovar rn para o serviço de importação habilitada",
        "executou a atividade formalizar contrato com a ação finalizar"
    ]
 
    # Percorre os blocos do histórico do mais recente para o mais antigo
    itens = soup.select("div.timelineItem")
    
    #itens = list(reversed(soup.select("div.timelineItem")))
    

    for idx, item in enumerate(itens):
        
        #Debug temporário
        #header = item.select_one("div.timelineItemContentHeader")
        #data_txt = header.get_text(strip=True) if header else None
        #print(f"[{idx}] {data_txt}")
        #/Debug temporário
        
        descricao_raw = item.select_one("div.timelineItemContent")
    
        if not descricao_raw:
            continue
    
        texto = descricao_raw.get_text(" ", strip=True)
        texto_normalizado = texto.replace("  ", " ").strip().lower()
    
        header = item.select_one("div.timelineItemContentHeader")
        data_txt = header.get_text(strip=True) if header else None
            
        # Regra 0: atividade prioritária (substitui Confirmar Recebimento para
        # fins de data de emissão da OC)
        for gatilho in gatilhos_prioritarios:
            if gatilho in texto_normalizado:

                if data_txt:
                    dt = data_hoje_ontem(data_txt)
                    if dt:
                        data_atividade_prioritaria = dt.strftime("%d/%m/%Y")
                break
    
        # Regra 1: confirmar recebimento habilitada
        if not data_oc_encontrada:
            if "confirmar recebimento" in texto_normalizado and "habilitada" in texto_normalizado:
                if data_txt:
                    dt = data_hoje_ontem(data_txt)
                    if dt:
                        data_emissao_oc = dt.strftime("%d/%m/%Y")
                        data_oc_encontrada = True
    
        # Regra 2: cancelamento
        if not data_oc_encontrada:
            if (
                "executou a atividade solicitar aquisição com a ação cancelar" in texto_normalizado
                or "atividade solicitar aquisição executada automaticamente com a ação finalizador" in texto_normalizado
            ):
                status_texto = "Cancelado"
                data_oc_encontrada = True  # Para de buscar a data da OC, mas continua o loop
    
        # Suspensão — roda sempre, independente das regras acima
        if "suspendeu a instância" in texto_normalizado:
            if data_txt:
                dt = data_hoje_ontem(data_txt)
                if dt:
                    data_inicio_suspensao = dt            
    
        elif "reativou a instância" in texto_normalizado and data_inicio_suspensao:
            if data_txt:
                dt = data_hoje_ontem(data_txt)
                if dt:
                    dias = (data_inicio_suspensao - dt).days
                    if dias > 0:
                        periodos_suspensao.append(dias)
                    data_inicio_suspensao = None
    
    dias_suspensos = sum(periodos_suspensao) if periodos_suspensao else 0


    # Prioridade sobre regra padrão apenas se não houver outra data definida
    if data_atividade_prioritaria:
        data_emissao_oc = data_atividade_prioritaria

    # Aqui poderemos adicionar mais elif com novas regras de verificação futuras
    # elif "outra condição" in texto_normalizado:
    #     fazer algo

    # Ao final do bloco, `data_conclusao_processo` ou `atividade_habilitada` estarão preenchidos

    #####################################

    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()

    driver.switch_to.window(janela_principal)
    
    dados_dos_chamados["Status"] = status_texto
    dados_dos_chamados["Data Emissão OC"] = data_emissao_oc
    dados_dos_chamados["Dias Suspenso"] = dias_suspensos

    print("Dados do chamado ", numchamado, " extraídos.")
    
    return dados_dos_chamados

#%%

def extrai_dados_com_retry(numchamado, tentativas=2, espera=20):
    janela_principal = driver.window_handles[0]
    for tentativa in range(1, tentativas + 1):
        try:
            return extrai_dados(numchamado)
        except TimeoutException:
            # Fecha todas as janelas extras antes de tentar de novo
            try:
                for janela in driver.window_handles:
                    if janela != janela_principal:
                        driver.switch_to.window(janela)
                        driver.close()
                driver.switch_to.window(janela_principal)
            except:
                pass
                        
            if tentativa < tentativas:
                print(f"⏳ Timeout no chamado {numchamado}. Tentativa {tentativa}/{tentativas}. Aguardando {espera}s...")
                sleep(espera)
            else:
                print(f"❌ Chamado {numchamado} falhou após {tentativas} tentativas. Pulando.")
                return None

#%% Google Sheets

def adicionar_gsheet():
    print(f"Adicionando dados à planilha {spreadsheet.title}...")

    # Limpa quebras de linha e espaços em todas as colunas
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.replace('\n', ' ').strip() if isinstance(x, str) else x
        )

    # Remove os pontos apenas da coluna "Valor R$"
    if "Valor R$" in df.columns:
        df["Valor R$"] = df["Valor R$"].apply(
            lambda x: x.replace('.', '') if isinstance(x, str) else x
        )

    dados = df.values.tolist()

    linha_final = len(worksheet.get_all_values()) + 1
    worksheet.insert_rows(dados, row=linha_final)

    print(f"Dados adicionados à aba '{worksheet.title}' da planilha {spreadsheet.title}.")
    print("Finalizando...")

#adicionar_gsheet()

#%%

todos_os_dados = []

cabecalhos_esperados = ["Código Unidade", "Unidade", "Data Aprovação GP", "Identificador",
                        "Atividade Habilitada", "Nome Projeto", "Apelido Projeto",
                        "Descrição", "Fonte", "CR", "Projeto", "Conta", "Rubrica",
                        "Valor R$", "Justificativa", "Justificativa GP", "Data Análise Célula",
                        "Analista", "Modalidade", "Apoio Consultivo", "Necessita Contrato",
                        "Tipo Item", "ANS", "Processo Compra Finalizado", "Data Aprovação Técnica",
                        "Ordem de Compra", "Data Prevista Recebimento", "Data Emissão OC",
                        "Dias Suspenso", "Data do Recebimento", "numero_oc_pdf",
                        "data_emissao_oc_pdf", "nome_fornecedor_pdf", "cnpj_fornecedor_pdf",
                        "prazo_entrega_pdf"]

valores_existentes = worksheet.get_all_records(expected_headers=cabecalhos_esperados)

linhas_existentes = worksheet.get_all_values()

mapa_identificador_linha = {
    str(linha[3]).zfill(6): idx + 1
    for idx, linha in enumerate(linhas_existentes[1:])
    if len(linha) > 3 and str(linha[3]).strip().replace('.', '', 1).isdigit()
}

hoje = datetime.now().strftime("%d/%m/%Y")

# Pares ja processados é montado novamente depois da extração dos manuais.
# Avaliar se pode ser retirado daqui futuramente. Mantido por questões de segurança
# para evitar extração em duplicidade.
pares_ja_processados = {
    (str(linha["Identificador"]).zfill(6), linha["Atividade Habilitada"]) for linha in valores_existentes
}

#%% Função Registrar Chamados na Planilha

def registrar_chamado(dados_dos_chamados, atividade, descricao, identificador, hoje, remover_manual=False):
    for col in ["Justificativa", "Justificativa GP"]:
        if isinstance(dados_dos_chamados.get(col), str):
            dados_dos_chamados[col] = dados_dos_chamados[col].replace('\n', ' ').strip()

    if isinstance(descricao, str):
        descricao = descricao.replace('\n', ' ').strip()
    dados_dos_chamados["Descrição"] = descricao

    dados_dos_chamados["Atividade Habilitada"] = atividade
    
    #Inserir apelido do projeto
    codigo_projeto_extraido = dados_dos_chamados.get("Projeto", "")
    dados_dos_chamados["Apelido Projeto"] = obter_apelido_projeto(codigo_projeto_extraido)


    if isinstance(dados_dos_chamados.get("Valor R$"), str):
        dados_dos_chamados["Valor R$"] = dados_dos_chamados["Valor R$"].replace('.', '')

    #dados_dos_chamados["Data Atualização"] = hoje

    # Cálculo do ANS com 3 níveis de prioridade:
    # 1. Se existir na aba ANS
    # 2. Se modalidade tiver dias no nome
    # 3. Regra padrão por Apoio + Tipo

    identificador_zfill = str(identificador).zfill(6)

    if identificador_zfill in mapa_ans_custom:
        try:
            ans_valor = float(mapa_ans_custom[identificador_zfill])
            if ans_valor.is_integer():
                ans_valor = int(ans_valor)
            dados_dos_chamados["ANS"] = ans_valor
        except:
            dados_dos_chamados["ANS"] = mapa_ans_custom[identificador_zfill]
    else:
        modalidade = dados_dos_chamados.get("Modalidade", "")
        ans_por_modalidade = {
            "ANS 13 Dias": 13,
            "ANS 20 Dias": 20,
            "ANS 30 Dias": 30,
            "ANS 4 Dias": 4,
            "ANS 8 Dias": 8
        }

        if modalidade in ans_por_modalidade:
            dados_dos_chamados["ANS"] = ans_por_modalidade[modalidade]
        else:
            apoio = dados_dos_chamados.get("Apoio Consultivo")
            tipo = dados_dos_chamados.get("Tipo Item")
            ans_map = {
                ("Não", "Produto Internacional"): 6,
                ("Não", "Produto Nacional"): 5,
                ("Não", "Serviço Internacional"): 6,
                ("Não", "Serviço Nacional"): 5,
                ("Sim", "Produto Internacional"): 30,
                ("Sim", "Produto Nacional"): 20,
                ("Sim", "Serviço Internacional"): 30,
                ("Sim", "Serviço Nacional"): 20
            }
            dados_dos_chamados["ANS"] = ans_map.get((apoio, tipo), "")

    status_final = dados_dos_chamados.get("Status", "")
    if status_final == "Cancelado" and dados_dos_chamados.get("ANS") == "":
        dados_dos_chamados["ANS"] = "Cancelado"
    elif dados_dos_chamados.get("ANS") == "":
        dados_dos_chamados["ANS"] = "Em análise"

    linha_ordenada = [dados_dos_chamados.get(col, "") for col in cabecalhos_esperados]
    linha_existente = mapa_identificador_linha.get(identificador)

    if linha_existente:
        worksheet.update(values=[linha_ordenada], range_name=f"A{linha_existente+1}")
        print(f"🔁 Chamado {identificador} atualizado na linha {linha_existente+1}.")
    else:
        worksheet.append_row(linha_ordenada)
        print(f"➕ Chamado {identificador} adicionado ao final da planilha.")

    if remover_manual:
        remover_chamado_manuais(worksheet_manuais, identificador)


#%% Primeiro: processa chamados manuais (forçar extração)

print("📌 Iniciando extração de chamados manuais...")
chamados_extraidos_com_sucesso = []

#Verifica dados a extrair manualmente e dados a ignorar
for idx, numero in enumerate(lista_manuais):
    numero_formatado = str(int(float(numero))).zfill(6)
    
    if numero_formatado in lista_ignorar:
        print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Chamado {numero_formatado} está na lista de ignorados. Pulando e removendo da lista manual.")
        remover_chamado_manuais(worksheet_manuais, numero_formatado)
        continue

    # Esta parte havia sido incluída para apoiar na inserção de muitos dados na planilha
    # Pulando chamados que já haviam sido finalizados. 
    #
    #atividades_para_ignorar = ["Encerrado", "Cancelado", "Confirmar recebimento  do item solicitado"]
    
    #if any((numero_formatado, atividade) in pares_ja_processados for atividade in atividades_para_ignorar):
    #    print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Chamado {numero_formatado} já encerrado ou em etapa final. Pulando extração.")
    #    remover_chamado_manuais(worksheet_manuais, numero_formatado)
    #    continue

        
    print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Acessando chamado {numero_formatado}")
    dados_dos_chamados = extrai_dados_com_retry(numero_formatado)
    
    # ⛔ Se a extração falhou, pula e mantém o chamado na lista
    if dados_dos_chamados is None:
        print(f"⚠️ Chamado {numero_formatado} não pôde ser extraído. Mantendo na lista manual.")
        continue
    
    atividade = df.loc[
        df["Identificador"].apply(lambda x: str(int(float(x))).zfill(6)) == numero_formatado,
        "AtividadeHabilitadaFiltrada"
    ]
    atividade_df = atividade.values[0] if not atividade.empty else "Indefinida"
    
    status_texto = dados_dos_chamados.get("Status", "")
    atividade_habilitada = status_texto if status_texto in ["Encerrado", "Cancelado", "Suspenso"] else atividade_df

    if dados_dos_chamados:
        registrar_chamado(
            dados_dos_chamados,
            atividade=atividade_habilitada,
            descricao=dados_dos_chamados.get("Descrição", ""),  # já veio do SE Suite
            identificador=numero_formatado,
            hoje=hoje,
            remover_manual=True
        )

print("✅ Encerrada a extração de chamados manuais. Continuando para os demais chamados...")

valores_existentes = worksheet.get_all_records(expected_headers=cabecalhos_esperados)

pares_ja_processados = {
    (str(linha["Identificador"]).zfill(6), linha["Atividade Habilitada"]) for linha in valores_existentes
}


#%% Segue com os chamados automáticos

print("📌 Iniciando extração dos chamados em aberto...")

for idx, numero in enumerate(num_chamados):
    #if (str(numero), atividadehabilitada[idx]) in pares_ja_processados:
    identificador_zfill = str(numero).zfill(6)
    
    if identificador_zfill in lista_ignorar:
        print(f"[{idx+1}/{len(num_chamados)}] Chamado {identificador_zfill} está na lista de ignorados. Pulando.")
        continue
    
    if (identificador_zfill, atividadehabilitada[idx]) in pares_ja_processados:
        #print(f"[{idx+1}/{len(num_chamados)}] Chamado {numero} sem alteração de status. Pulando.")
        continue

    print(f"[{idx+1}/{len(num_chamados)}] Acessando chamado {identificador_zfill}")
    dados_dos_chamados = extrai_dados_com_retry(numero)

    if dados_dos_chamados:
        registrar_chamado(
            dados_dos_chamados,
            atividade=atividadehabilitada[idx],
            descricao=objetos_compra[idx],
            identificador=str(numero),
            hoje=hoje,
            remover_manual=False
        )

print("✅ Encerrada a extração dos chamados em aberto.")

#%% Verifica chamados que saíram do XLS, mas não foram encerrados nem cancelados

print("🔎 Atualizando chamados encerrados recentemente...")

# 1. Conjunto de chamados atuais no XLS
conjunto_chamados_xls = set(num_chamados)

# 2. Chamados da planilha com status diferente de "Encerrado" ou "Cancelado"
chamados_para_verificar = [
    linha for linha in valores_existentes
    if linha.get("Atividade Habilitada") not in ["Encerrado", "Cancelado"]
]

total_que_saiu = len(chamados_para_verificar) - len(conjunto_chamados_xls)

# 3. Para cada chamado da planilha, verifica se ele saiu do XLS
for idx, linha in enumerate(chamados_para_verificar):
    identificador = str(linha["Identificador"]).zfill(6)
    if identificador not in conjunto_chamados_xls:
        print(f"[{idx+1}/{total_que_saiu}] 🔁 Chamado {identificador} saiu do XLS. Extraindo novamente...")
        dados_dos_chamados = extrai_dados_com_retry(identificador)

        if dados_dos_chamados:
            status_texto = dados_dos_chamados.get("Status", "")
            atividade_atualizada = status_texto if status_texto in ["Encerrado", "Cancelado", "Suspenso"] else linha.get("Atividade Habilitada", "Indefinida")
            descricao_existente = linha.get("Descrição", "")

            registrar_chamado(
                dados_dos_chamados,
                atividade=atividade_atualizada,
                descricao=descricao_existente,
                identificador=identificador,
                hoje=hoje,
                remover_manual=False
            )

print("✅ Encerrada a extração dos chamados finalizados recentemente.")

#%% Exportar a aba "Dados_v1_1" para CSV

print("💾 Exportando aba 'Dados_v1_1' para CSV...")

dados_worksheet = worksheet.get_all_values()

caminho_csv = r"C:\RPA\se_suite_xls\Dados_v1_1.csv"
with open(caminho_csv, mode='w', newline='', encoding='utf-8') as arquivo_csv:
    writer = csv.writer(arquivo_csv)
    writer.writerows(dados_worksheet)

print(f"✅ Arquivo CSV exportado com sucesso para {caminho_csv}")

driver.quit()

print("Finalizando...")

sleep(3)

#%% DEBUGS úteis para identificar os frames do SE Suite


    ### DEBUG: Salvar HTML
    # with open("html_debug.html", "w", encoding="utf-8") as f:
    #     f.write(driver.page_source)

    
    # #DEBUG: Salvar cada iframe como HTML para inspeção.
    # driver.switch_to.default_content()
    # iframes = driver.find_elements(By.TAG_NAME, "iframe")
    
    # for i, iframe in enumerate(iframes):
    #     nome = iframe.get_attribute("name") or iframe.get_attribute("id") or f"iframe_{i}"
    #     nome = nome.replace("/", "_").replace("\\", "_")
    
    #     try:
    #         driver.switch_to.frame(iframe)
    #         html = driver.page_source
    #         with open(f"{nome}.html", "w", encoding="utf-8") as f:
    #             f.write(html)
    #         print(f"✅ {nome}.html salvo com sucesso.")
    
    #         # Tenta capturar iframes internos
    #         iframes_internos = driver.find_elements(By.TAG_NAME, "iframe")
    #         for j, iframe_interno in enumerate(iframes_internos):
    #             nome_interno = iframe_interno.get_attribute("name") or iframe_interno.get_attribute("id") or f"{nome}_interno_{j}"
    #             nome_interno = nome_interno.replace("/", "_").replace("\\", "_")
    #             try:
    #                 driver.switch_to.frame(iframe_interno)
    #                 html_interno = driver.page_source
    #                 with open(f"{nome_interno}.html", "w", encoding="utf-8") as f:
    #                     f.write(html_interno)
    #                 print(f"✅ {nome_interno}.html salvo com sucesso.")
    #                 driver.switch_to.parent_frame()
    #             except Exception as e:
    #                 print(f"❌ Erro ao salvar iframe interno {nome_interno}: {e}")
    #         driver.switch_to.parent_frame()
    #     except Exception as e:
    #         print(f"❌ Erro ao salvar iframe {nome}: {e}")
    #         driver.switch_to.default_content()
    # #/DEBUG: Salvar cada iframe como HTML para inspeção.
    
    # DEBUG
    # Exporta HTML completo do iframe de histórico
    # try:
    #     html_history = driver.page_source
    #     with open("iframe_history_renderizado.html", "w", encoding="utf-8") as f:
    #         f.write(html_history)
    #     print("📄 HTML do histórico salvo como 'iframe_history_renderizado.html'.")
    # except Exception as e:
    #     print(f"❌ Erro ao salvar HTML do histórico: {e}")


    # ########## DEBUG 1
    # #driver.switch_to.default_content()
    
    # iframes = driver.find_elements(By.TAG_NAME, "iframe")
    # print("DEBUG 1")
    # print(f"🔎 {len(iframes)} iframe(s) no debug 1:")
    # for i, iframe in enumerate(iframes):
    #     print(f"[{i}] name={iframe.get_attribute('name')} | id={iframe.get_attribute('id')}")
    # ######### /DEBUG1
    
    
    
    # print("\n🔎 PREVIEW DOS IFRAMES (1000 primeiros caracteres)\n")
    
    # for i, iframe in enumerate(iframes):
    #     nome = iframe.get_attribute("name")
    #     iframe_id = iframe.get_attribute("id")
    
    #     try:
    #         driver.switch_to.frame(iframe)
    #         html = driver.page_source
    #         preview = html.replace("\n", " ").replace("\r", " ")[:1000]
    #         print(f"[{i}] name={nome} | id={iframe_id}")
    #         print(f"    {preview}\n")
    #     except Exception as e:
    #         print(f"[{i}] name={nome} | id={iframe_id}")
    #         print("    ❌ Erro ao acessar iframe\n")
    #     finally:
    #         driver.switch_to.parent_frame()    
    

    # ########## DEBUG 2
    # #driver.switch_to.default_content()
    
    # iframes = driver.find_elements(By.TAG_NAME, "iframe")
    # print("DEBUG 2")
    # print(f"🔎 {len(iframes)} iframe(s) no debug 1:")
    # for i, iframe in enumerate(iframes):
    #     print(f"[{i}] name={iframe.get_attribute('name')} | id={iframe.get_attribute('id')}")
    # ########## /DEBUG2