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
from selenium.webdriver.common.keys import Keys
#from selenium.common.exceptions import TimeoutException
#from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from datetime import datetime
import os
import ctypes
import win32com.client as win32
import gspread

#%%

options = Options()
#options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")

options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\RPA\se_suite_xls",
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = Chrome(options=options)

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

WebDriverWait(driver, 30).until(lambda d: d.execute_script('return document.readyState') == 'complete')

sleep(1)

driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=tracking,104,2')

WebDriverWait(driver, 30).until(lambda d: d.execute_script('return document.readyState') == 'complete')

sleep(1)

WebDriverWait(driver, 10).until(
    EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe"))
)

sleep(1)

# botão seta
botao_seta = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "se_admin_btnreport-menuButton"))
)
botao_seta.click()

# Aguarda e clica em "Exportar para Excel"
botao_exportar = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Exportar para Excel")]'))
)
botao_exportar.click()

print("Baixando arquivo XLS...")

sleep(1)

#%%

caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"
inicio = time.time()
timeout = 600  # até 10 minutos

while time.time() - inicio < timeout:
    if os.path.exists(caminho) and not os.path.exists(caminho + ".crdownload"):
        print("Convertendo arquivo para XLSX...")
        break
    time.sleep(2)
else:
    raise TimeoutError("Download não terminou dentro do tempo esperado.")

for janela in driver.window_handles:
    if janela != janela_principal:
        driver.switch_to.window(janela)
        driver.close()

# Garante que está de volta na janela principal
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

caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"

desbloquear_arquivo_excel(caminho)

#%%

def converter_xls_para_xlsx(caminho_xls, caminho_xlsx):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # Excel rodando "invisível"
    excel.DisplayAlerts = False  # Evita pop-ups e confirmações

    wb = excel.Workbooks.Open(caminho_xls)
    wb.SaveAs(caminho_xlsx, FileFormat=51)  # 51 = .xlsx
    wb.Close(False)
    excel.Quit()

converter_xls_para_xlsx(
    caminho,
    r"C:\RPA\se_suite_xls\relatorio_convertido.xlsx"
)

if os.path.exists(caminho):
    os.remove(caminho)
    print("Arquivo original excluído.")
else:
    print("Arquivo original não encontrado para exclusão.")


#%%

df = pd.read_excel(r"C:\RPA\se_suite_xls\relatorio_convertido.xlsx")
#print(df.head())

#%%

df = df.iloc[1:].reset_index(drop=True)
df = df.drop(columns=['P', 'S', 'SW', 'SLA', 'PR', 'D', 'A', 'Executor', 'Processo', 'Tipo de workflow'])
df = df[~df["Atividade habilitada"].str.startswith("Confirmar recebimento  do item solicitado", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Analisar pertinência da solicitação", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Tomar ciência da negativa da solicitação", na=False)]
df["AtividadeHabilitadaFiltrada"] = df["Atividade habilitada"].str.split("(", n=1).str[0].str.strip()
print(df["AtividadeHabilitadaFiltrada"].value_counts())
print()
print('Total de chamados: {}' .format(df["Identificador"].count()))


#%%

num_chamados = df["Identificador"].astype(int).tolist()
objetos_compra = df["Título"].tolist()
atividadehabilitada = df["AtividadeHabilitadaFiltrada"].tolist()

#%%

def extrai_dados (numchamado):
    sleep(1)
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
    
    janela_principal = driver.window_handles[0]
    
    xpaths_input = [
        '//*[@id="st-container"]/div/div[1]/div/div[1]/ul[3]/div/div/div[1]/input',
        '//*[@id="st-container"]/div/div/div/div[1]/ul[3]/div/div/div[2]/input'
    ]
    
    for xpath_input in xpaths_input:
        try:
            inserir_compra = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath_input))
            )
            break
        except:
            continue
    
    inserir_compra.clear()
    sleep(0.5)
    inserir_compra.send_keys(str(numchamado))
    sleep(0.5)
    inserir_compra.send_keys(Keys.ENTER)
    
    print("Aguardando SE Suite...")
    
    primeiro_item = WebDriverWait(driver, 600).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="st-container"]/div/div/div/div[4]/div/div[2]/div/div/div[2]/div/div[2]/div[1]/span'))
    )
    print("Chamado localizado. Extraindo dados...")
    #primeiro_item.click()
           
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
        
    # Troca para o frame
    WebDriverWait(driver, 50).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "ribbonFrame"))
    )
    
    # Clica no botão "Solicitação de aquisição ISI"
    for tentativa in range(5):
        try:
            botao = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[text()="Solicitação de aquisição ISI"]/ancestor::a'))
            )
            botao.click()
            break  # Sucesso
        except:
            print(f"Tentativa {tentativa+1}: botão não encontrado, tentando novamente...")
            sleep(2)    
        
    # Espera e entra no iframe
    WebDriverWait(driver, 100).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "frame_form_8a3449076f9f6db3016ff76aba7472f3"))
    )
    
    modalidade_map = {
    "d2801b01f3eafc41709cbb42567ab8c0": "AQUISIÇÃO DIRETA",
    "548b6278c989e3fa6efa6c46dc292848": "AVALIAÇÃO COMPETITIVA (EMBRAPII)",
    "00f807948514d8310e6a84226f3f2e74": "CONTRATAÇÃO DIRETA (EMBRAPII)",
    "1653d026b250b711bf6ee4edcdcf874f": "DISPENSA DE LICITAÇÃO",
    "e77f1a812ccb40258280b3b07db1d824": "SIMPLES COTAÇÃO (EMBRAPII)",
    "6c9c19595306f579a3bf2eb4d2bd9972": "COMPRA SIMPLIFICADA",
    "a3782c54787727b5f76fdb1d5a660a8c": "INEXIGIBILIDADE"
    }
    
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
        ("Valor R$", '//*[@id="field_8a3449076f9f6db3016fc922d7cd109b"]'),
        ("Justificativa", '//*[@id="field_8a3449076f9f6db3016fc921c3a2107d"]'),
        ("Justificativa GP", '//*[@id="field_8a3449076f9f6db3016fc936726114cd"]'),
        ("Data Análise Célula", '//*[@id="field_8a3449076f9f6db3016fc93bb7e515bc"]'),
        ("Modalidade", '//*[@id="oidzoom_8a3449076f9f6db3016ff872820c0ff2"]'),
        ("Processo Compra Finalizado", '//*[@id="field_8a3449076f9f6db3016fc95433971a26"]'),
        ("Data Aprovação Técnica", '//*[@id="field_8a3449076f9f6db3016fc9666f801d12"]'),
        ("Data Prevista Recebimento", '//*[@id="field_8a34490772473ce70172c30fab5e3842"]'),
    ]
    
    print("Dados do chamado ", numchamado, " extraídos.")
    
    dados_dos_chamados = {}
            
    for nome, xpath in campos:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        dados_dos_chamados[nome] = element.get_attribute("value")
    
    codigo_modalidade = dados_dos_chamados.get("Modalidade")
    dados_dos_chamados["Modalidade"] = modalidade_map.get(codigo_modalidade, codigo_modalidade)
    
    #driver.close()
    #sleep(1)
    #driver.switch_to.window(driver.window_handles[0])
     
    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()

    driver.switch_to.window(janela_principal)
    
    return dados_dos_chamados

#%% Google Sheets

gc = gspread.service_account(filename=os.path.join(os.path.dirname(os.getcwd()), 'crested-century-386316-01c90985d6e4.json'))

spreadsheet = gc.open("Acompanhamento_Aquisições_Teste")
worksheet = spreadsheet.worksheet("Dados")

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

cabecalhos_esperados = ["Unidade", "Data Aprovação GP", "Identificador", "Atividade Habilitada",
                        "Nome Projeto", "Descrição", "Fonte", "CR", "Projeto",
                        "Conta", "Rubrica", "Valor R$", "Justificativa", "Justificativa GP",
                        "Data Análise Célula", "Modalidade", "Processo Compra Finalizado",
                        "Data Aprovação Técnica", "Data Prevista Recebimento",
                        "Data Atualização"]
valores_existentes = worksheet.get_all_records(expected_headers=cabecalhos_esperados)

linhas_existentes = worksheet.get_all_values()
mapa_identificador_linha = {
    str(linha[2]): idx + 1  # col 2 = "Identificador", +1 porque gspread começa em 1
    for idx, linha in enumerate(linhas_existentes[1:])  # pula cabeçalho
    if len(linha) > 2  # ignora linhas incompletas
}

hoje = datetime.now().strftime("%d/%m/%Y")

# ids_ja_processados = {
#     str(linha["Identificador"]) for linha in valores_existentes if linha.get("Data Atualização") == hoje
# }

pares_ja_processados = {
    (str(linha["Identificador"]), linha["Atividade Habilitada"]) for linha in valores_existentes
}

for idx, numero in enumerate(num_chamados):
    #if str(numero) in ids_ja_processados:
    if (str(numero), atividadehabilitada[idx]) in pares_ja_processados:
        print(f"[{idx+1}/{len(num_chamados)}] Chamado {numero} sem alteração de status. Pulando.")
        continue

    print(f"[{idx+1}/{len(num_chamados)}] Acessando chamado {numero}")
    dados_dos_chamados = extrai_dados(numero)

    if dados_dos_chamados:
        colunas_para_limpar = ["Justificativa", "Justificativa GP"]
        for col in colunas_para_limpar:
            if isinstance(dados_dos_chamados.get(col), str):
                dados_dos_chamados[col] = dados_dos_chamados[col].replace('\n', ' ').strip()
        
        descricao = objetos_compra[idx]
        if isinstance(descricao, str):
            descricao = descricao.replace('\n', ' ').strip()
        dados_dos_chamados["Descrição"] = descricao   
        
        dados_dos_chamados["Atividade Habilitada"] = atividadehabilitada[idx]
        
        if isinstance(dados_dos_chamados.get("Valor R$"), str):
            dados_dos_chamados["Valor R$"] = dados_dos_chamados["Valor R$"].replace('.', '')
            
        dados_dos_chamados["Data Atualização"] = hoje
        
        linha_ordenada = [dados_dos_chamados.get(col, "") for col in cabecalhos_esperados]

        #worksheet.append_row(linha_ordenada)
        
        identificador_str = str(numero)
        linha_existente = mapa_identificador_linha.get(identificador_str)
        
        if linha_existente:
            worksheet.update(values=[linha_ordenada], range_name=f"A{linha_existente+1}")
            print(f"🔁 Chamado {identificador_str} atualizado na linha {linha_existente+1}.")
        else:
            worksheet.append_row(linha_ordenada)
            print(f"➕ Chamado {identificador_str} adicionado ao final da planilha.")
        
        #dados_dos_chamados["Descrição"] = objetos_compra[idx]
        #colunas = df.columns.tolist()
        #colunas.remove("Descrição")
        #colunas.insert(4, "Descrição")  # posição 2 = terceira coluna (0-based)
        #dados_dos_chamados = dados_dos_chamados[colunas]
        #dados_dos_chamados["Data Atualização"] = hoje

        #worksheet.append_row(list(dados_dos_chamados.values()))

#worksheet.format("L2:L", {"numberFormat": {"type": "CURRENCY"}})

#driver.quit()

#df = pd.DataFrame(todos_os_dados)

#df["Descrição"] = objetos_compra

# Reorganiza a ordem das colunas
#colunas = df.columns.tolist()
#colunas.remove("Descrição")
#colunas.insert(4, "Descrição")  # posição 2 = terceira coluna (0-based)

#df = df[colunas]

#print(df)
















