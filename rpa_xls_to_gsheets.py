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
from selenium.common.exceptions import TimeoutException
from datetime import datetime
import os
import ctypes
import win32com.client as win32
import gspread

#%%

options = Options()
options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")

options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\RPA\se_suite_xls",
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = Chrome(options=options)

#Dados Google Sheets
gc = gspread.service_account(filename=os.path.join(os.path.dirname(os.getcwd()), 'crested-century-386316-01c90985d6e4.json'))

spreadsheet = gc.open("Acompanhamento_Aquisições_RPA")
worksheet = spreadsheet.worksheet("Dados")


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

WebDriverWait(driver, 100).until(lambda d: d.execute_script('return document.readyState') == 'complete')

sleep(2)

driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=tracking,104,2')

WebDriverWait(driver, 100).until(lambda d: d.execute_script('return document.readyState') == 'complete')

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

print("Baixando arquivo XLS...")

sleep(1)

#%%

caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"
inicio = time.time()
timeout = 600

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

caminho = r"C:\RPA\se_suite_xls\Gestão de workflow.xls"

desbloquear_arquivo_excel(caminho)

#%%

def converter_xls_para_xlsx(caminho_xls, caminho_xlsx):
    excel = win32.DispatchEx('Excel.Application')  # cria nova instância
    #excel = win32.gencache.EnsureDispatch('Excel.Application')
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

#Lê arquivo baixado do SE Suite
df = pd.read_excel(r"C:\RPA\se_suite_xls\relatorio_convertido.xlsx")

# Acessa as abas "Manuais" e "Ignorar" da mesma planilha
worksheet_manuais = spreadsheet.worksheet("Manuais")
worksheet_ignorar = spreadsheet.worksheet("Ignorar")

# Lê os valores da coluna A (sem cabeçalho)
valores_manuais = worksheet_manuais.col_values(1)
valores_ignorar = worksheet_ignorar.col_values(1)

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
#df = df[~df["Atividade habilitada"].str.startswith("Confirmar recebimento  do item solicitado", na=False)]
#df = df[~df["Atividade habilitada"].str.startswith("Analisar pertinência da solicitação", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Solicitar aquisição", na=False)]
df = df[~df["Atividade habilitada"].str.startswith("Tomar ciência da negativa da solicitação", na=False)]
df["AtividadeHabilitadaFiltrada"] = df["Atividade habilitada"].str.split("(", n=1).str[0].str.strip()
print(df["AtividadeHabilitadaFiltrada"].value_counts())
print()
print('Total de chamados: {}' .format(df["Identificador"].count()))


#%%

num_chamados = df["Identificador"].apply(lambda x: str(int(float(x))).zfill(6) if pd.notnull(x) else "").tolist()
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
        
    try:
        primeiro_item = WebDriverWait(driver, 200).until(
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
    
    titulo_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="headerTitle"]'))
    )
    titulo_completo = titulo_element.text.strip()
    titulo_limpo = titulo_completo.split(" - ", 1)[1] if " - " in titulo_completo else ""
    
    # Status do chamado
    status_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="statusTextSpan"]'))
    )
    status_texto = status_element.text.strip()
            
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
    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "frame_form_8a3449076f9f6db3016ff76aba7472f3"))
        )
    except TimeoutException:
        print("❌ Frame não carregou. Pulando chamado.")
        return None
    
    #Modalidade de Aquisição
    modalidade_map = {
    "d2801b01f3eafc41709cbb42567ab8c0": "AQUISIÇÃO DIRETA",
    "548b6278c989e3fa6efa6c46dc292848": "AVALIAÇÃO COMPETITIVA (EMBRAPII)",
    "00f807948514d8310e6a84226f3f2e74": "CONTRATAÇÃO DIRETA (EMBRAPII)",
    "1653d026b250b711bf6ee4edcdcf874f": "DISPENSA DE LICITAÇÃO",
    "e77f1a812ccb40258280b3b07db1d824": "SIMPLES COTAÇÃO (EMBRAPII)",
    "6c9c19595306f579a3bf2eb4d2bd9972": "COMPRA SIMPLIFICADA",
    "a3782c54787727b5f76fdb1d5a660a8c": "INEXIGIBILIDADE"
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
        ("Modalidade", '//*[@id="oidzoom_8a3449076f9f6db3016ff872820c0ff2"]'),
        ("Processo Compra Finalizado", '//*[@id="field_8a3449076f9f6db3016fc95433971a26"]'),
        ("Data Aprovação Técnica", '//*[@id="field_8a3449076f9f6db3016fc9666f801d12"]'),
        ("Data Prevista Recebimento", '//*[@id="field_8a34490772473ce70172c30fab5e3842"]'),
    ]
    
    print("Dados do chamado ", numchamado, " extraídos.")
                
    for nome, xpath in campos:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        dados_dos_chamados[nome] = element.get_attribute("value")
    
    dados_dos_chamados["Descrição"] = titulo_limpo
    
    valor_final = dados_dos_chamados.get("Valor Final")
    valor_inicial = dados_dos_chamados.get("Valor Inicial")
    dados_dos_chamados["Valor R$"] = valor_final if valor_final else valor_inicial
    
    codigo_modalidade = dados_dos_chamados.get("Modalidade")
    dados_dos_chamados["Modalidade"] = modalidade_map.get(codigo_modalidade, codigo_modalidade)
             
    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()

    driver.switch_to.window(janela_principal)
    
    dados_dos_chamados["Status"] = status_texto
    
    return dados_dos_chamados

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

cabecalhos_esperados = ["Unidade", "Data Aprovação GP", "Identificador", "Atividade Habilitada",
                        "Nome Projeto", "Descrição", "Fonte", "CR", "Projeto",
                        "Conta", "Rubrica", "Valor R$", "Justificativa", "Justificativa GP",
                        "Data Análise Célula", "Modalidade", "Processo Compra Finalizado",
                        "Data Aprovação Técnica", "Data Prevista Recebimento",
                        "Data Atualização"]
valores_existentes = worksheet.get_all_records(expected_headers=cabecalhos_esperados)

linhas_existentes = worksheet.get_all_values()

mapa_identificador_linha = {
    str(int(float(linha[2]))).zfill(6): idx + 1
    for idx, linha in enumerate(linhas_existentes[1:])
    if len(linha) > 2 and linha[2].replace('.', '', 1).isdigit()
}

hoje = datetime.now().strftime("%d/%m/%Y")

pares_ja_processados = {
    (str(linha["Identificador"]).zfill(6), linha["Atividade Habilitada"]) for linha in valores_existentes
}

# Primeiro: processa chamados manuais
print("📌 Iniciando extração de chamados manuais...")
chamados_extraidos_com_sucesso = []

#Verifica dados a extrair manualmente e dados a ignorar
for idx, numero in enumerate(lista_manuais):
    numero_formatado = str(int(float(numero))).zfill(6)
    
    if numero_formatado in lista_ignorar:
        print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Chamado {numero_formatado} está na lista de ignorados. Pulando e removendo da lista manual.")
        remover_chamado_manuais(worksheet_manuais, numero_formatado)
        continue

    if (numero_formatado, "Encerrado") in pares_ja_processados or (numero_formatado, "Cancelado") in pares_ja_processados:
    #if (numero_formatado, "Chamado Encerrado") in pares_ja_processados:
        print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Chamado {numero_formatado} já encerrado. Pulando extração.")
        remover_chamado_manuais(worksheet_manuais, numero_formatado)
        continue
        
    print(f"[MANUAL {idx+1}/{len(lista_manuais)}] Acessando chamado {numero_formatado}")
    dados_dos_chamados = extrai_dados(numero_formatado)
    
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
        for col in ["Justificativa", "Justificativa GP"]:
            if isinstance(dados_dos_chamados.get(col), str):
                dados_dos_chamados[col] = dados_dos_chamados[col].replace('\n', ' ').strip()

        #dados_dos_chamados["Descrição"] = ""
        dados_dos_chamados["Atividade Habilitada"] = atividade_habilitada

        if isinstance(dados_dos_chamados.get("Valor R$"), str):
            dados_dos_chamados["Valor R$"] = dados_dos_chamados["Valor R$"].replace('.', '')

        dados_dos_chamados["Data Atualização"] = hoje

        linha_ordenada = [dados_dos_chamados.get(col, "") for col in cabecalhos_esperados]
        linha_existente = mapa_identificador_linha.get(numero_formatado)

        if linha_existente:
            worksheet.update(values=[linha_ordenada], range_name=f"A{linha_existente+1}")
            print(f"🔁 Chamado {numero_formatado} atualizado na linha {linha_existente+1}.")
        else:
            worksheet.append_row(linha_ordenada)
            print(f"➕ Chamado {numero_formatado} adicionado ao final da planilha.")

        remover_chamado_manuais(worksheet_manuais, numero_formatado)

# Segue com os chamados automáticos
for idx, numero in enumerate(num_chamados):
    #if (str(numero), atividadehabilitada[idx]) in pares_ja_processados:
    identificador_zfill = str(numero).zfill(6)
    
    if identificador_zfill in lista_ignorar:
        print(f"[{idx+1}/{len(num_chamados)}] Chamado {identificador_zfill} está na lista de ignorados. Pulando.")
        continue
    
    if (identificador_zfill, atividadehabilitada[idx]) in pares_ja_processados:
        print(f"[{idx+1}/{len(num_chamados)}] Chamado {numero} sem alteração de status. Pulando.")
        continue

    print(f"[{idx+1}/{len(num_chamados)}] Acessando chamado {identificador_zfill}")
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
        
        identificador_str = str(numero)
        linha_existente = mapa_identificador_linha.get(identificador_str)
        
        if linha_existente:
            worksheet.update(values=[linha_ordenada], range_name=f"A{linha_existente+1}")
            print(f"🔁 Chamado {identificador_str} atualizado na linha {linha_existente+1}.")
        else:
            worksheet.append_row(linha_ordenada)
            print(f"➕ Chamado {identificador_str} adicionado ao final da planilha.")

#%%        
















