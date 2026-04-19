# -*- coding: utf-8 -*-
"""
Created on Sun Apr 12 17:04:02 2026

@author: davi.hulse
"""

import re
import pdfplumber

# =============================================
# ALTERE O CAMINHO DO PDF AQUI
# =============================================
#CAMINHO_PDF = r"C:\Users\davi.hulse\Downloads\AF 2156.pdf"
#CAMINHO_PDF = r"C:\Users\davi.hulse\Downloads\893.pdf"
#CAMINHO_PDF = r"C:\Users\davi.hulse\Downloads\258.pdf" #Esse dá erro mesmo, está em imagem
CAMINHO_PDF = r"C:\Users\davi.hulse\Downloads\596.pdf"
#CAMINHO_PDF = r"C:\Users\davi.hulse\Downloads\696erro.pdf"



with pdfplumber.open(CAMINHO_PDF) as pdf:
    for page in pdf.pages:
        print(page.extract_text())
        
# with pdfplumber.open(CAMINHO_PDF) as pdf:
#     for i, page in enumerate(pdf.pages):
#         print(f"--- Página {i+1} ---")
#         print(page.extract_text()[:200])

#%%
# =============================================

CNPJS_INTERNOS = ["03.774.688/0054-67", "03.774.688/0055-48"]

def extrair_dados_oc(texto_pdf):
    numero_oc = ""
    data_emissao = ""
    nome_fornecedor_pdf = ""
    cnpj_fornecedor_pdf = ""
    prazo_entrega_pdf = ""

    primeiras_linhas = "\n".join(texto_pdf.splitlines()[:10])
    print(f"🔍 Primeiras linhas:\n{primeiras_linhas}\n")

    if "Número AF:" in primeiras_linhas:
        print("📋 Modelo identificado: Modelo 1")
        texto_limpo = re.sub(r'\(cid:\d+\)', ' ', texto_pdf)
        match_num = re.search(r'Número AF:\s*([\d.]+)', texto_limpo)
        match_data = re.search(r'Data:\s*(\d{2}/\d{2}/\d{4})', texto_limpo)
        match_fornecedor = re.search(r'Razão social:\s*(.+)', texto_limpo)
        match_cnpj = re.search(r'DADOS DO FORNECEDOR.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_limpo, re.DOTALL)
        match_prazo = re.search(r'Prazos de entrega.*?/\s*(\d{2}/\d{2}/\d{4})', texto_limpo, re.DOTALL)
        numero_oc = match_num.group(1).replace('.', '') if match_num else ""
        data_emissao = match_data.group(1) if match_data else ""
        nome_fornecedor_pdf = match_fornecedor.group(1).strip() if match_fornecedor else ""
        cnpj_raw = match_cnpj.group(1).strip() if match_cnpj else ""
        cnpj_fornecedor_pdf = "" if cnpj_raw in CNPJS_INTERNOS else cnpj_raw
        prazo_entrega_pdf = match_prazo.group(1).strip() if match_prazo else ""

    elif "Ordem de compra" in primeiras_linhas:
        print("📋 Modelo identificado: Modelo 2")
        match_num = re.search(r'Nº\s+(\d+)\s+Valor Total:', texto_pdf)
        match_data = re.search(r'DATA EMISSÃO\s+(\d{2}/\d{2}/\d{4})', texto_pdf)
        match_fornecedor = re.search(r'Empresa Fornecedora:\s*(.+?)\s*CNPJ:', texto_pdf)
        match_cnpj = re.search(r'Empresa Fornecedora:.*?CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_pdf, re.DOTALL)
        match_prazo = next((re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l) for l in texto_pdf.split('\n') if re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l)), None)
        numero_oc = match_num.group(1) if match_num else ""
        data_emissao = match_data.group(1) if match_data else ""
        nome_fornecedor_pdf = match_fornecedor.group(1).strip() if match_fornecedor else ""
        cnpj_raw = match_cnpj.group(1).strip() if match_cnpj else ""
        cnpj_fornecedor_pdf = "" if cnpj_raw in CNPJS_INTERNOS else cnpj_raw
        prazo_entrega_pdf = match_prazo.group(1).strip() if match_prazo else ""

    elif re.search(r'Ordem\s+\d+', primeiras_linhas, re.DOTALL):
        print("📋 Modelo identificado: Modelo 3")
        match_num_raw = re.search(r'Ordem\s+(\d[\d,]+)', texto_pdf, re.DOTALL)
        #match_valor = re.search(r'Valor Total:\s*\n.*?(\d[\d,]+)', texto_pdf, re.DOTALL)
        match_valor = re.search(r'Valor Total:\s*[\r\n]+\S+[\r\n]+(\d[\d,.]+)', texto_pdf)
        print(f"DEBUG match_num_raw: {match_num_raw.group(1) if match_num_raw else 'None'}")
        print(f"DEBUG match_valor: {match_valor.group(1) if match_valor else 'None'}")
        num_raw = match_num_raw.group(1).split(',')[0] if match_num_raw else ""
      
        #match_valor_item = re.search(r'([\d.]+),([\d]{2})\d{2}/\d{2}/\d{4}', texto_pdf)
        #match_valor_item = re.search(r'[\d,]+(\d{1,3}(?:\.\d{3})*),\d{2}\d{2}/\d{2}/\d{4}', texto_pdf)
        #match_valor_item = re.search(r'(?<!\d)(\d{1,3}(?:\.\d{3})*),\d{2}\d{2}/\d{2}/\d{4}', texto_pdf)
        
        match_valor_item = re.search(r',(\d{1,3}(?:\.\d{3})*),\d{2}(?=\d{2}/\d{2}/\d{4})', texto_pdf)
        print(f"DEBUG num_raw: {num_raw}")
        print(f"DEBUG match_valor_item: {match_valor_item.group(1) if match_valor_item else 'None'}")
        
        match_valor_item = re.search(r',(\d{1,3}(?:\.\d{3})*),\d{2}(?=\d{2}/\d{2}/\d{4})', texto_pdf)
        if match_valor_item:
            valor_str = match_valor_item.group(1).replace('.', '')
            numero_oc = num_raw[:-len(valor_str)] if num_raw.endswith(valor_str) else num_raw
        else:
            numero_oc = num_raw
      
        # match_valor_item = re.search(r'1,0000\s+([\d.]+),([\d]+)\1', texto_pdf)
        # if match_valor_item:
        #     valor_str = match_valor_item.group(1).replace('.', '')
        #     numero_oc = num_raw[:-len(valor_str)] if num_raw.endswith(valor_str) else num_raw
        # else:
        #     numero_oc = num_raw
        
        
        
        match_data = re.search(r'DATA EMISSÃO\s+(\d{2}/\d{2}/\d{4})', texto_pdf, re.DOTALL)
        match_fornecedor = re.search(r'Empresa Fornecedora:\s*(.+?)\s*CNPJ:', texto_pdf)
        match_cnpj = re.search(r'Empresa Fornecedora:.*?CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_pdf, re.DOTALL)
        match_prazo = next((re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l) for l in texto_pdf.split('\n') if re.search(r'(\d{2}/\d{2}/\d{4})\s*$', l)), None)
        data_emissao = match_data.group(1) if match_data else ""
        nome_fornecedor_pdf = match_fornecedor.group(1).strip() if match_fornecedor else ""
        cnpj_raw = match_cnpj.group(1).strip() if match_cnpj else ""
        cnpj_fornecedor_pdf = "" if cnpj_raw in CNPJS_INTERNOS else cnpj_raw
        prazo_entrega_pdf = match_prazo.group(1).strip() if match_prazo else ""

    else:
        print("⚠️ Nenhum modelo identificado.")

    return numero_oc, data_emissao, nome_fornecedor_pdf, cnpj_fornecedor_pdf, prazo_entrega_pdf


def main():
    print(f"📂 Lendo PDF: {CAMINHO_PDF}\n")
    with pdfplumber.open(CAMINHO_PDF) as pdf:
        #texto_pdf = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
        texto_pdf = "\n".join(page.extract_text(x_tolerance=1) for page in pdf.pages if page.extract_text(x_tolerance=1))

    print(f"📄 Texto extraído do PDF:\n{texto_pdf}\n")
    print("=" * 60)

    numero_oc, data_emissao, nome_fornecedor, cnpj_fornecedor, prazo_entrega = extrair_dados_oc(texto_pdf)

    print("\n✅ Dados extraídos:")
    print(f"  Número OC:       {numero_oc}")
    print(f"  Data Emissão:    {data_emissao}")
    print(f"  Nome Fornecedor: {nome_fornecedor}")
    print(f"  CNPJ Fornecedor: {cnpj_fornecedor}")
    print(f"  Prazo Entrega:   {prazo_entrega}")


if __name__ == "__main__":
    main()