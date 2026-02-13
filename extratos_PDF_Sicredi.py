import pdfplumber
import pandas as pd
import re
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

# ==========================================
# 1️⃣ Selecionar pasta
# ==========================================

Tk().withdraw()
pasta = askdirectory(title="Selecione a pasta com os extratos Sicredi")

arquivos_pdf = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]

if not arquivos_pdf:
    print("❌ Nenhum PDF encontrado.")
else:
    print(f"📄 {len(arquivos_pdf)} PDFs encontrados.")

# ==========================================
# 2️⃣ Funções auxiliares
# ==========================================

def converter(valor):
    valor = valor.replace(".", "").replace(",", ".")
    return float(valor)

def eh_valor_monetario(texto):
    return re.match(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$", texto)

padrao_data = re.compile(r"\d{2}/\d{2}/\d{4}")

# ==========================================
# 3️⃣ Processamento
# ==========================================

df_final = pd.DataFrame()
log_inconsistencias = []

for arquivo in arquivos_pdf:
    
    caminho = os.path.join(pasta, arquivo)
    print(f"\n🔎 Processando: {arquivo}")
    
    linhas = []
    saldo_anterior = 0
    
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                for linha in texto.split("\n"):
                    linha = linha.rstrip()
                    
                    if "SALDO ANTERIOR" in linha.upper():
                        match_saldo = re.search(r"-?\d{1,3}(?:\.\d{3})*,\d{2}", linha)
                        if match_saldo:
                            saldo_anterior = converter(match_saldo.group())
                    
                    linhas.append(linha)
    
    registros = []
    
    for linha in linhas:
        
        if not padrao_data.match(linha):
            continue
        
        try:
            # 1️⃣ Data = primeiros 10 caracteres
            data = linha[:10]
            
            restante = linha[10:].strip()
            
            # 2️⃣ Separar da direita para esquerda
            partes = restante.split()
            
            if len(partes) < 2:
                continue
            
            # Últimos dois são Saldo e Valor
            saldo_str = partes[-1]
            valor_str = partes[-2]
            
            if not eh_valor_monetario(saldo_str) or not eh_valor_monetario(valor_str):
                log_inconsistencias.append((arquivo, linha))
                continue
            
            saldo = converter(saldo_str)
            valor = converter(valor_str)
            
            # Documento pode existir ou não
            documento = ""
            descricao_partes = partes[:-2]
            
            if len(descricao_partes) >= 1:
                # Se houver algo antes do valor
                documento = descricao_partes[-1]
                descricao = " ".join(descricao_partes[:-1])
                
                # Se o documento for claramente parte do texto (ex: PIX_CRED ok)
                # mas se for numérico ou alfanumérico simples também aceitamos
            else:
                descricao = ""
            
            registros.append([
                data,
                descricao.strip(),
                documento.strip(),
                valor,
                saldo,
                saldo_anterior,
                arquivo
            ])
        
        except Exception:
            log_inconsistencias.append((arquivo, linha))
    
    if not registros:
        continue
    
    df = pd.DataFrame(registros, columns=[
        "Data", "Descrição", "Documento",
        "Valor", "Saldo",
        "Saldo_Anterior",
        "Arquivo_Origem"
    ])
    
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
    
    df["Débito"] = df["Valor"].apply(lambda x: x if x < 0 else 0)
    df["Crédito"] = df["Valor"].apply(lambda x: x if x > 0 else 0)
    
    # Conferência de saldo
    df["Saldo_Calculado"] = df["Saldo_Anterior"].iloc[0] + df["Valor"].cumsum()
    df["Diferença_Saldo"] = df["Saldo"] - df["Saldo_Calculado"]
    
    df_final = pd.concat([df_final, df], ignore_index=True)

# ==========================================
# 4️⃣ Consolidação Final
# ==========================================

if not df_final.empty:
    
    df_final = df_final.sort_values(["Data", "Arquivo_Origem"])
    
    df_final = df_final[
        ["Data", "Descrição", "Documento",
         "Débito", "Crédito",
         "Saldo", "Saldo_Calculado",
         "Diferença_Saldo",
         "Arquivo_Origem"]
    ]
    
    caminho_saida = os.path.join(pasta, "Extrato_Sicredi_Consolidado.xlsx")
    df_final.to_excel(caminho_saida, index=False)
    
    print("\n✅ Arquivo consolidado gerado com sucesso.")
    print(f"📁 Local: {caminho_saida}")
    
    if log_inconsistencias:
        print(f"\n⚠️ {len(log_inconsistencias)} linhas com possível inconsistência.")
    else:
        print("\n✔ Nenhuma inconsistência estrutural detectada.")

else:
    print("❌ Nenhum dado estruturado foi extraído.")
