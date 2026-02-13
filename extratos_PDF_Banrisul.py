import pdfplumber
import pandas as pd
import re
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

Tk().withdraw()

# Selecionar pasta
pasta = askdirectory(title="Selecione a pasta com os extratos Banrisul")

if not pasta:
    raise Exception("Nenhuma pasta selecionada.")

dados_totais = []

meses = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

padrao_valor = re.compile(r"^\d{1,3}(\.\d{3})*,\d{2}-?$")

# Percorrer PDFs da pasta
for arquivo in os.listdir(pasta):
    if not arquivo.lower().endswith(".pdf"):
        continue

    caminho = os.path.join(pasta, arquivo)
    linhas = []

    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                linhas.extend(texto.split("\n"))

    # Capturar competência
    mes = None
    ano = None

    for linha in linhas:
        if "PERIODO:" in linha:
            periodo = linha.split(":")[1].strip()
            nome_mes, ano = periodo.split("/")
            ano = int(ano)
            mes = meses[nome_mes.upper()]
            break

    if not mes:
        continue

    meses_abrev = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr",
    5: "mai", 6: "jun", 7: "jul", 8: "ago",
    9: "set", 10: "out", 11: "nov", 12: "dez"
    }
    
    competencia = f"{meses_abrev[mes]}/{str(ano)[-2:]}"

    dia_atual = None

    for linha in linhas:

        linha = linha.strip()

        if not linha or \
           "SALDO ANT" in linha or \
           "SALDO NA DATA" in linha or \
           "MOVIMENTOS" in linha or \
           "DIA HISTORICO" in linha:
            continue

        match_dia = re.match(r"^(\d{2})\s+(.*)", linha)

        if match_dia:
            dia_atual = int(match_dia.group(1))
            restante = match_dia.group(2)
        else:
            if dia_atual is None:
                continue
            restante = linha

        partes = restante.split()

        if len(partes) < 2:
            continue

        valor_str = partes[-1]

        if not padrao_valor.match(valor_str):
            continue

        documento = partes[-2]
        historico = " ".join(partes[:-2])

        negativo = valor_str.endswith("-")
        valor_str = valor_str.replace("-", "")
        valor = float(valor_str.replace(".", "").replace(",", "."))

        data = pd.Timestamp(year=ano, month=mes, day=dia_atual)

        debito = valor if negativo else 0.0
        credito = valor if not negativo else 0.0

        dados_totais.append([
            data.strftime("%d/%m/%Y"),
            competencia,
            historico,
            documento,
            debito,
            credito
        ])

# Criar DataFrame consolidado
df = pd.DataFrame(dados_totais, columns=[
    "Data",
    "Competência",
    "Histórico",
    "Documento",
    "Débito",
    "Crédito"
])

# Ordenar
df["Data_ord"] = pd.to_datetime(df["Data"], dayfirst=True)
df = df.sort_values(["Data_ord"]).drop(columns="Data_ord")

# ----------------------------
# Exportar Excel
# ----------------------------

arquivo_saida = os.path.join(pasta, "Extratos_Banrisul_Consolidado.xlsx")

with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:

    # Aba Consolidado
    df.to_excel(writer, sheet_name="Consolidado", index=False)

    # Abas por competência
    for comp in df["Competência"].unique():
        df_comp = df[df["Competência"] == comp]
        nome_aba = comp.replace("/", "_")
        df_comp.to_excel(writer, sheet_name=nome_aba, index=False)

print("Arquivo gerado com sucesso:")
print(arquivo_saida)
