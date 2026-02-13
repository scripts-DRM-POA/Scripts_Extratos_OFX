import os
import re
import hashlib
import pandas as pd
from tkinter import Tk, filedialog

# ==========================
# AUXILIARES
# ==========================

def escolher_diretorio():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Selecione a pasta com os extratos TXT")


def gerar_hash(data, historico, valor):
    base = f"{data}_{historico}_{valor}"
    return hashlib.md5(base.encode()).hexdigest()


def mes_nome_para_numero(mes):
    meses = {
        "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04",
        "MAI": "05", "JUN": "06", "JUL": "07", "AGO": "08",
        "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"
    }
    return meses.get(mes.upper())


# ==========================
# PROCESSAMENTO TXT BANRISUL
# ==========================

def processar_txt(caminho_txt):

    movimentos = []
    data_atual = None
    competencia_atual = None
    ano = None
    mes_num = None

    with open(caminho_txt, "r", encoding="latin-1") as f:
        linhas = f.readlines()

    for linha in linhas:

        linha = linha.rstrip("\n")

        # Detectar competência pelo conteúdo
        mov_match = re.search(r"MOVIMENTOS\s+([A-Z]{3})/(\d{4})", linha)
        if mov_match:
            mes_nome = mov_match.group(1)
            ano = mov_match.group(2)
            mes_num = mes_nome_para_numero(mes_nome)
            if mes_num:
                competencia_atual = f"{mes_num}/{ano}"
            continue

        if not competencia_atual:
            continue

        if "SALDO NA DATA" in linha:
            continue

        # Identificar dia
        match_dia = re.match(r"^\s*(\d{2})\s+(.*)", linha)

        if match_dia:
            dia = match_dia.group(1)
            data_atual = dia
        else:
            if not data_atual:
                continue

        # Extrair valor no final
        match_valor = re.search(r"(-?\d{1,3}(?:\.\d{3})*,\d{2}-?)\s*$", linha)

        if match_valor and data_atual:

            valor_str = match_valor.group(1)

            linha_sem_valor = linha.replace(valor_str, "").strip()

            # Documento (6 dígitos antes do valor)
            match_doc = re.search(r"(\d{6})\s*$", linha_sem_valor)
            documento = match_doc.group(1) if match_doc else ""

            if documento:
                historico = linha_sem_valor.replace(documento, "").strip()
            else:
                historico = linha_sem_valor.strip()

            # Converter valor
            valor = valor_str.replace(".", "").replace(",", ".")
            if valor.endswith("-"):
                valor = -float(valor[:-1])
            else:
                valor = float(valor)

            data_completa = f"{data_atual}/{mes_num}/{ano}"

            hash_id = gerar_hash(data_completa, historico, valor)

            movimentos.append([
                data_completa,
                historico,
                documento,
                valor,
                competencia_atual,
                hash_id
            ])

    return movimentos


# ==========================
# EXECUÇÃO PRINCIPAL
# ==========================

def main():

    pasta = escolher_diretorio()
    if not pasta:
        print("Nenhuma pasta selecionada.")
        return

    todos_movimentos = []

    for arquivo in os.listdir(pasta):
        if arquivo.lower().endswith(".txt"):
            caminho = os.path.join(pasta, arquivo)
            movimentos = processar_txt(caminho)
            todos_movimentos.extend(movimentos)

    if not todos_movimentos:
        print("Nenhum lançamento encontrado.")
        return

    df = pd.DataFrame(todos_movimentos, columns=[
        "Data", "Descrição", "Documento",
        "Valor", "Competência", "Hash"
    ])

    # Remove duplicados entre arquivos
    df = df.drop_duplicates(subset="Hash")

    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
    df = df.sort_values(["Competência", "Data"])

    # Crédito / Débito
    df["Crédito"] = df["Valor"].apply(lambda x: x if x > 0 else 0)
    df["Débito"] = df["Valor"].apply(lambda x: abs(x) if x < 0 else 0)

    # Saldo acumulado por competência
    df["Saldo"] = df.groupby("Competência")["Valor"].cumsum()

    # Resumo mensal
    resumo = df.groupby("Competência").agg({
        "Crédito": "sum",
        "Débito": "sum",
        "Valor": "sum"
    }).reset_index()

    resumo.rename(columns={"Valor": "Saldo_Líquido"}, inplace=True)

    # Exportação
    caminho_saida = os.path.join(pasta, "Extrato_Consolidado.xlsx")

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:

        colunas = [
            "Data", "Descrição", "Documento",
            "Crédito", "Débito", "Saldo", "Competência"
        ]

        df[colunas].to_excel(writer, sheet_name="Consolidado", index=False)

        for comp in df["Competência"].unique():
            df_comp = df[df["Competência"] == comp]
            nome_aba = comp.replace("/", "-")
            df_comp[colunas].to_excel(writer, sheet_name=nome_aba, index=False)

        resumo.to_excel(writer, sheet_name="Resumo_Mensal", index=False)

    print("\nProcessamento concluído com sucesso.")
    print(f"Arquivo gerado em:\n{caminho_saida}")


if __name__ == "__main__":
    main()
