# ======================================================
# 🔧 Imports
# ======================================================
from tkinter import Tk, filedialog
from pathlib import Path
from datetime import datetime
import re
import xml.etree.ElementTree as ET
import pandas as pd


# ======================================================
# 🧱 Função 1: Corrigir arquivo OFX para XML válido
# ======================================================
def corrigir_ofx_para_xml(caminho_ofx: Path, caminho_xml: Path) -> bool:
    with open(caminho_ofx, 'r', encoding='latin1', errors='ignore') as f:
        linhas = f.readlines()

    # Procura o início do <OFX>
    for i, ln in enumerate(linhas):
        if re.match(r'<\s*ofx\s*>', ln.strip(), re.IGNORECASE):
            linhas = linhas[i:]
            break

    # Monta XML corrigido
    corr = ['<?xml version="1.0" encoding="UTF-8"?>']
    padrao = re.compile(r'^<(\w+)>([^<]+)$')
    for l in linhas:
        ln = l.strip()
        m = padrao.match(ln)
        if m:
            tag, valor = m.groups()
            corr.append(f"<{tag}>{valor.strip()}</{tag}>")
        else:
            corr.append(ln)

    # Grava o arquivo temporário corrigido
    with open(caminho_xml, 'w', encoding='utf-8') as f:
        f.write("\n".join(corr))

    # Verifica se é XML válido
    try:
        ET.parse(caminho_xml)
        return True
    except ET.ParseError:
        return False


# ======================================================
# 🧱 Função 2: Extrair transações para DataFrame
# ======================================================
def extrair_dataframe(caminho_xml: Path):
    tree = ET.parse(caminho_xml)
    root = tree.getroot()
    trans = root.findall(".//STMTTRN")
    if not trans:
        return None

    rows = []
    for t in trans:
        rows.append({el.tag: (el.text or '').strip() for el in t})
    df = pd.DataFrame(rows)

    # Ajusta datas
    if 'DTPOSTED' in df.columns:
        df['DTPOSTED'] = (
            df['DTPOSTED']
            .str.extract(r'(\d{8,14})')[0]
            .apply(lambda x: pd.to_datetime(
                x,
                format='%Y%m%d%H%M%S' if pd.notna(x) and len(x) > 8 else '%Y%m%d',
                errors='coerce'
            ))
            .dt.strftime('%d/%m/%Y')
        )

    # Ajusta valores
    if 'TRNAMT' in df.columns:
        df['TRNAMT'] = df['TRNAMT'].str.replace(',', '.', regex=False).astype(float)
        df['CREDITO'] = df['TRNAMT'].apply(lambda x: x if x > 0 else "")
        df['DEBITO'] = df['TRNAMT'].apply(lambda x: x if x < 0 else "")
        df['TIPO'] = df['TRNAMT'].apply(lambda x: 'C' if x > 0 else ('D' if x < 0 else ''))
    else:
        df['CREDITO'] = ""
        df['DEBITO'] = ""
        df['TIPO'] = ""

    df.rename(columns={'DTPOSTED': 'DATA', 'TRNAMT': 'VALOR', 'MEMO': 'HISTORICO', 'CHECKNUM': 'DOCUMENTO'}, inplace=True)
    cols = ['DATA', 'VALOR', 'TIPO', 'HISTORICO', 'DOCUMENTO', 'CREDITO', 'DEBITO']
    return df[[c for c in cols if c in df.columns]]


# ======================================================
# 🧱 Função 3: Selecionar pasta e processar arquivos
# ======================================================
def process_dir() -> Path:
    # 🪟 Abre janela para selecionar a pasta
    Tk().withdraw()
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta com arquivos .OFX")

    if not pasta_selecionada:
        raise FileNotFoundError("Nenhuma pasta selecionada.")

    pasta = Path(pasta_selecionada)
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    arquivos = list(pasta.glob("*.ofx"))
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo .ofx na pasta selecionada.")

    # Gera nome do arquivo de saída
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    saida = pasta / f"consolidado_ofx_{ts}.xlsx"

    writer = pd.ExcelWriter(saida, engine="openpyxl")
    dfs = []

    # Processa cada OFX
    for arq in arquivos:
        print(f"➡️ Processando: {arq.name}")
        xml_corr = arq.with_name(arq.stem + "_corrigido.xml")
        if corrigir_ofx_para_xml(arq, xml_corr):
            df = extrair_dataframe(xml_corr)
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=arq.stem[:31], index=False)
                df['ARQUIVO'] = arq.name
                dfs.append(df)

    # Consolida tudo em uma planilha
    if dfs:
        total = pd.concat(dfs, ignore_index=True)
        ordem = ['ARQUIVO', 'DATA', 'VALOR', 'TIPO', 'HISTORICO', 'DOCUMENTO', 'CREDITO', 'DEBITO']
        total = total[[c for c in ordem if c in total.columns]]
        total.to_excel(writer, sheet_name="CONSOLIDADO", index=False)

    writer.close()
    print(f"✅ Consolidado salvo em: {saida}")
    return saida


# ======================================================
# ▶️ Execução automática ao rodar o script
# ======================================================
if __name__ == "__main__":
    try:
        process_dir()
    except Exception as e:
        print(f"❌ Erro: {e}")



