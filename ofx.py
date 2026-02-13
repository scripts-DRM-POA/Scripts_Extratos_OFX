# ============================================================
# PROCESSADOR OFX / XML → XLSX CONSOLIDADO (VERSÃO FINAL)
# ============================================================

from pathlib import Path
from datetime import datetime
import re
import uuid
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

# ============================================================
# UTILITÁRIOS DE NORMALIZAÇÃO (BLINDADOS)
# ============================================================

def normalizar_valor_br(valor: str) -> float | None:
    """
    Normaliza valores no padrão brasileiro:
    - aceita "- 1.234,56"
    - aceita "1.234.567,89"
    - aceita "-1234.56"
    Retorna float ou None
    """
    if not valor:
        return None

    v = valor.strip()

    # remove espaços entre sinal e número
    v = re.sub(r'^-\s+', '-', v)

    # mantém apenas dígitos, ponto, vírgula e sinal
    v = re.sub(r'[^\d\-,\.]', '', v)

    # se tiver vírgula, assume padrão brasileiro
    if ',' in v:
        v = v.replace('.', '').replace(',', '.')
    else:
        # se houver mais de um ponto, mantém só o último como decimal
        if v.count('.') > 1:
            partes = v.split('.')
            v = ''.join(partes[:-1]) + '.' + partes[-1]

    try:
        return float(v)
    except:
        return None


def normalizar_data(dt: str) -> str:
    if not dt:
        return ""
    m = re.search(r'(\d{8,14})', dt)
    if not m:
        return ""
    s = m.group(1)
    try:
        ano = int(s[:4])
        ano_atual = datetime.now().year
        if ano < 1900 or ano > ano_atual + 1:
            return datetime.now().strftime("%Y%m%d")
        return s[:8]
    except:
        return ""


# ============================================================
# 1. CORREÇÃO / RECONSTRUÇÃO OFX → XML
# ============================================================

def corrigir_ofx_para_xml(ofx: Path, xml_saida: Path) -> bool:
    texto = ofx.read_text(encoding="latin1", errors="ignore")
    texto = texto[texto.find("<"):]

    stmts = re.findall(r'<STMTTRN>(.*?)</STMTTRN>', texto, flags=re.S | re.I)
    transacoes = []

    for raw in stmts:
        def campo(tag):
            m = re.search(rf'<{tag}>(.*?)($|<)', raw, flags=re.I | re.S)
            return m.group(1).strip() if m else ""

        valor = normalizar_valor_br(campo("TRNAMT"))

        transacoes.append({
            "DTPOSTED": normalizar_data(campo("DTPOSTED")),
            "TRNAMT": valor,
            "MEMO": campo("MEMO"),
            "CHECKNUM": campo("CHECKNUM"),
            "FITID": campo("FITID") or uuid.uuid4().hex
        })

    if not transacoes:
        return False

    dtstart = min(t["DTPOSTED"] for t in transacoes if t["DTPOSTED"])
    dtend = max(t["DTPOSTED"] for t in transacoes if t["DTPOSTED"])

    xml = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<OFX><BANKMSGSRSV1><STMTTRNRS><STMTRS>',
        '<BANKACCTFROM><BANKID>000</BANKID><ACCTID>000</ACCTID><ACCTTYPE>CHECKING</ACCTTYPE></BANKACCTFROM>',
        f'<BANKTRANLIST><DTSTART>{dtstart}</DTSTART><DTEND>{dtend}</DTEND>'
    ]

    for t in transacoes:
        xml.extend([
            '<STMTTRN>',
            '<TRNTYPE>OTHER</TRNTYPE>',
            f'<DTPOSTED>{t["DTPOSTED"]}</DTPOSTED>',
            f'<TRNAMT>{t["TRNAMT"]}</TRNAMT>',
            f'<FITID>{t["FITID"]}</FITID>',
            f'<MEMO>{t["MEMO"]}</MEMO>',
            f'<CHECKNUM>{t["CHECKNUM"]}</CHECKNUM>',
            '</STMTTRN>'
        ])

    xml.extend([
        '</BANKTRANLIST></STMTRS></STMTTRNRS></BANKMSGSRSV1></OFX>'
    ])

    xml_saida.write_text("\n".join(xml), encoding="utf-8")

    try:
        ET.parse(xml_saida)
        return True
    except:
        return False


# ============================================================
# 2. EXTRAÇÃO XML → DATAFRAME
# ============================================================

def extrair_dataframe(xml: Path) -> pd.DataFrame | None:
    tree = ET.parse(xml)
    root = tree.getroot()

    rows = []
    for t in root.findall(".//STMTTRN"):
        get = lambda x: (t.findtext(x) or "").strip()

        valor = normalizar_valor_br(get("TRNAMT"))

        rows.append({
            "DATA": pd.to_datetime(normalizar_data(get("DTPOSTED")), format="%Y%m%d", errors="coerce"),
            "VALOR": valor,
            "HISTORICO": get("MEMO"),
            "DOCUMENTO": get("CHECKNUM"),
            "FITID": get("FITID")
        })

    df = pd.DataFrame(rows).dropna(subset=["VALOR", "DATA"])

    # DÉBITO / CRÉDITO
    df["CREDITO"] = df["VALOR"].apply(lambda x: x if x > 0 else "")
    df["DEBITO"] = df["VALOR"].apply(lambda x: x if x < 0 else "")
    df["TIPO"] = df["VALOR"].apply(lambda x: "C" if x > 0 else "D")

    df["DATA"] = df["DATA"].dt.strftime("%d/%m/%Y")

    return df


# ============================================================
# 3. PROCESSO PRINCIPAL (SELEÇÃO DE PASTA)
# ============================================================

def processar_pasta():
    root = tk.Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta com arquivos OFX/XML")

    if not pasta:
        print("Nenhuma pasta selecionada.")
        return

    arquivos = list(Path(pasta).glob("*.ofx")) + list(Path(pasta).glob("*.xml"))

    saida = Path(pasta) / f"consolidado_ofx_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    writer = pd.ExcelWriter(saida, engine="openpyxl")

    todos = []

    for arq in arquivos:
        print("Processando:", arq.name)

        if arq.suffix.lower() == ".ofx":
            xml = arq.with_suffix(".corrigido.xml")
            if not corrigir_ofx_para_xml(arq, xml):
                continue
        else:
            xml = arq

        df = extrair_dataframe(xml)
        if df is None or df.empty:
            continue

        df["ARQUIVO"] = arq.name
        todos.append(df)

        df.to_excel(writer, sheet_name=arq.stem[:31], index=False)

    if todos:
        consolidado = pd.concat(todos, ignore_index=True)

        # DEDUPLICAÇÃO DEFINITIVA
        consolidado = consolidado.drop_duplicates(
            subset=["FITID", "DATA", "VALOR"]
        )

        ordem = ["ARQUIVO", "DATA", "VALOR", "TIPO", "HISTORICO", "DOCUMENTO", "CREDITO", "DEBITO", "FITID"]
        consolidado = consolidado[ordem]

        consolidado.to_excel(writer, sheet_name="CONSOLIDADO", index=False)

    writer.close()

    load_workbook(saida).close()
    print("Arquivo final gerado com sucesso:", saida)


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    processar_pasta()
