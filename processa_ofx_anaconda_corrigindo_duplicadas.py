# ============================================================
# PROCESSADOR DE ARQUIVOS OFX/XML ➜ XLSX CONSOLIDADO
# (com reparo de OFX quebrado + FITID + deduplicação correta)
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

# ------------------------------------------------------------
# Utilitários
# ------------------------------------------------------------
def _normalize_decimal(valor: str) -> str:
    if valor is None:
        return ""
    v = valor.strip().replace(".", "").replace(",", ".") if re.match(r'^[\d\.\,]+$', valor.strip()) else valor.strip()
    if v.count('.') > 1 and re.search(r'\d+\.\d+$', v):
        parts = v.split('.')
        v = ''.join(parts[:-1]) + '.' + parts[-1]
    return v

def _normalize_date(dt_text: str) -> str:
    if not dt_text:
        return ""
    m = re.search(r'(\d{8,14})', dt_text)
    if not m:
        return ""
    s = m.group(1)
    try:
        year = int(s[0:4])
        ano_atual = datetime.now().year
        if year < 1900 or year > ano_atual + 1:
            return datetime.now().strftime("%Y%m%d%H%M%S")
        return s
    except:
        return datetime.now().strftime("%Y%m%d%H%M%S")

# ============================================================
# 1. CORRIGE OFX QUEBRADO ➜ XML VÁLIDO
# ============================================================
def corrigir_ofx_para_xml(caminho_ofx: Path, caminho_xml: Path) -> bool:
    try:
        texto = caminho_ofx.read_text(encoding="latin1", errors="ignore")
    except Exception as e:
        print(f"Erro lendo {caminho_ofx.name}: {e}")
        return False

    idx = re.search(r'<', texto)
    texto_body = texto[idx.start():] if idx else texto

    m_bankacct = re.search(r'<BANKACCTFROM>(.*?)</BANKACCTFROM>', texto_body, flags=re.S | re.I)
    bankacct = m_bankacct.group(0) if m_bankacct else ""

    m_ledger = re.search(r'<LEDGERBAL>(.*?)</LEDGERBAL>', texto_body, flags=re.S | re.I)
    ledgerbal = m_ledger.group(0) if m_ledger else ""

    m_banktran = re.search(r'<BANKTRANLIST>(.*?)</BANKTRANLIST>', texto_body, flags=re.S | re.I)
    banktran_inner = m_banktran.group(1) if m_banktran else ""

    stmts_raw = re.findall(r'<STMTTRN>(.*?)</STMTTRN>', banktran_inner, flags=re.S | re.I)
    stmt_entries = []

    for raw in stmts_raw:
        campos = {}
        for tag in ["TRNTYPE", "DTPOSTED", "TRNAMT", "FITID", "MEMO", "CHECKNUM"]:
            rx = re.search(rf'<{tag}>(.*?)($|<)', raw, flags=re.S | re.I)
            campos[tag] = rx.group(1).strip() if rx else ""
        campos["TRNAMT"] = _normalize_decimal(campos["TRNAMT"])
        campos["DTPOSTED"] = _normalize_date(campos["DTPOSTED"])
        if not campos["FITID"]:
            campos["FITID"] = uuid.uuid4().hex
        stmt_entries.append(campos)

    if not stmt_entries:
        print(f"❌ Nenhuma transação localizada em {caminho_ofx.name}")
        return False

    dt_values = [s["DTPOSTED"] for s in stmt_entries if s["DTPOSTED"]]
    dtstart = min(dt_values)
    dtend = max(dt_values)

    bal_amt = ""
    if ledgerbal:
        m_bal = re.search(r'<BALAMT>(.*?)($|<)', ledgerbal, flags=re.S | re.I)
        if m_bal:
            bal_amt = _normalize_decimal(m_bal.group(1))

    if not bankacct:
        bankacct = (
            "<BANKACCTFROM>"
            "<BANKID>UNKNOWN</BANKID>"
            "<BRANCHID>UNKNOWN</BRANCHID>"
            "<ACCTID>UNKNOWN</ACCTID>"
            "<ACCTTYPE>CHECKING</ACCTTYPE>"
            "</BANKACCTFROM>"
        )

    stmts_xml = []
    for s in stmt_entries:
        trn = [
            "<STMTTRN>",
            f"<TRNTYPE>{s['TRNTYPE'] or 'OTHER'}</TRNTYPE>",
            f"<DTPOSTED>{s['DTPOSTED']}</DTPOSTED>",
            f"<TRNAMT>{s['TRNAMT']}</TRNAMT>",
            f"<FITID>{s['FITID']}</FITID>",
        ]
        if s["MEMO"]:
            trn.append(f"<MEMO>{s['MEMO']}</MEMO>")
        if s["CHECKNUM"]:
            trn.append(f"<CHECKNUM>{s['CHECKNUM']}</CHECKNUM>")
        trn.append("</STMTTRN>")
        stmts_xml.append("\n".join(trn))

    ofx_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<OFX>
<BANKMSGSRSV1>
<STMTTRNRS>
<STMTRS>
{bankacct}
<BANKTRANLIST>
<DTSTART>{dtstart}</DTSTART>
<DTEND>{dtend}</DTEND>
{chr(10).join(stmts_xml)}
</BANKTRANLIST>
{"<LEDGERBAL><BALAMT>"+bal_amt+"</BALAMT><DTASOF>"+dtend+"</DTASOF></LEDGERBAL>" if bal_amt else ""}
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>
"""

    try:
        caminho_xml.write_text(ofx_xml, encoding="utf-8")
        ET.parse(caminho_xml)
        return True
    except Exception as e:
        print(f"❌ Falha ao gerar XML: {e}")
        return False

# ============================================================
# 2. EXTRAI DATAFRAME DO XML
# ============================================================
def extrair_dataframe(caminho_xml: Path):
    try:
        tree = ET.parse(caminho_xml)
    except:
        return None

    root = tree.getroot()
    trans = root.findall(".//STMTTRN")
    if not trans:
        return None

    rows = [{el.tag: (el.text or '').strip() for el in t} for t in trans]
    df = pd.DataFrame(rows)

    df["DATA"] = pd.to_datetime(
        df["DTPOSTED"].str.extract(r"(\d{8})")[0],
        format="%Y%m%d",
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    df["VALOR"] = pd.to_numeric(df["TRNAMT"].str.replace(",", "."), errors="coerce")
    df["CREDITO"] = df["VALOR"].apply(lambda x: x if x > 0 else 0)
    df["DEBITO"] = df["VALOR"].apply(lambda x: abs(x) if x < 0 else 0)
    df["TIPO"] = df["VALOR"].apply(lambda x: "C" if x > 0 else "D")

    df.rename(columns={
        "MEMO": "HISTORICO",
        "CHECKNUM": "DOCUMENTO"
    }, inplace=True)

    cols = ["FITID", "DATA", "VALOR", "TIPO", "HISTORICO", "DOCUMENTO", "CREDITO", "DEBITO"]
    return df[cols]

# ============================================================
# 3. PROCESSO PRINCIPAL
# ============================================================
def process_dir():
    root = tk.Tk()
    root.withdraw()

    arquivos = filedialog.askopenfilenames(
        title="Selecione OFX/XML",
        filetypes=[("OFX/XML", "*.ofx *.xml")]
    )

    if not arquivos:
        return

    arquivos = [Path(a) for a in arquivos]
    saida = arquivos[0].parent / f"consolidado_ofx_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    writer = pd.ExcelWriter(saida, engine="openpyxl")
    dfs = []

    for arq in arquivos:
        print(f"➡ {arq.name}")
        if arq.suffix.lower() == ".xml":
            df = extrair_dataframe(arq)
        else:
            xml = arq.with_name(arq.stem + "_corrigido.xml")
            df = extrair_dataframe(xml) if corrigir_ofx_para_xml(arq, xml) else None

        if df is not None and not df.empty:
            df.copy().to_excel(writer, sheet_name=arq.stem[:31], index=False)
            df_consol = df.copy()
            df_consol["ARQUIVO"] = arq.name
            dfs.append(df_consol)

    if dfs:
        total = pd.concat(dfs, ignore_index=True)
        antes = len(total)
        total = total.drop_duplicates(subset=["FITID"])
        print(f"🔁 Deduplicação: {antes - len(total)} removidos")
        total = total.sort_values("DATA")

        ordem = ["ARQUIVO", "FITID", "DATA", "VALOR", "TIPO",
                 "HISTORICO", "DOCUMENTO", "CREDITO", "DEBITO"]
        total[ordem].to_excel(writer, sheet_name="CONSOLIDADO", index=False)

    writer.close()
    load_workbook(saida).close()
    print(f"\n✅ Arquivo gerado: {saida}")

# ============================================================
# 4. EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    process_dir()
