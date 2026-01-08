# ============================================================
# PROCESSADOR DE ARQUIVOS OFX/XML ➜ XLSX CONSOLIDADO
# (inclui reparo padrão automático de OFX quebrado)
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
    """Troca vírgula por ponto e limpa espaços."""
    if valor is None:
        return ""
    v = valor.strip().replace(".", "").replace(",", ".") if re.match(r'^[\d\.\,]+$', valor.strip()) else valor.strip()
    # tenta remover múltiplos pontos indevidos mantendo último como separador decimal
    if v.count('.') > 1 and re.search(r'\d+\.\d+$', v):
        parts = v.split('.')
        v = ''.join(parts[:-1]) + '.' + parts[-1]
    return v

def _normalize_date(dt_text: str) -> str:
    """Tenta extrair YYYYMMDD (ou YYYYMMDDHHMMSS) e corrige datas absurdas.
       Retorna string no mesmo formato encontrado (YYYYMMDD...).
    """
    if not dt_text:
        return ""
    # extrai sequência de dígitos de 8 a 14 chars
    m = re.search(r'(\d{8,14})', dt_text)
    if not m:
        return ""
    s = m.group(1)
    try:
        year = int(s[0:4])
        # se ano absurdo (muito antigo ou futuro > ano_atual+1), corrige para data atual
        ano_atual = datetime.now().year
        if year < 1900 or year > ano_atual + 1:
            return datetime.now().strftime("%Y%m%d%H%M%S")
        return s
    except:
        return datetime.now().strftime("%Y%m%d%H%M%S")


# ============================================================
# 1. CORRIGE/RECONSTRÓI OFX QUEBRADO (REPARO PADRÃO - OPÇÃO B)
# ============================================================
def corrigir_ofx_para_xml(caminho_ofx: Path, caminho_xml: Path) -> bool:
    """
    Reparo padrão:
    - tenta extrair blocos existentes (BANKACCTFROM, BANKTRANLIST, LEDGERBAL, STMTTRN)
    - normaliza decimais e datas
    - gera FITID quando vazio
    - reconstrói a estrutura OFX mínima e salva como XML parseável
    """
    try:
        texto = caminho_ofx.read_text(encoding="latin1", errors="ignore")
    except Exception as e:
        print(f"Erro lendo {caminho_ofx.name}: {e}")
        return False

    # 1) Remove eventual header OFX key:val até chegar em tag '<'
    # Alguns arquivos têm cabeçalho com linhas A: B ; remover até primeira linha com '<'
    idx_first_tag = re.search(r'<', texto)
    if idx_first_tag:
        texto_body = texto[idx_first_tag.start():]
    else:
        texto_body = texto

    # 2) Extrair blocos se existirem (captura como strings brutas)
    # Busca bloco BANKACCTFROM
    m_bankacct = re.search(r'<BANKACCTFROM>(.*?)</BANKACCTFROM>', texto_body, flags=re.S | re.I)
    bankacct = m_bankacct.group(0) if m_bankacct else ""

    # Busca LEDGERBAL (balanço)
    m_ledger = re.search(r'<LEDGERBAL>(.*?)</LEDGERBAL>', texto_body, flags=re.S | re.I)
    ledgerbal = m_ledger.group(0) if m_ledger else ""

    # Busca conteúdo de BANKTRANLIST (que contém DTSTART, DTEND e STMTTRN)
    m_banktran = re.search(r'<BANKTRANLIST>(.*?)</BANKTRANLIST>', texto_body, flags=re.S | re.I)
    banktran_inner = m_banktran.group(1) if m_banktran else ""

    # A partir do inner, buscar todos os STMTTRN
    stmts_raw = re.findall(r'<STMTTRN>(.*?)</STMTTRN>', banktran_inner, flags=re.S | re.I)
    stmt_entries = []
    for raw in stmts_raw:
        # pra cada entrada, vamos extrair campos internos como tags simples
        campos = {}
        for tag in ["TRNTYPE", "DTPOSTED", "TRNAMT", "FITID", "MEMO", "CHECKNUM"]:
            rx = re.search(rf'<{tag}>(.*?)($|<)', raw, flags=re.S | re.I)
            if rx:
                campos[tag] = rx.group(1).strip()
            else:
                campos[tag] = ""
        # Normalizações:
        campos["TRNAMT"] = _normalize_decimal(campos.get("TRNAMT", ""))
        campos["DTPOSTED"] = _normalize_date(campos.get("DTPOSTED", ""))
        # FITID automático se vazio
        if not campos.get("FITID"):
            campos["FITID"] = uuid.uuid4().hex
        stmt_entries.append(campos)

    # Se não encontramos STMTTRN via regex, tentaremos localizar linhas que pareçam transações soltas
    if not stmt_entries:
        # tentativa de captar linhas que contenham DTPOSTED ou TRNAMT
        possible = re.findall(r'(<STMTTRN>.*?</STMTTRN>)', texto_body, flags=re.S | re.I)
        for raw in possible:
            campos = {}
            for tag in ["TRNTYPE", "DTPOSTED", "TRNAMT", "FITID", "MEMO", "CHECKNUM"]:
                rx = re.search(rf'<{tag}>(.*?)($|<)', raw, flags=re.S | re.I)
                if rx:
                    campos[tag] = rx.group(1).strip()
                else:
                    campos[tag] = ""
            campos["TRNAMT"] = _normalize_decimal(campos.get("TRNAMT", ""))
            campos["DTPOSTED"] = _normalize_date(campos.get("DTPOSTED", ""))
            if not campos.get("FITID"):
                campos["FITID"] = uuid.uuid4().hex
            stmt_entries.append(campos)

    # Se ainda vazio, tentar extrair manualmente por linhas que contenham algo parecido com valor/date
    if not stmt_entries:
        # busca por linhas com padrão de data YYYYMMDD e valor com vírgula ou ponto
        lines = texto_body.splitlines()
        cur = {}
        for ln in lines:
            # data
            mdate = re.search(r'(\d{8,14})', ln)
            mval = re.search(r'(-?\d+[\.,]\d{2})', ln)
            if mdate and mval:
                cur = {}
                cur["DTPOSTED"] = _normalize_date(mdate.group(1))
                cur["TRNAMT"] = _normalize_decimal(mval.group(1))
                cur["FITID"] = uuid.uuid4().hex
                # tentativa de memo: resto da linha
                memo = re.sub(r'(\d{8,14})', '', ln)
                memo = re.sub(r'(-?\d+[\.,]\d{2})', '', memo).strip()
                cur["MEMO"] = memo
                cur["TRNTYPE"] = "UNKNOWN"
                cur["CHECKNUM"] = ""
                stmt_entries.append(cur)

    # Caso não tenhamos nenhuma transação, o reparo falha
    if not stmt_entries:
        print(f"❌ Não foi possível localizar transações em: {caminho_ofx.name}")
        return False

    # 3) Reconstruir DTSTART/DTEND a partir das transações
    dt_values = [s["DTPOSTED"] for s in stmt_entries if s.get("DTPOSTED")]
    dtstart = min(dt_values) if dt_values else datetime.now().strftime("%Y%m%d%H%M%S")
    dtend = max(dt_values) if dt_values else datetime.now().strftime("%Y%m%d%H%M%S")

    # 4) Normalizar ledgerbal se existir
    bal_amt = ""
    if ledgerbal:
        m_bal = re.search(r'<BALAMT>(.*?)($|<)', ledgerbal, flags=re.S | re.I)
        if m_bal:
            bal_amt = _normalize_decimal(m_bal.group(1))
        else:
            bal_amt = ""

    # 5) Montar XML OFX reconstruído
    # montar BANKACCTFROM: se não existe, criar placeholders
    if not bankacct:
        bankacct = (
            "<BANKACCTFROM>\n"
            "  <BANKID>UNKNOWN</BANKID>\n"
            "  <BRANCHID>UNKNOWN</BRANCHID>\n"
            "  <ACCTID>UNKNOWN</ACCTID>\n"
            "  <ACCTTYPE>CHECKING</ACCTTYPE>\n"
            "</BANKACCTFROM>"
        )

    # montar transações
    stmts_xml = []
    for s in stmt_entries:
        trn = [
            "<STMTTRN>",
            f"<TRNTYPE>{(s.get('TRNTYPE') or '').strip() or 'OTHER'}</TRNTYPE>",
            f"<DTPOSTED>{s.get('DTPOSTED') or datetime.now().strftime('%Y%m%d%H%M%S')}</DTPOSTED>",
            f"<TRNAMT>{(s.get('TRNAMT') or '').replace(',', '.')}</TRNAMT>",
            f"<FITID>{s.get('FITID') or uuid.uuid4().hex}</FITID>",
        ]
        # MEMO and CHECKNUM optional
        memo = s.get('MEMO', '').strip()
        if memo:
            trn.append(f"<MEMO>{memo}</MEMO>")
        check = s.get('CHECKNUM', '').strip()
        if check:
            trn.append(f"<CHECKNUM>{check}</CHECKNUM>")
        trn.append("</STMTTRN>")
        stmts_xml.append("\n".join(trn))

    banktranlist_block = (
        "<BANKTRANLIST>\n"
        f"<DTSTART>{dtstart}</DTSTART>\n"
        f"<DTEND>{dtend}</DTEND>\n"
        + "\n".join(stmts_xml) +
        "\n</BANKTRANLIST>"
    )

    ledgerbal_block = ""
    if bal_amt:
        ledgerbal_block = (
            "<LEDGERBAL>\n"
            f"<BALAMT>{bal_amt.replace(',', '.')}</BALAMT>\n"
            f"<DTASOF>{dtend}</DTASOF>\n"
            "</LEDGERBAL>"
        )

    ofx_recon = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        "<OFX>\n"
        "  <BANKMSGSRSV1>\n"
        "    <STMTTRNRS>\n"
        "      <STMTRS>\n"
        f"        {bankacct}\n"
        f"        {banktranlist_block}\n"
        f"        {ledgerbal_block}\n"
        "      </STMTRS>\n"
        "    </STMTTRNRS>\n"
        "  </BANKMSGSRSV1>\n"
        "</OFX>\n"
    )

    # 6) Salvar e validar
    try:
        caminho_xml.write_text(ofx_recon, encoding="utf-8")
    except Exception as e:
        print(f"Erro salvando XML em {caminho_xml}: {e}")
        return False

    # tenta parse
    try:
        ET.parse(caminho_xml)
        return True
    except Exception as e:
        print(f"❌ Falha ao parsear XML reconstruído: {e}")
        # em caso de falha, salva uma versão para debug
        debug_path = caminho_xml.with_suffix(".debug.xml")
        try:
            debug_path.write_text(ofx_recon, encoding="utf-8")
            print(f"Arquivo de debug salvo em: {debug_path}")
        except:
            pass
        return False


# ============================================================
# 2. Extrair dados de um XML válido
# ============================================================
def extrair_dataframe(caminho_xml: Path):
    try:
        tree = ET.parse(caminho_xml)
    except Exception as e:
        print(f"Erro ao carregar XML {caminho_xml}: {e}")
        return None

    root = tree.getroot()
    # tenta localizar os nós STMTTRN
    trans = root.findall(".//STMTTRN")
    if not trans:
        # algumas variações usam TRANSACTION ou LANCAMENTO
        trans = root.findall(".//TRANSACTION") or root.findall(".//LANCAMENTO")

    if not trans:
        print(f"⚠ Nenhuma transação encontrada em {caminho_xml.name}")
        return None

    rows = []
    for t in trans:
        rows.append({el.tag: (el.text or '').strip() for el in t})

    df = pd.DataFrame(rows)

    # Data
    if "DTPOSTED" in df.columns:
        df["DTPOSTED"] = (
            df["DTPOSTED"]
            .str.extract(r"(\d{8,14})")[0]
            .apply(lambda x: pd.to_datetime(
                x,
                format="%Y%m%d%H%M%S" if len(str(x)) > 8 else "%Y%m%d",
                errors="coerce"
            ))
            .dt.strftime("%d/%m/%Y")
        )

    # Valores
    if "TRNAMT" in df.columns:
        df["TRNAMT"] = df["TRNAMT"].astype(str).str.replace(",", ".", regex=False)
        # remove espaços e possíveis separadores de milhares
        df["TRNAMT"] = df["TRNAMT"].str.replace(r"[^\d\.\-]", "", regex=True)
        df["TRNAMT"] = pd.to_numeric(df["TRNAMT"], errors="coerce")
        df["CREDITO"] = df["TRNAMT"].apply(lambda x: x if x > 0 else "")
        df["DEBITO"] = df["TRNAMT"].apply(lambda x: x if x < 0 else "")
        df["TIPO"] = df["TRNAMT"].apply(lambda x: "C" if x > 0 else ("D" if x < 0 else ""))
    else:
        df["CREDITO"] = ""
        df["DEBITO"] = ""
        df["TIPO"] = ""

    df.rename(columns={
        "DTPOSTED": "DATA",
        "TRNAMT": "VALOR",
        "MEMO": "HISTORICO",
        "CHECKNUM": "DOCUMENTO"
    }, inplace=True)

    cols = ["DATA", "VALOR", "TIPO", "HISTORICO", "DOCUMENTO", "CREDITO", "DEBITO"]
    return df[[c for c in cols if c in df.columns]]


# ============================================================
# 3. Processo principal → seleciona arquivos e gera XLSX
# ============================================================
def process_dir() -> Path:
    print("\n============================================")
    print("   SELECIONE ARQUIVOS OFX OU XML PARA PROCESSAR")
    print("============================================\n")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    arquivos_selecionados = filedialog.askopenfilenames(
        title="Selecione arquivos OFX ou XML",
        filetypes=[("Arquivos OFX", "*.ofx"), ("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]
    )

    if not arquivos_selecionados:
        print("Nenhum arquivo selecionado. Encerrando.")
        return None

    arquivos = [Path(a) for a in arquivos_selecionados]

    pasta_saida = arquivos[0].parent
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    saida = pasta_saida / f"consolidado_ofx_{ts}.xlsx"

    writer = pd.ExcelWriter(saida, engine="openpyxl")
    dfs = []

    for arq in arquivos:
        print(f"➡ Processando: {arq.name}")

        if arq.suffix.lower() == ".xml":
            df = extrair_dataframe(arq)
        else:
            xml_corr = arq.with_name(arq.stem + "_corrigido.xml")
            ok = corrigir_ofx_para_xml(arq, xml_corr)
            if ok:
                df = extrair_dataframe(xml_corr)
            else:
                print(f"❌ Falha ao converter/reparar OFX: {arq.name}")
                df = None

        if df is not None and not df.empty:
            df.to_excel(writer, sheet_name=arq.stem[:31], index=False)
            df["ARQUIVO"] = arq.name
            dfs.append(df)

    if dfs:
        total = pd.concat(dfs, ignore_index=True)
        ordem = ["ARQUIVO", "DATA", "VALOR", "TIPO", "HISTORICO", "DOCUMENTO", "CREDITO", "DEBITO"]
        total = total[[c for c in ordem if c in total.columns]]
        total.to_excel(writer, sheet_name="CONSOLIDADO", index=False)

    writer.close()

    print(f"\nArquivo gerado: {saida}")
    print("🔎 Validando arquivo final...")

    # Validação do XLSX
    try:
        wb = load_workbook(saida)
        wb.close()
        print("✅ Arquivo validado com sucesso! Nenhuma corrupção detectada.")
    except Exception as e:
        print("\n❌ ERRO: O arquivo XLSX gerado está corrompido!")
        print(f"Motivo técnico: {e}")
        print("⚠ Recomenda-se repetir o processamento.\n")

    return saida


# ============================================================
# 4. EXECUÇÃO DIRETA
# ============================================================
if __name__ == "__main__":
    process_dir()
