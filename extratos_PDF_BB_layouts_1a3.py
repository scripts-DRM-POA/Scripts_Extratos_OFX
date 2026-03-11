# -*- coding: utf-8 -*-

import os, glob, re
from datetime import datetime
from tkinter import Tk, filedialog

import fitz  # PyMuPDF
import pandas as pd


# ============================================================
# Util: extrair linhas do PDF
# ============================================================
def extract_text_lines(pdf_path: str):
    doc = fitz.open(pdf_path)
    lines = []
    for pg in doc:
        txt = pg.get_text("text")
        page_lines = [ln.replace("\xa0", " ").strip() for ln in txt.split("\n")]
        lines.extend([ln for ln in page_lines if ln])
    return lines


def norm_space(s: str) -> str:
    return re.sub(r"\s{2,}", " ", s.replace("\xa0", " ").strip())


def money_to_float_br(s: str) -> float:
    # aceita "R$ 1.234,56" ou "1.234,56"
    s = s.replace("R$", "").strip()
    return float(s.replace(".", "").replace(",", "."))


def is_numeric_token(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,25}", s.strip()))


META_ANY = re.compile(
    r"(saldo\s+anterior|saldo\s+do\s+dia|\bS\s*A\s*L\s*D\s*O\b|total\s+aplica|saldos\s+por\s+dia|posição|"
    r"dispon[ií]vel|bloqueado|limite|vencimento|folha|central de atendimento|sac|ouvidoria|relatório emitido|"
    r"este relatório não é um comprovante)",
    re.IGNORECASE
)


# ============================================================
# LAYOUT 1 (Extrato CC): "1.500,00 (-)" numa linha e a DATA na linha seguinte
# ============================================================
RE_VALSIGN_LINE = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})\s+\((?P<pm>[+-])\)\s*$")
RE_DATE_LINE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
RE_DATE_ZERO = re.compile(r"^00/00/0000$")

HEADERS_A = {
    "extrato de conta corrente", "lançamentos", "dia", "lote", "documento",
    "histórico", "valor", "cliente"
}

def parse_layout_valsign_date_next(lines):
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        m = RE_VALSIGN_LINE.match(lines[i])
        if not m:
            i += 1
            continue

        val = money_to_float_br(m.group("val"))
        if m.group("pm") == "-":
            val = -val

        # procurar data logo após
        j = i + 1
        while j < N and not (RE_DATE_LINE.match(lines[j]) or RE_DATE_ZERO.match(lines[j])):
            j += 1
        if j >= N:
            break
        if RE_DATE_ZERO.match(lines[j]):
            i = j + 1
            continue

        date_str = lines[j]
        dt = datetime.strptime(date_str, "%d/%m/%Y")

        # descrição até próximo valor(+/-)
        desc_parts = []
        k = j + 1
        while k < N and not RE_VALSIGN_LINE.match(lines[k]):
            ln = lines[k].strip()
            if not ln or META_ANY.search(ln):
                k += 1
                continue
            low = ln.lower()
            if low in HEADERS_A:
                k += 1
                continue
            if is_numeric_token(ln):
                k += 1
                continue
            desc_parts.append(ln)
            k += 1

        desc = norm_space(" ".join(desc_parts)).strip(" -")
        if desc:
            rows.append([dt.strftime("%d/%m/%Y"), dt.year, dt.month, desc, val, "C" if val > 0 else "D"])

        i = k

    return pd.DataFrame(rows, columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])


# ============================================================
# LAYOUT 2 (tabela): data "dd.mm.aaaa" sozinha + descrição + linhas "X,XX C/D" (saldo vem depois)
# ============================================================
RE_DATE_DOT = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
RE_MONEY_CD_LINE = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})\s+(?P<dc>[CD])$")

HEADERS_B = {
    "origem", "banco", "lote", "saldo - r$", "valor - r$", "documento", "histórico",
    "agência (prefixo/dv)", "conta nº / dv", "posição", "data da emissão", "data lançamento",
    "folha", "data contábil", "correntista", "extrato conta corrente", "nome", "data da abertura",
    "cnpj", "cpf"
}

def parse_layout_dot_table(lines):
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not RE_DATE_DOT.match(lines[i]):
            i += 1
            continue

        dt = datetime.strptime(lines[i].replace(".", "/"), "%d/%m/%Y")

        # bloco até a próxima data
        block = []
        j = i + 1
        while j < N and not RE_DATE_DOT.match(lines[j]):
            block.append(lines[j])
            j += 1

        block_clean = [b.strip() for b in block if b.strip() and not META_ANY.search(b)]

        # primeiro valor C/D do bloco = valor da transação (o saldo geralmente vem depois)
        val = None
        for b in block_clean:
            mv = RE_MONEY_CD_LINE.match(b)
            if mv and mv.group("val") != "0,00":
                val = money_to_float_br(mv.group("val"))
                if mv.group("dc") == "D":
                    val = -val
                break

        if val is None:
            i = j
            continue

        # descrição = linhas não-numéricas, não-cabeçalho, não valor C/D
        desc_parts = []
        for b in block_clean:
            low = b.lower()
            if low in HEADERS_B:
                continue
            if RE_MONEY_CD_LINE.match(b):
                continue
            if is_numeric_token(b):
                continue
            desc_parts.append(b)

        desc = norm_space(" ".join(desc_parts)).strip(" -")
        if desc:
            rows.append([dt.strftime("%d/%m/%Y"), dt.year, dt.month, desc, val, "C" if val > 0 else "D"])

        i = j

    return pd.DataFrame(rows, columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])


# ============================================================
# LAYOUT 3 (Relatório de Pagamentos): "dd/mm/aaaa NOME..." e valor em linha "R$ X"
# ============================================================
RE_REPORT_ROW = re.compile(r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<name>.+)$")
RE_REPORT_VAL = re.compile(r"^R\$\s*(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})$")

def parse_layout_payments_report(lines):
    rows = []
    i = 0
    N = len(lines)

    current_date = None
    desc_buf = []

    while i < N:
        ln = lines[i].strip()
        if not ln:
            i += 1
            continue
        if META_ANY.search(ln) or ln.startswith("____"):
            i += 1
            continue

        mrow = RE_REPORT_ROW.match(ln)
        if mrow:
            current_date = mrow.group("date")
            desc_buf = [norm_space(mrow.group("name"))]
            i += 1
            continue

        mval = RE_REPORT_VAL.match(ln)
        if mval and current_date and desc_buf:
            dt = datetime.strptime(current_date, "%d/%m/%Y")
            val = money_to_float_br(mval.group("val"))
            # relatório de pagamentos => débito
            val = -abs(val)
            desc = norm_space(" ".join(desc_buf))
            rows.append([dt.strftime("%d/%m/%Y"), dt.year, dt.month, desc, val, "D"])
            current_date = None
            desc_buf = []
            i += 1
            continue

        # continuação de descrição do item
        if current_date and desc_buf:
            if re.match(r"^(CNPJ|CPF)\s*:", ln, flags=re.IGNORECASE):
                i += 1
                continue
            if ln.lower().startswith(("bco:", "ag:", "conta:")):
                i += 1
                continue
            if ln.isdigit():
                i += 1
                continue
            desc_buf.append(ln)

        i += 1

    return pd.DataFrame(rows, columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])


# ============================================================
# Auto-detecção: roda os 3 parsers e escolhe o que extrai mais linhas
# ============================================================
def parse_auto(lines):
    parsers = [
        ("bb_valsign_date_next", parse_layout_valsign_date_next),
        ("bb_dot_table", parse_layout_dot_table),
        ("bb_payments_report", parse_layout_payments_report),
    ]
    best_name = None
    best_df = pd.DataFrame()
    for name, fn in parsers:
        df = fn(lines)
        if len(df) > len(best_df):
            best_df = df
            best_name = name
    return best_name, best_df


def listar_pdfs(pasta):
    return sorted(glob.glob(os.path.join(pasta, "*.pdf")) + glob.glob(os.path.join(pasta, "*.PDF")))


def processar_pasta(pasta_pdf: str, saida_xlsx: str = None):
    pdfs = listar_pdfs(pasta_pdf)
    if not pdfs:
        raise FileNotFoundError(f"Não encontrei PDFs em: {pasta_pdf}")

    dfs = []
    erros = []

    print(f"Processando {len(pdfs)} PDF(s) em: {pasta_pdf}")

    for idx, pdf_path in enumerate(pdfs, start=1):
        nome = os.path.basename(pdf_path)
        try:
            print(f"[{idx}/{len(pdfs)}] Lendo: {nome}")
            lines = extract_text_lines(pdf_path)
            layout, df = parse_auto(lines)

            if df.empty:
                erros.append((nome, "Sem lançamentos extraídos (todos os parsers retornaram vazio)."))
                print("    -> AVISO: nenhum lançamento extraído")
                continue

            df.insert(0, "Arquivo", nome)
            df.insert(1, "LayoutDetectado", layout or "")
            dfs.append(df)
            print(f"    -> OK: {len(df)} lançamento(s) | layout={layout}")
        except Exception as e:
            erros.append((nome, str(e)))
            print(f"    -> ERRO: {e}")

    if not dfs:
        raise ValueError("Nenhum lançamento foi extraído de qualquer PDF (todos vazios/erro).")

    df_all = pd.concat(dfs, ignore_index=True)

    if saida_xlsx is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        saida_xlsx = os.path.join(pasta_pdf, f"consolidado_lancamentos_{ts}.xlsx")

    print(f"Gerando Excel: {saida_xlsx}")

    with pd.ExcelWriter(saida_xlsx, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Lancamentos")
        wb = writer.book
        ws = writer.sheets["Lancamentos"]
        money_fmt = wb.add_format({"num_format": "R$ #,##0.00; -R$ #,##0.00"})
        date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

        ws.set_column("A:A", 40)            # Arquivo
        ws.set_column("B:B", 22)            # LayoutDetectado
        ws.set_column("C:C", 12, date_fmt)  # Data
        ws.set_column("D:D", 6)             # Ano
        ws.set_column("E:E", 5)             # Mês
        ws.set_column("F:F", 110)           # Descrição
        ws.set_column("G:G", 16, money_fmt) # Valor
        ws.set_column("H:H", 5)             # Tipo

        if erros:
            pd.DataFrame(erros, columns=["Arquivo", "Erro"]).to_excel(writer, index=False, sheet_name="Erros")

    print("Concluído.")
    if erros:
        print(f"Atenção: {len(erros)} arquivo(s) com erro (ver aba 'Erros').")

    return df_all, saida_xlsx


if __name__ == "__main__":
    Tk().withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs (Banco do Brasil)")
    if not pasta:
        raise SystemExit("Nenhuma pasta selecionada.")

    _, arquivo_saida = processar_pasta(pasta)
    print("Arquivo gerado:", arquivo_saida)