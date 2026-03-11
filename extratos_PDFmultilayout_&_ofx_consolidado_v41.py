# -*- coding: utf-8 -*-
"""
Extrator multi-layout de extratos bancários (PDF + OFX, sem OCR / sem Tesseract) - v41

Uso:
    python extrator_multilayout_consolidado_v41.py
ou
    python extrator_multilayout_consolidado_v41.py --pasta "C:\\caminho\\dos\\pdfs"

Gera um XLSX com duas abas:
    - Consolidado
    - Logs

Colunas do Consolidado:
    Arquivo | Data | Descrição | Documento | Valor | Tipo | Débito | Crédito

Colunas do Logs:
    Arquivo | n_transações_obtidas
"""

import sys
import os
import re
import glob
import argparse
from datetime import datetime

try:
    import fitz  # PyMuPDF
except Exception:
    print("\nERRO: biblioteca 'pymupdf' não está instalada.")
    print("Instale com:\n    pip install pymupdf\n")
    input("Pressione ENTER para sair...")
    sys.exit()

try:
    import pandas as pd
except Exception:
    print("\nERRO: biblioteca 'pandas' não está instalada.")
    print("Instale com:\n    pip install pandas\n")
    input("Pressione ENTER para sair...")
    sys.exit()


MESES_PT = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}
MESES_BAN = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}


def norm_space(s: str) -> str:
    return re.sub(r"\s{2,}", " ", s.replace("\xa0", " ").replace("\uf166", " ").replace("\ue90a", " ").replace("\uf18f", " ").strip())


def money_to_float(tok: str) -> float:
    t = tok.replace("R$", "").strip().replace("−", "-")
    neg = False
    if t.endswith("-"):
        neg = True
        t = t[:-1]
    if t.startswith("-"):
        neg = True
        t = t[1:]
    val = float(t.replace(".", "").replace(",", "."))
    return -val if neg else val


DOC_LABELS = {"documento", "doc", "nº documento", "no documento", "nr documento", "nro. documento", "nro documento"}
LOT_LABELS = {"lote", "ag. origem", "ag origem", "origem", "banco"}


def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def is_probable_document_token(tok: str) -> bool:
    t = norm_space(tok)
    d = only_digits(t)
    if not d or len(d) <= 6:
        return False
    if len(set(d)) == 1:
        return False
    return bool(re.fullmatch(r"[\d./-]+", t))


def clean_document_token(tok: str) -> str:
    t = norm_space(tok)
    return t if is_probable_document_token(t) else ""


def clean_document_token_flexible(tok: str) -> str:
    t = norm_space(tok)
    if not t:
        return ""
    if re.fullmatch(r"[\d./-]{4,}", t):
        d = only_digits(t)
        if d and len(set(d)) > 1:
            return t
    return ""


def clean_document_token_sicredi(tok: str) -> str:
    t = norm_space(tok)
    if not t:
        return ""
    up = t.upper()
    if up in {"TARIFA", "FGTS", "DAS", "DEPOSI", "CAPTACAO", "PIX_DEB", "SEG_ICATU"}:
        return t
    if re.fullmatch(r"[A-Z0-9_./-]{3,25}", up):
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", t):
            return ""
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", t):
            return ""
        return t
    return ""


def extract_document_from_block_sicredi(block):
    normalized = [norm_space(b) for b in block if norm_space(b)]
    for idx, b in enumerate(normalized):
        if b.lower() in {"documento", "doc"}:
            for cand in normalized[idx + 1: idx + 4]:
                cleaned = clean_document_token_sicredi(cand)
                if cleaned:
                    return cleaned
    candidates = []
    for idx, b in enumerate(normalized):
        cleaned = clean_document_token_sicredi(b)
        if cleaned:
            candidates.append((idx, cleaned))
    if not candidates:
        return ""
    return candidates[-1][1]


def is_balance_or_summary_line(s: str) -> bool:
    low = norm_space(s).lower()
    if not low:
        return True
    return any(x in low for x in [
        "saldo anterior", "saldo do dia", "saldo final", "saldo total", "saldo disponível", "saldo disponivel",
        "saldo bloqueado", "saldo em c/c", "saldo da conta", "saldo na data", "saldo após", "saldo apos",
        "saldo de conta corrente", "saldo por transação", "saldo por transacao", "saldo dev", "saldo cred", "resumo", "a transportar",
        "totalizador", "totais"
    ])


def normalize_text_for_dedupe(s: str) -> str:
    s = norm_space(str(s)).lower()
    s = re.sub(r"[^\w]+", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()


def extract_document_from_block(block, doc_labels=None):
    doc_labels = {x.lower() for x in (doc_labels or DOC_LABELS)}
    normalized = [norm_space(b) for b in block if norm_space(b)]

    for idx, b in enumerate(normalized):
        if b.lower() in doc_labels:
            for cand in normalized[idx + 1: idx + 5]:
                cleaned = clean_document_token(cand)
                if cleaned:
                    return cleaned

    for b in normalized:
        m = re.search(r"(?:documento|doc)\s*[:.-]?\s*([\d./-]{4,})", b, re.I)
        if m:
            cleaned = clean_document_token(m.group(1))
            if cleaned:
                return cleaned

    candidates = []
    for idx, b in enumerate(normalized):
        low = b.lower()
        if low in LOT_LABELS or low in doc_labels:
            continue
        cleaned = clean_document_token(b)
        if cleaned:
            candidates.append((len(only_digits(cleaned)), idx, cleaned))

    if not candidates:
        return ""
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return candidates[0][2]


def standardize(df: pd.DataFrame, doc_cleaner=clean_document_token) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    out = df.copy()
    out["Data"] = out["Data"].astype(str).str.strip()
    out["Descrição"] = out["Descrição"].astype(str).map(norm_space).str.strip(" -")
    if "Documento" not in out.columns:
        out["Documento"] = ""
    out["Documento"] = out["Documento"].fillna("").astype(str).map(doc_cleaner)
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")

    out = out[out["Valor"].notna()]
    out = out[out["Data"].str.match(r"^\d{2}/\d{2}/\d{4}$", na=False)]
    out = out[out["Valor"] != 0]
    out = out[~out["Descrição"].map(is_balance_or_summary_line)]

    bad = re.compile(
        r"saldo anterior|saldo do dia|saldo final|saldo total|saldo disponível|saldo disponivel|saldo bloqueado|"
        r"solicitado em:|fale com a gente|ouvidoria|sac:|cpf/cnpj:|instituição:|instituicao:|agência:|agencia:|"
        r"conta:|período:|periodo:|filtros aplicados|relatório gerado em|relatorio gerado em|extrato financeiro|"
        r"tipo de saldo|tipo de transação|tipo de transacao|a transportar|versão |versao |extrato consolidado inteligente|"
        r"internet banking empresarial|consultas, informações|consultas, informacoes|redes sociais|resumo - |"
        r"saldo de conta corrente em|movimentação|movimentacao|lançamentos|lancamentos$|saldo dev|saldo cred",
        re.IGNORECASE
    )
    out = out[~out["Descrição"].str.contains(bad, na=False)]

    out["Tipo"] = out["Valor"].apply(lambda x: "C" if x > 0 else "D")
    out["Débito"] = out["Valor"].apply(lambda x: x if x < 0 else "")
    out["Crédito"] = out["Valor"].apply(lambda x: x if x > 0 else "")

    return out[["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"]]

def extract_lines(pdf_path: str):
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        txt = page.get_text("text") or ""
        for ln in txt.splitlines():
            ln = norm_space(ln)
            if ln:
                lines.append(ln)
    return lines


def list_input_files(folder: str):
    seen = {}
    patterns = ["*.pdf", "*.PDF", "*.ofx", "*.OFX"]
    for pattern in patterns:
        for p in glob.glob(os.path.join(folder, pattern)):
            key = os.path.abspath(p).lower()
            if key not in seen:
                seen[key] = os.path.abspath(p)
    return sorted(seen.values())


def normalize_ofx_date(dt: str) -> str:
    if not dt:
        return ""
    m = re.search(r"(\d{8})", dt)
    if not m:
        return ""
    s = m.group(1)
    try:
        return datetime.strptime(s, "%Y%m%d").strftime("%d/%m/%Y")
    except Exception:
        return ""


def parse_ofx_file(ofx_path):
    try:
        texto = open(ofx_path, "r", encoding="latin1", errors="ignore").read()
    except Exception:
        texto = open(ofx_path, "r", encoding="utf-8", errors="ignore").read()

    blocos = re.findall(r"<STMTTRN>(.*?)</STMTTRN>", texto, flags=re.S | re.I)
    rows = []

    def campo(raw: str, tag: str) -> str:
        m = re.search(rf"<{tag}>(.*?)(?:$|<)", raw, flags=re.I | re.S)
        return norm_space(m.group(1)) if m else ""

    for raw in blocos:
        data = normalize_ofx_date(campo(raw, "DTPOSTED"))
        valor_txt = campo(raw, "TRNAMT")
        if not data or not valor_txt:
            continue

        try:
            valor = float(str(valor_txt).replace(",", "."))
        except Exception:
            continue

        descricao = campo(raw, "MEMO") or campo(raw, "NAME") or campo(raw, "TRNTYPE") or "Lançamento OFX"
        documento = campo(raw, "CHECKNUM") or campo(raw, "REFNUM") or campo(raw, "FITID")

        rows.append([data, descricao, documento, valor])

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
                      doc_cleaner=lambda x: norm_space(str(x))[:80])


# ---------------- Banco do Brasil ----------------

def parse_bb_layout1(lines):
    re_val = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})\s+\((?P<pm>[+-])\)\s*$")
    re_date = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        m = re_val.match(lines[i])
        if not m:
            i += 1
            continue

        val = money_to_float(m.group("val"))
        val = -abs(val) if m.group("pm") == "-" else abs(val)

        j = i + 1
        while j < N and not re_date.match(lines[j]):
            j += 1
        if j >= N:
            break

        dt = lines[j]
        desc_parts = []
        doc = ""
        k = j + 1

        while k < N and not re_val.match(lines[k]):
            ln = lines[k]
            if re_date.match(ln):
                break
            if ln.lower() in {"extrato de conta corrente", "lançamentos", "dia", "lote", "documento", "histórico", "valor", "cliente"}:
                k += 1
                continue
            if re.fullmatch(r"\d{4,}", ln):
                if not doc:
                    doc = ln
                k += 1
                continue
            if "saldo" in ln.lower():
                k += 1
                continue
            desc_parts.append(ln)
            k += 1

        desc = norm_space(" ".join(desc_parts))
        if desc:
            rows.append([dt, desc, doc, val])

        i = k

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]), doc_cleaner=clean_document_token_flexible)


def parse_bb_layout2(lines):
    re_date = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    re_money_cd = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})\s+(?P<dc>[CD])$")
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not re_date.match(lines[i]):
            i += 1
            continue

        dt = datetime.strptime(lines[i].replace(".", "/"), "%d/%m/%Y").strftime("%d/%m/%Y")
        block = []
        j = i + 1
        while j < N and not re_date.match(lines[j]):
            if lines[j]:
                block.append(lines[j].strip())
            j += 1

        if not block or any("saldo anterior" in b.lower() for b in block):
            i = j
            continue

        val = None
        val_idx = None
        for idx_b, b in enumerate(block):
            m = re_money_cd.match(b)
            if m and m.group("val") != "0,00":
                val = money_to_float(m.group("val"))
                val = -abs(val) if m.group("dc") == "D" else abs(val)
                val_idx = idx_b
                break

        if val is None:
            i = j
            continue

        pre_value_block = block[:val_idx]
        doc = extract_document_from_block(pre_value_block)
        desc_parts = []
        for b in pre_value_block:
            low = b.lower()

            if low in {
                "origem", "banco", "lote", "saldo - r$", "valor - r$", "documento", "histórico",
                "agência (prefixo/dv)", "conta nº / dv", "posição", "data da emissão", "data lançamento",
                "folha", "data contábil", "correntista", "extrato conta corrente", "nome", "data da abertura",
                "cnpj", "cpf"
            }:
                continue

            if clean_document_token(b) == doc:
                continue

            if re.fullmatch(r"\d{1,6}", b) or re.fullmatch(r"[\d\.]{1,10}", b):
                continue

            desc_parts.append(b)

        desc = norm_space(" ".join(desc_parts))
        if desc and desc != "0,00":
            rows.append([dt, desc, doc, val])

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))

def parse_bb_layout3(lines):
    re_date = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    re_amt = re.compile(r"^(?P<amt>\d{1,3}(?:\.\d{3})*,\d{2})\s+(?P<dc>[CD])$")
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not re_date.match(lines[i]):
            i += 1
            continue

        dt = datetime.strptime(lines[i].replace(".", "/"), "%d/%m/%Y").strftime("%d/%m/%Y")
        j = i + 1

        if j < N and re.fullmatch(r"\d{1,4}", lines[j]):
            j += 1  # coluna Origem opcional
            while j < N and lines[j].lower() in {"origem", "histórico", "historico", "documento", "lote"}:
                j += 1

        if j >= N:
            break

        hist = norm_space(lines[j])
        if is_balance_or_summary_line(hist):
            i = j + 1
            continue
        j += 1

        doc = ""
        if j < N and re.fullmatch(r"[A-Za-z0-9./-]{3,25}", lines[j]):
            cand = norm_space(lines[j])
            if j + 1 < N and re.fullmatch(r"\d{3,6}", lines[j + 1]):
                doc = clean_document_token_flexible(cand) or cand
                j += 1
            elif j + 1 < N and re_amt.match(lines[j + 1]):
                doc = clean_document_token_flexible(cand)
                j += 1

        if j < N and re.fullmatch(r"\d{3,6}", lines[j]):
            j += 1
        if j >= N:
            break

        m = re_amt.match(lines[j])
        if not m:
            i += 1
            continue

        val = money_to_float(m.group("amt"))
        val = -abs(val) if m.group("dc") == "D" else abs(val)

        rows.append([dt, hist, doc, val])
        i = j + 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]), doc_cleaner=clean_document_token_flexible)

def parse_bb_layout4(lines):
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    val_line_re = re.compile(
        r"^(?:(?P<doc>[\d./-]{3,25})\s+)?(?P<amt>\d{1,3}(?:\.\d{3})*,\d{2})\s+(?P<dc>[CD])"
        r"(?:\s+\d{1,3}(?:\.\d{3})*,\d{2}\s+[CD])?$"
    )
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not date_re.match(lines[i]):
            i += 1
            continue

        dt = lines[i]
        j = i + 1
        while j < N and date_re.match(lines[j]):
            dt = lines[j]
            j += 1

        block = []
        while j < N and not date_re.match(lines[j]):
            ln = lines[j].strip()
            if ln:
                block.append(ln)
            j += 1

        if not block or "saldo anterior" in " ".join(block).lower():
            i = j
            continue

        val = None
        val_idx = None
        doc = ""
        for idx_b, b in enumerate(block):
            m = val_line_re.match(b)
            if m and m.group("amt") != "0,00":
                val = money_to_float(m.group("amt"))
                val = -abs(val) if m.group("dc") == "D" else abs(val)
                val_idx = idx_b
                doc = clean_document_token_flexible(m.group("doc") or "")
                break

        if val is None:
            i = j
            continue

        pre_value_block = block[:val_idx]
        post_value_block = block[val_idx + 1:]

        if not doc:
            doc = extract_document_from_block(pre_value_block)
        if not doc and pre_value_block:
            tail = norm_space(pre_value_block[-1])
            if re.fullmatch(r"[\d./-]{3,25}", tail):
                doc = clean_document_token_flexible(tail) or tail

        desc_parts = []
        for b in pre_value_block + post_value_block:
            low = b.lower()
            if "saldo" in low:
                continue
            if low in {"lançamentos", "dt. balancete", "dt. movimento ag. origem lote histórico", "documento", "valor r$", "saldo"}:
                continue
            if b == "0000":
                continue
            if doc and norm_space(b) == doc:
                continue
            if re.fullmatch(r"[\d./-]{3,25}", b) and (clean_document_token_flexible(b) or b == doc):
                continue
            b2 = re.sub(r"^\d+\s+\d+\s+", "", b).strip()
            b2 = re.sub(r"\s+\d{1,3}(?:\.\d{3})*,\d{2}\s+[CD](?:\s+\d{1,3}(?:\.\d{3})*,\d{2}\s+[CD])?$", "", b2).strip()
            if doc and b2.endswith(doc):
                b2 = b2[:-len(doc)].strip(' -')
            if b2 and not is_balance_or_summary_line(b2):
                desc_parts.append(b2)

        desc = norm_space(" ".join(desc_parts))
        if desc and "saldo" not in desc.lower() and desc.replace(" ", "").lower() != "saldo":
            rows.append([dt, desc, doc, val])

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]), doc_cleaner=clean_document_token_flexible)

def parse_bb_payments_report(lines):
    re_row = re.compile(r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<name>.+)$")
    re_val = re.compile(r"^R\$\s*(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})$")
    rows = []
    cur = None
    buf = []
    i = 0
    N = len(lines)

    while i < N:
        ln = lines[i]
        m = re_row.match(ln)
        if m:
            cur = m.group("date")
            buf = [m.group("name")]
            i += 1
            continue

        mv = re_val.match(ln)
        if mv and cur and buf:
            rows.append([cur, norm_space(" ".join(buf)), "", -abs(money_to_float(mv.group("val")))])
            cur = None
            buf = []
            i += 1
            continue

        if cur:
            if re.match(r"^(CNPJ|CPF)\s*:", ln, re.I) or ln.lower().startswith(("bco:", "ag:", "conta:")) or ln.isdigit():
                i += 1
                continue
            buf.append(ln)

        i += 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


def parse_bb_auto(lines):
    candidates = [
        ("bb_layout1", parse_bb_layout1(lines)),
        ("bb_layout2", parse_bb_layout2(lines)),
        ("bb_layout3", parse_bb_layout3(lines)),
        ("bb_layout4", parse_bb_layout4(lines)),
        ("bb_report", parse_bb_payments_report(lines)),
    ]
    return max(candidates, key=lambda t: len(t[1]))


# ---------------- ABC ----------------

def parse_abc(lines):
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not date_re.match(lines[i]):
            i += 1
            continue

        dt = lines[i]
        if i + 2 < N and "SALDO ANTERIOR" in lines[i + 2].upper():
            i += 1
            continue

        doc = "" if i + 1 >= N or lines[i + 1] == "-" else lines[i + 1]
        j = i + 2
        desc_parts = []
        val = None

        while j < N:
            ln = lines[j]
            if date_re.match(ln):
                break

            if ln.lower() in {"credito", "crédito", "debito", "débito"}:
                k = j + 1
                while k < N and (lines[k] == "-" or not money_re.match(lines[k])):
                    if date_re.match(lines[k]):
                        break
                    k += 1
                if k < N and money_re.match(lines[k]):
                    val = money_to_float(lines[k])
                j = k + 1
                break

            if ln.lower() in {"data", "nro. documento", "histórico", "historico", "operação", "operacao", "valor (r$)", "saldo diário (r$)", "saldo diario (r$)"} or ln == "-":
                j += 1
                continue

            desc_parts.append(ln)
            j += 1

        if val is not None:
            rows.append([dt, norm_space(" ".join(desc_parts)), doc, val])

        i = j if j > i else i + 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


# ---------------- Banrisul ----------------

def parse_banrisul(lines):
    mes = None
    ano = None
    for ln in lines:
        if "PERIODO:" in ln.upper() or "PERÍODO:" in ln.upper():
            try:
                periodo = ln.split(":", 1)[1].strip()
                nome_mes, a = periodo.split("/")
                ano = int(re.sub(r"\D", "", a))
                mes = MESES_BAN.get(nome_mes.strip().upper())
            except Exception:
                pass
            break

    if not (mes and ano):
        return standardize(pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor"]))

    rows = []
    dia = None
    padrao = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}-?$")

    for ln in lines:
        t = ln.strip()
        if not t:
            continue
        up = t.upper()
        if any(k in up for k in ["SALDO ANT", "SALDO NA DATA", "MOVIMENTOS", "DIA HISTORICO"]):
            continue

        m = re.match(r"^(\d{2})\s+(.*)", t)
        if m:
            dia = int(m.group(1))
            restante = m.group(2)
        else:
            if dia is None:
                continue
            restante = t

        parts = restante.split()
        if len(parts) < 2 or not padrao.match(parts[-1]):
            continue

        val = money_to_float(parts[-1])
        if "saldo" in restante.lower():
            continue

        doc = ""
        desc_tokens = parts[:-1]
        if desc_tokens and clean_document_token(desc_tokens[-1]):
            doc = clean_document_token(desc_tokens[-1])
            desc_tokens = desc_tokens[:-1]
        desc = " ".join(desc_tokens)

        dt = datetime(ano, mes, dia).strftime("%d/%m/%Y")
        rows.append([dt, desc, doc, val])

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))




# ---------------- Sicredi ----------------

def parse_sicredi(lines):
    rows = []
    i = 0
    N = len(lines)
    date_line = re.compile(r'^(?P<date>\d{2}/\d{2}/\d{4})(?:\s+(?P<rest>.+))?$')
    money_re = re.compile(r'^-?\d{1,3}(?:\.\d{3})*,\d{2}$')
    money_pair_re = re.compile(r'^(?P<amt>-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(?P<saldo>-?\d{1,3}(?:\.\d{3})*,\d{2})$')
    headers = {"Data", "Descrição", "Documento", "Valor (R$)", "Saldo (R$)", "Extrato", "SALDO", "SALDO ANTERIOR"}
    skip_prefixes = ("Associado:", "Cooperativa:", "Conta Corrente:", "Conta:", "Dados referentes ao período", "Extrato (Período", "Impresso em", "Sicredi Fone", "SAC ", "Ouvidoria")

    while i < N:
        ln = lines[i]
        if ln in headers or any(ln.startswith(p) for p in skip_prefixes):
            i += 1
            continue
        if ln in {"SALDO", "SALDO ANTERIOR"}:
            i += 2
            continue

        m = date_line.match(ln)
        if not m:
            i += 1
            continue

        dt = m.group('date')
        rest = (m.group('rest') or '').strip()

        j = i + 1
        block = [rest] if rest else []
        while j < N:
            cur = lines[j]
            if cur in headers or any(cur.startswith(p) for p in skip_prefixes):
                j += 1
                continue
            if cur in {"SALDO", "SALDO ANTERIOR"}:
                break
            if date_line.match(cur):
                break
            block.append(cur)
            j += 1

        block = [b for b in block if b and not is_balance_or_summary_line(b)]

        amt = None
        amt_idx = None
        for idx, b in enumerate(block):
            mp = money_pair_re.match(b)
            if mp:
                amt = money_to_float(mp.group('amt'))
                amt_idx = idx
                break
            if money_re.match(b):
                amt = money_to_float(b)
                amt_idx = idx
                break

        if amt is None:
            i = j
            continue

        pre_value_block = block[:amt_idx]
        doc = extract_document_from_block_sicredi(pre_value_block)
        desc_parts = []
        for b in pre_value_block:
            if doc and norm_space(b) == doc:
                continue
            desc_parts.append(b)

        desc = norm_space(" ".join(desc_parts))
        if desc and not is_balance_or_summary_line(desc):
            rows.append([dt, desc, doc, amt])

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]), doc_cleaner=clean_document_token_sicredi)

# ---------------- Inter ----------------

RE_INTER_DAY1 = re.compile(r"^(Segunda|Terça|Terca|Quarta|Quinta|Sexta|Sábado|Sabado|Domingo),\s+(\d{1,2})\s+de\s+([A-Za-zç]+)\s+de\s+(\d{4})$", re.I)
RE_INTER_DAY2 = re.compile(r"^(\d{1,2})\s+de\s+([A-Za-zç]+)\s+de\s+(\d{4})(?:\s+Saldo do dia:.*)?$", re.I)
RE_INTER_VAL = re.compile(r"^[+-]?R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}$")
RE_INTER_BAL = re.compile(r"^R\$\s*[-−]?\d{1,3}(?:\.\d{3})*,\d{2}$")

def parse_inter(lines):
    rows = []
    current = None
    desc_buf = []
    i = 0
    N = len(lines)

    meta_starts = (
        "solicitado em:", "cpf/cnpj:", "instituição:", "agência:", "conta:", "período",
        "saldo total", "saldo disponível", "saldo bloqueado", "fale com a gente", "sac",
        "ouvidoria", "deficiência"
    )

    while i < N:
        ln = lines[i]
        low = ln.lower()

        if any(low.startswith(x) for x in meta_starts) or low in {"valor", "saldo por transação", "(bloqueado + disponível)", "(bloqueado + disponivel)"}:
            i += 1
            continue

        m = RE_INTER_DAY1.match(ln) or RE_INTER_DAY2.match(ln)
        if m:
            if len(m.groups()) == 4:
                day = int(m.group(2)); mon = m.group(3); year = m.group(4)
            else:
                day = int(m.group(1)); mon = m.group(2); year = m.group(3)
            current = f"{day:02d}/{MESES_PT[mon.lower()]:02d}/{year}"
            desc_buf = []
            i += 1
            continue

        if current is None:
            i += 1
            continue

        if low.startswith("saldo do dia"):
            i += 1
            continue

        if RE_INTER_VAL.match(ln):
            desc = norm_space(" ".join(desc_buf))
            if desc:
                rows.append([current, desc, "", money_to_float(ln)])
            desc_buf = []
            if i + 1 < N and RE_INTER_BAL.match(lines[i + 1]):
                i += 2
            else:
                i += 1
            continue

        if RE_INTER_BAL.match(ln):
            i += 1
            continue

        if not is_balance_or_summary_line(ln):
            desc_buf.append(ln)
        i += 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))





RE_SANT_DAY = re.compile(r"^(Segunda|Terça|Terca|Quarta|Quinta|Sexta|Sábado|Sabado|Domingo),\s+(\d{1,2})\s+de\s+([A-Za-zç]+)\s+de\s+(\d{4})$", re.I)

# ---------------- Santander ----------------

def parse_santander_layout1_from_pdf(pdf_path):
    import pdfplumber, re, pandas as pd

    texto = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    linhas = [re.sub(r"\s{2,}", " ", l.strip()) for l in texto.splitlines() if l.strip()]

    start_idx = None
    for idx, l in enumerate(linhas):
        if l.strip().lower() == "movimentação":
            start_idx = idx + 1
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    end_idx = len(linhas)
    for idx in range(start_idx, len(linhas)):
        if "saldos por período" in linhas[idx].lower():
            end_idx = idx
            break

    sect = []
    for l in linhas[start_idx:end_idx]:
        low = l.lower()
        if low.startswith("extrato_pj_") or low.startswith("balp_uy_") or low.startswith("pagina:"):
            continue
        if low == "extrato consolidado inteligente" or re.fullmatch(r'[A-Za-zç]+/\d{4}', l, re.I):
            continue
        if low in {
            "data descrição nº documento movimentos (r$) saldo (r$)",
            "créditos débitos", "créditos", "creditos", "débitos", "debitos", "conta corrente", "movimentação"
        }:
            continue
        sect.append(l)

    year_ctx = "2024"
    mhead = re.search(r'([A-Za-zç]+)/(\d{4})', texto, re.I)
    if mhead:
        year_ctx = mhead.group(2)

    re_date = re.compile(r'^(?P<data>\d{2}/\d{2})\s+(?P<rest>.+)$')
    re_tx = re.compile(
        r'^(?P<desc>.+?)'
        r'(?:\s+(?P<doc>\d{4,}|-))?'
        r'\s+(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)'
        r'(?:\s+(?P<saldo>\d{1,3}(?:\.\d{3})*,\d{2}-?))?$'
    )

    rows = []
    current_date = None
    last_idx = None

    for line in sect:
        low = line.lower()
        if low.startswith("saldo em "):
            continue

        m = re_date.match(line)
        if m:
            current_date = f"{m.group('data')}/{year_ctx}"
            rest = m.group('rest').strip()
            mt = re_tx.match(rest)
            if mt:
                desc = mt.group('desc').strip(" -")
                doc = mt.group('doc') or ""
                if doc == "-":
                    doc = ""
                val = money_to_float(mt.group('val'))
                rows.append([current_date, desc, doc, val])
                last_idx = len(rows) - 1
            else:
                rows.append([current_date, rest.strip(" -"), "", None])
                last_idx = len(rows) - 1
            continue

        if current_date is None:
            continue

        mt = re_tx.match(line)
        if mt and re.search(r'[A-Za-zÀ-ÿ]', mt.group('desc') or ''):
            desc = mt.group('desc').strip(" -")
            doc = mt.group('doc') or ""
            if doc == "-":
                doc = ""
            val = money_to_float(mt.group('val'))
            rows.append([current_date, desc, doc, val])
            last_idx = len(rows) - 1
            continue

        if last_idx is not None:
            low2 = line.lower()
            if re.fullmatch(r'\d{1,3}(?:\.\d{3})*,\d{2}-?', line):
                continue
            if low2.startswith("saldo em ") or "saldos por período" in low2 or "saldos por periodo" in low2:
                continue
            if line.startswith('“') or line.startswith('"'):
                continue
            if any(x in low2 for x in ["este demonstrativo", "central de atendimento", "sac", "ouvidoria", "extrato_pj_", "balp_uy_", "adiantamento a depositantes", "jurosmoratórios", "jurosmoratorios", "produtocontratado", "saldo devedor"]):
                continue
            if low2.startswith('periodo:') or low2.startswith('período:'):
                line = re.sub(r'^per[íi]odo:\s*', 'PERIODO ', line, flags=re.I)
            rows[last_idx][1] = norm_space(((rows[last_idx][1] or "") + " " + line).strip())

    df = pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"])
    df = df[df["Valor"].notna()].copy()
    return standardize(df)

def parse_santander_layout2(lines):
    rows = []
    current = None
    i = 0
    N = len(lines)

    while i < N:
        ln = lines[i]
        m = RE_SANT_DAY.match(ln)
        if m:
            current = f"{int(m.group(2)):02d}/{MESES_PT[m.group(3).lower()]:02d}/{m.group(4)}"
            i += 1
            continue

        if current is None:
            i += 1
            continue

        if ln in {"CREDITO", "DEBITO"} or any(x in ln.lower() for x in ["solicitado em", "internet banking empresarial", "exibindo resultados", "para consultas abaixo", "agência:", "conta:", "banco santander"]):
            i += 1
            continue

        if i + 2 < N and lines[i + 1] in {"CREDITO", "DEBITO"} and re.match(r"^[+-]?R\$\s*\d", lines[i + 2]):
            tipo = lines[i + 1]
            val = money_to_float(lines[i + 2])
            val = -abs(val) if tipo == "DEBITO" else abs(val)
            rows.append([current, ln, "", val])
            i += 3
            continue

        i += 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


def parse_santander(pdf_path, lines):
    re_sant_day_local = re.compile(
        r"^(Segunda|Terça|Terca|Quarta|Quinta|Sexta|Sábado|Sabado|Domingo),\s+(\d{1,2})\s+de\s+([A-Za-zç]+)\s+de\s+(\d{4})$",
        re.I
    )
    if any(re_sant_day_local.match(ln) for ln in lines[:200]):
        return parse_santander_layout2(lines)
    return parse_santander_layout1_from_pdf(pdf_path)

def detect_year(lines):
    sample = " ".join(lines[:400]).lower()
    m = re.search(r"\b(20\d{2})\b", sample)
    return int(m.group(1)) if m else datetime.now().year


def parse_itau(lines):
    year = detect_year(lines)
    rows = []
    current = None
    in_mov = False
    i = 0
    N = len(lines)

    header_noise = (
        "data", "descrição", "descricao", "entradas r$", "saídas r$", "saidas r$", "saldo r$",
        "(créditos)", "(debitos)", "(débitos)", "a = agendamento", "b = ações movimentadas",
        "c = crédito a compensar", "d = débito a compensar", "g = aplicação programada",
        "p = poupança automática", "para demais siglas, consulte as notas", "explicativas no final do extrato"
    )

    while i < N:
        ln = lines[i]
        low = ln.lower()

        if re.search(r"Conta\s+Corrente\s*\|\s*Movimenta", ln, re.I):
            in_mov = True
            current = None
            i += 1
            continue

        if not in_mov:
            i += 1
            continue

        # fim da seção correta
        if re.search(r"^Conta\s+Corrente\s*\|\s*Aplica", ln, re.I):
            break

        # datas
        if re.fullmatch(r"\d{2}/\d{2}", ln):
            current = datetime.strptime(f"{ln}/{year}", "%d/%m/%Y")
            i += 1
            continue
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", ln):
            current = datetime.strptime(ln, "%d/%m/%Y")
            i += 1
            continue

        if current is None:
            i += 1
            continue

        # ignorar linhas de resumo / saldo
        if low in header_noise or any(k in low for k in [
            "saldo anterior", "saldo final", "saldo da conta corrente", "saldo total disponível",
            "conta corrente | saldo", "conta corrente | cheque especial",
            "saldo aplic aut", "saldo aplic aut mais", "saldo aplic", "saldo em c/c",
            "entrada r$", "saída r$", "saida r$", "na conta corrente", "bruto", "líquido", "liquido", "total",
            "totalizador de aplicações automáticas", "os valores referentes ao totalizador"
        ]):
            i += 1
            continue

        # descrição de transação dentro da seção movimentação
        desc = ln

        # descarte de linhas que são só número/resumo
        if not re.search(r"[A-Za-zÀ-ÿ]", desc):
            i += 1
            continue

        # próxima linha deve ser valor da transação
        if i + 1 < N and re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}-?", lines[i + 1]):
            val = money_to_float(lines[i + 1])
            rows.append([current.strftime("%d/%m/%Y"), desc, "", val])
            i += 2
            continue

        i += 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))

def parse_efi(lines):
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    value_re = re.compile(r"^[+-]\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []
    i = 0
    N = len(lines)

    while i < N:
        if not date_re.match(lines[i]):
            i += 1
            continue

        dt = lines[i]
        if i + 1 < N and lines[i + 1].lower().startswith("saldo do dia"):
            i += 3
            continue

        j = i + 1
        desc_parts = []
        doc = ""
        val = None

        while j < N and not date_re.match(lines[j]):
            ln = lines[j]
            low = ln.lower()

            if any(x in low for x in [
                "efí s.a.", "ouvidoria:", "tecbiz - tecnologia", "banco 364", "agência ",
                "período", "tipo de saldo", "tipo de transação", "filtros aplicados",
                "relatório gerado em", "todos"
            ]) or ln in {"Valor", "Descrição", "Data", "Protocolo", "Valor (R$)", "Lançamentos", "Extrato ﬁnanceiro", "Extrato financeiro"}:
                j += 1
                continue

            if re.fullmatch(r"\d{6,}", ln):
                doc = ln
                j += 1
                continue

            if value_re.match(ln):
                val = money_to_float(ln)
                rows.append([dt, norm_space(" ".join(desc_parts)), doc, val])
                j += 1
                break

            desc_parts.append(ln)
            j += 1

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))




# ---------------- Unicred ----------------

try:
    import pdfplumber
except Exception:
    pdfplumber = None

def parse_unicred(pdf_path):
    # parser fiel ao script individual, com adaptação para saída padronizada
    if pdfplumber is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    import itertools
    from decimal import Decimal, ROUND_HALF_UP

    TOL = Decimal("0.01")

    def br_to_decimal(valor_str):
        return Decimal(valor_str.replace('.', '').replace(',', '.')).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    def extrair_valores_linha(linha):
        return re.findall(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})', linha)

    texto = ""
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto += t + "\n"

    linhas = texto.splitlines()
    registros = []
    saldo_anterior = None

    for linha in linhas:
        if "Saldo Anterior" in linha:
            valores = extrair_valores_linha(linha)
            if valores:
                saldo_anterior = br_to_decimal(valores[-1])
            continue

        if re.match(r"\d{2}/\d{2}/\d{4}", linha):
            data = linha[:10]
            valores = extrair_valores_linha(linha)
            if not valores:
                continue
            valor_mov = br_to_decimal(valores[0])
            saldo_info = br_to_decimal(valores[1]) if len(valores) > 1 else None
            historico = linha[11:].strip()
            if len(valores) > 1:
                historico = re.sub(r'\s+' + re.escape(valores[-1]) + r'\s*$', '', historico).strip()
            registros.append({
                "Data": data,
                "Historico": historico,
                "Valor": valor_mov,
                "Saldo_Informado": saldo_info
            })

    df = pd.DataFrame(registros).reset_index(drop=True)
    if df.empty or saldo_anterior is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    def resolver_bloco(bloco, saldo_ant, saldo_final):
        bloco = bloco.copy()
        bloco["Tipo"] = None
        livres = []
        for i, row in bloco.iterrows():
            hist = row["Historico"].upper()
            if "INTEGR PARC CAPITAL" in hist:
                bloco.at[i, "Tipo"] = "D"
            elif "RECEB" in hist:
                bloco.at[i, "Tipo"] = "C"
            else:
                livres.append(i)

        saldo_base = saldo_ant
        for i, row in bloco.iterrows():
            if row["Tipo"] == "C":
                saldo_base += row["Valor"]
            elif row["Tipo"] == "D":
                saldo_base -= row["Valor"]

        for combinacao in itertools.product(["C","D"], repeat=len(livres)):
            saldo_teste = saldo_base
            for idx, tipo in zip(livres, combinacao):
                valor = bloco.at[idx, "Valor"]
                saldo_teste = saldo_teste + valor if tipo == "C" else saldo_teste - valor
            if abs(saldo_teste - saldo_final) <= TOL:
                for idx, tipo in zip(livres, combinacao):
                    bloco.at[idx, "Tipo"] = tipo
                return bloco, True

        return bloco, False

    saldo_confirmado = saldo_anterior
    inicio_bloco = 0
    df["Tipo"] = None

    for i in range(len(df)):
        if pd.notna(df.loc[i, "Saldo_Informado"]):
            saldo_final = df.loc[i, "Saldo_Informado"]
            bloco = df.loc[inicio_bloco:i].copy()
            bloco_resolvido, ok = resolver_bloco(bloco, saldo_confirmado, saldo_final)
            if ok:
                df.loc[inicio_bloco:i, "Tipo"] = bloco_resolvido["Tipo"]
            saldo_confirmado = saldo_final
            inicio_bloco = i + 1

    out = pd.DataFrame({
        "Data": df["Data"],
        "Descrição": df["Historico"],
        "Documento": "",
        "Valor": [
            float(v) if t == "C" else -float(v) if t == "D" else None
            for v, t in zip(df["Valor"], df["Tipo"])
        ]
    })
    out = out[out["Valor"].notna()]
    return standardize(out)


# ---------------- Dispatcher ----------------

def parse_one_pdf(pdf_path):
    lines = extract_lines(pdf_path)
    if not lines:
        return "", pd.DataFrame()

    txt = " ".join(lines[:250]).lower()
    name = os.path.basename(pdf_path).lower()

    if ("consultas - extrato de conta corrente" in txt) or ("sisbb" in txt) or name.startswith("bb_layout"):
        layout, df = parse_bb_auto(lines)
        return layout, df

    if "banco abc" in txt or name == "abc.pdf":
        return "abc", parse_abc(lines)

    if "banrisul" in txt:
        return "banrisul", parse_banrisul(lines)

    if "sicredi" in txt or "cooperativa:" in txt:
        return "sicredi", parse_sicredi(lines)

    if "banco inter" in txt or "saldo por transação" in txt or name.startswith("inter"):
        return "inter", parse_inter(lines)

    if "santander" in txt or "contamax" in txt or name.startswith("santander"):
        return "santander", parse_santander(pdf_path, lines)

    if "itaú" in txt or "itau" in txt or "extrato mensal" in txt or name == "itau.pdf":
        return "itau", parse_itau(lines)

    if "unicred" in txt or "instituição financeira:  136" in txt.lower() or name == "unicred.pdf":
        return "unicred", parse_unicred(pdf_path)

    if re.search(r"\bef[ií]\b", txt) or "extrato financeiro" in txt or name == "efi_bank.pdf":
        return "efi", parse_efi(lines)

    layout, df = parse_bb_auto(lines)
    if not df.empty:
        return layout, df

    return "desconhecido", pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor"])


def parse_one_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".ofx":
        return "ofx", parse_ofx_file(file_path)
    return parse_one_pdf(file_path)


# ---------------- XLSX e fluxo ----------------

def export_xlsx(out_path, df_all, df_logs):
    engine = "xlsxwriter"
    try:
        __import__("xlsxwriter")
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(out_path, engine=engine) as writer:
        df_all.to_excel(writer, index=False, sheet_name="Consolidado")
        df_logs.to_excel(writer, index=False, sheet_name="Logs")

        if engine == "xlsxwriter":
            wb = writer.book
            ws = writer.sheets["Consolidado"]
            ws_logs = writer.sheets["Logs"]

            money_fmt = wb.add_format({"num_format": "R$ #,##0.00;[Red]-R$ #,##0.00"})
            date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

            ws.set_column("A:A", 35)
            ws.set_column("B:B", 12, date_fmt)
            ws.set_column("C:C", 110)
            ws.set_column("D:D", 22)
            ws.set_column("E:E", 16, money_fmt)
            ws.set_column("F:F", 6)
            ws.set_column("G:H", 16, money_fmt)

            ws_logs.set_column("A:A", 35)
            ws_logs.set_column("B:B", 20)
        else:
            ws = writer.sheets["Consolidado"]
            ws_logs = writer.sheets["Logs"]
            widths = {"A": 35, "B": 12, "C": 110, "D": 22, "E": 16, "F": 6, "G": 16, "H": 16}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
            ws_logs.column_dimensions["A"].width = 35
            ws_logs.column_dimensions["B"].width = 20
            for row in ws.iter_rows(min_row=2):
                row[1].number_format = 'DD/MM/YYYY'
                for idx in (4, 6, 7):
                    row[idx].number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'


def escolher_pasta():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta = filedialog.askdirectory(title="Selecione a pasta com PDFs")
        root.destroy()
        return pasta or None
    except Exception:
        return None


def processar_pasta(folder):
    arquivos = list_input_files(folder)
    if not arquivos:
        raise FileNotFoundError(f"Não encontrei arquivos PDF/OFX em: {folder}")

    dados = []
    logs = []
    total = len(arquivos)

    for idx, file_path in enumerate(arquivos, start=1):
        nome = os.path.basename(file_path)
        try:
            layout, df = parse_one_file(file_path)
            if df.empty:
                logs.append([nome, "erro"])
                print(f"[{idx}/{total}] ERRO - {nome} | sem transações extraídas")
                continue

            df.insert(0, "Arquivo", nome)
            dados.append(df)
            logs.append([nome, int(len(df))])
            print(f"[{idx}/{total}] OK   - {nome} | {layout} | {len(df)} transação(ões)")

        except Exception as e:
            logs.append([nome, "erro"])
            print(f"[{idx}/{total}] ERRO - {nome} | {e}")

    if dados:
        df_all = pd.concat(dados, ignore_index=True)
        df_all = df_all[["Arquivo", "Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"]]
    else:
        df_all = pd.DataFrame(columns=["Arquivo", "Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    df_logs = pd.DataFrame(logs, columns=["Arquivo", "n_transações_obtidas"])

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(folder, f"consolidado_lancamentos_{stamp}.xlsx")
    export_xlsx(out_path, df_all, df_logs)
    return df_all, df_logs, out_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pasta", help="Pasta com PDFs (se omitido, abre seletor)")
    args = parser.parse_args()

    folder = args.pasta or escolher_pasta()
    if not folder:
        print("Nenhuma pasta selecionada.")
        return

    try:
        df_all, df_logs, out_path = processar_pasta(folder)
        print("\nArquivo gerado:")
        print(out_path)
        print(f"Total de transações: {len(df_all)}")
        print(f"Arquivos com erro: {(df_logs['n_transações_obtidas'] == 'erro').sum()}")
    except Exception as e:
        print("\nERRO GERAL:")
        print(str(e))


if __name__ == "__main__":
    main()
