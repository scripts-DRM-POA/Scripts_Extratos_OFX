# -*- coding: utf-8 -*-
"""
Extrator multi-layout de extratos bancários (PDF + OFX) - v69

Observação: a extração permanece textual para os layouts já mapeados.
Somente o Itaú layout7 em PDF-imagem usa OCR (Tesseract) como fallback localizado.

Uso:
    python extratos_PDFmultilayout_ofx_e_comSenhas_v69_BB_layout5.py
ou
    python extratos_PDFmultilayout_ofx_e_comSenhas_v69_BB_layout5.py --pasta "C:\\caminho\\dos\\pdfs"

Gera um XLSX com três abas:
    - Consolidado
    - Logs
    - Log_Lacunas

Colunas do Consolidado:
    Arquivo | Data | Descrição | Documento | Valor | Tipo | Débito | Crédito | Classificação provável | Categoria

Colunas do Logs:
    Arquivo | n_transações_obtidas
"""

import sys
import os
import re
import unicodedata
import glob
import argparse
import getpass
import calendar
import shutil
from datetime import datetime

# Saída de console compatível com .bat/Prompt do Windows.
# Não força UTF-8 no stdout: isso evita mojibake quando o .bat/console está em ANSI/CP1252.
# Em vez disso, normaliza Unicode para NFC e codifica a mensagem usando o encoding atual do stream,
# substituindo apenas caracteres realmente incompatíveis com esse ambiente.
import builtins


def _console_safe_text(value, encoding=None) -> str:
    text = unicodedata.normalize("NFC", str(value))
    enc = encoding or "utf-8"
    try:
        return text.encode(enc, errors="replace").decode(enc, errors="replace")
    except Exception:
        return text.encode("utf-8", errors="replace").decode("utf-8", errors="replace")


def print(*args, sep=" ", end="\n", file=None, flush=False):
    stream = file if file is not None else sys.stdout
    enc = getattr(stream, "encoding", None) or "utf-8"
    safe_args = [_console_safe_text(arg, enc) for arg in args]
    safe_sep = _console_safe_text(sep, enc)
    safe_end = _console_safe_text(end, enc)
    builtins.print(*safe_args, sep=safe_sep, end=safe_end, file=stream, flush=flush)

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

# OCR é opcional e usado SOMENTE no Itaú layout7 em PDF-imagem.
# Os demais layouts continuam sem depender de OCR/Tesseract.
try:
    import pytesseract
    from PIL import Image, ImageOps
except Exception:
    pytesseract = None
    Image = None
    ImageOps = None


PDF_PASSWORD_CACHE = []
PDF_PASSWORD_BY_FILE = {}


def _prompt_pdf_password(pdf_path: str):
    prompt_title = "Senha do PDF"
    prompt_text = f"Informe a senha para abrir o PDF:\n{os.path.basename(pdf_path)}"
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        senha = simpledialog.askstring(prompt_title, prompt_text, show="*", parent=root)
        root.destroy()
        return senha
    except Exception:
        try:
            return getpass.getpass(f"Senha do PDF '{os.path.basename(pdf_path)}' (ENTER cancela): ")
        except Exception:
            return None


def open_pdf_with_password(pdf_path: str):
    doc = fitz.open(pdf_path)
    if not doc.needs_pass:
        PDF_PASSWORD_BY_FILE[os.path.abspath(pdf_path)] = None
        return doc, None

    candidates = []
    file_key = os.path.abspath(pdf_path)
    if file_key in PDF_PASSWORD_BY_FILE and PDF_PASSWORD_BY_FILE[file_key] is not None:
        candidates.append(PDF_PASSWORD_BY_FILE[file_key])
    for pw in PDF_PASSWORD_CACHE:
        if pw not in candidates:
            candidates.append(pw)

    for pw in candidates:
        try:
            if doc.authenticate(pw):
                PDF_PASSWORD_BY_FILE[file_key] = pw
                if pw and pw not in PDF_PASSWORD_CACHE:
                    PDF_PASSWORD_CACHE.append(pw)
                return doc, pw
        except Exception:
            pass

    while True:
        senha = _prompt_pdf_password(pdf_path)
        if senha is None or senha == "":
            doc.close()
            raise RuntimeError("PDF protegido por senha; operação cancelada pelo usuário")
        try:
            if doc.authenticate(senha):
                PDF_PASSWORD_BY_FILE[file_key] = senha
                if senha not in PDF_PASSWORD_CACHE:
                    PDF_PASSWORD_CACHE.append(senha)
                return doc, senha
        except Exception:
            pass


MESES_PT = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}
MESES_BAN = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}


def norm_space(s: str) -> str:
    s = str(s or "")
    # Normaliza Unicode para a forma composta (NFC), convertendo sequências como
    # "c" + cedilha combinante (\u0327) em "ç" quando possível.
    s = unicodedata.normalize("NFC", s)
    s = (
        s.replace("\xa0", " ")
         .replace("\uf166", " ")
         .replace("\ue90a", " ")
         .replace("\uf18f", " ")
    )
    return re.sub(r"\s{2,}", " ", s.strip())


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
    doc, password = open_pdf_with_password(pdf_path)
    try:
        lines = []
        for page in doc:
            txt = page.get_text("text") or ""
            for ln in txt.splitlines():
                ln = norm_space(ln)
                if ln:
                    lines.append(ln)
        return lines, password
    finally:
        doc.close()


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
    """Normaliza datas OFX em formatos bancários usuais.

    Mantém o suporte original a AAAAMMDDhhmmss e acrescenta suporte a
    datas já formatadas, como DD.MM.AAAA, DD/MM/AAAA e DD-MM-AAAA.
    Alguns OFX do Banco do Brasil exportam <DTPOSTED>04.01.2021</DTPOSTED>,
    sem a forma AAAAMMDD.
    """
    if not dt:
        return ""
    raw = norm_space(str(dt))

    # Formato clássico OFX: AAAAMMDD, eventualmente seguido de hora/fuso.
    m = re.search(r"(\d{8})", raw)
    if m:
        s = m.group(1)
        candidates = []
        if 1900 <= int(s[:4]) <= 2099:
            candidates.append((s, "%Y%m%d"))
        if 1900 <= int(s[4:]) <= 2099:
            candidates.append((s, "%d%m%Y"))
        for value, fmt in candidates:
            try:
                return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
            except Exception:
                pass

    # Formatos encontrados em alguns OFX exportados por internet banking.
    m = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", raw)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).strftime("%d/%m/%Y")
        except Exception:
            return ""

    m = re.search(r"\b(\d{4})[./-](\d{1,2})[./-](\d{1,2})\b", raw)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%d/%m/%Y")
        except Exception:
            return ""

    return ""


def parse_ofx_amount(valor_txt: str, trntype: str = ""):
    """Converte <TRNAMT> de OFX para float.

    Além do padrão OFX comum (-123.45), aceita exportações que trazem
    valores em formato brasileiro com sufixo C/D, como:
        "10.200,00 C"
        "2.390,72 D"
    O sufixo C/D prevalece sobre TRNTYPE quando presente.
    """
    raw = norm_space(str(valor_txt or "")).replace("−", "-")
    if not raw:
        return None, False

    explicit_cd = None
    m_cd = re.search(r"(?:^|\s)([CD])\s*$", raw, flags=re.I)
    if m_cd:
        explicit_cd = m_cd.group(1).upper()
        raw = raw[:m_cd.start(1)].strip()

    raw = raw.replace("R$", "").replace("+", " ").strip()
    neg_by_sign = raw.startswith("-") or raw.endswith("-")

    m_num = re.search(r"-?\d[\d.,]*-?", raw)
    if not m_num:
        return None, explicit_cd is not None

    num = m_num.group(0).strip()
    if num.startswith("-") or num.endswith("-"):
        neg_by_sign = True
    num = num.strip("-")

    if "," in num and "." in num:
        if num.rfind(",") > num.rfind("."):
            normalized = num.replace(".", "").replace(",", ".")
        else:
            normalized = num.replace(",", "")
    elif "," in num:
        normalized = num.replace(".", "").replace(",", ".")
    elif "." in num:
        parts = num.split(".")
        if len(parts) > 2 and len(parts[-1]) == 3:
            normalized = "".join(parts)
        else:
            normalized = num
    else:
        normalized = num

    try:
        value = float(normalized)
    except Exception:
        return None, explicit_cd is not None

    if explicit_cd == "D":
        value = -abs(value)
    elif explicit_cd == "C":
        value = abs(value)
    elif neg_by_sign:
        value = -abs(value)

    return value, explicit_cd is not None


def _decode_ofx_bytes(raw: bytes) -> str:
    """Decodifica OFX evitando mojibake em textos UTF-8/Windows-1252/Latin-1."""
    head = raw[:4096].decode("ascii", errors="ignore")
    candidates = []

    # Respeita pistas explícitas quando existirem, mas mantém fallback robusto.
    for m in re.finditer(r"(?:CHARSET|ENCODING)\s*:\s*([^\r\n]+)|encoding=[\"']([^\"']+)[\"']", head, flags=re.I):
        enc = (m.group(1) or m.group(2) or "").strip().upper()
        if not enc:
            continue
        if "UTF" in enc:
            candidates.extend(["utf-8-sig", "utf-8"])
        elif enc in {"1252", "WINDOWS-1252", "CP1252"}:
            candidates.append("cp1252")
        elif "ISO" in enc or "LATIN" in enc or enc in {"USASCII", "US-ASCII", "NONE"}:
            # Muitos OFX declaram USASCII, mas trazem acentuação em cp1252/latin1.
            candidates.extend(["utf-8-sig", "utf-8", "cp1252", "latin1"])

    candidates.extend(["utf-8-sig", "utf-8", "cp1252", "latin1"])

    seen = set()
    ordered = []
    for enc in candidates:
        if enc not in seen:
            seen.add(enc)
            ordered.append(enc)

    for enc in ordered:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin1", errors="ignore")


def fix_mojibake_ptbr(s: str) -> str:
    """Corrige casos como 'AplicaÃ§Ã£o' -> 'Aplicação', sem alterar texto já correto."""
    s = norm_space(str(s or ""))
    if not s:
        return ""
    if not re.search(r"[ÃÂâ]\S*", s):
        return s
    try:
        fixed = s.encode("latin1", errors="strict").decode("utf-8", errors="strict")
        # Só aceita se a correção reduzir marcas típicas de mojibake.
        bad_before = len(re.findall(r"[ÃÂâ]", s))
        bad_after = len(re.findall(r"[ÃÂâ]", fixed))
        if bad_after < bad_before:
            return norm_space(fixed)
    except Exception:
        pass
    return s


def is_ofx_balance_transaction(descricao: str) -> bool:
    """Identifica lançamentos OFX meramente informativos de saldo.

    Alguns bancos exportam saldos dentro de blocos <STMTTRN>, como se fossem
    transações: SALDO TOTAL DISPONÍVEL DIA, SALDO MOVIMENTAÇÃO CONTA,
    SALDO APLIC. AUT., SALDO ANTERIOR, SALDO FINAL etc. Esses registros não
    representam movimentação bancária e devem ser ignorados apenas no parser OFX.
    """
    desc = norm_space(str(descricao or "")).upper()
    if not desc:
        return False

    # Regra conservadora: remove somente memos/names que contenham a palavra SALDO.
    # Mantém aplicações/resgates/rendimentos reais, como APL APLIC AUT MAIS,
    # RES APLIC AUT MAIS e REND PAGO APLIC AUT MAIS.
    return bool(re.search(r"\bSALDO\b", desc))


def ofx_contains_cpf_cnpj(s: str) -> bool:
    """Detecta CPF/CNPJ com ou sem pontuação em trecho textual OFX."""
    s = str(s or "")
    patterns = [
        r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b",
        r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b",
        r"\b\d{11}\b",
        r"\b\d{14}\b",
    ]
    return any(re.search(p, s) for p in patterns)


def ofx_has_letters(s: str) -> bool:
    return bool(re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", str(s or "")))


def ofx_text_key(s: str) -> str:
    s = fix_mojibake_ptbr(s).lower()
    s = re.sub(r"[^a-z0-9à-öø-ÿ]+", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def ofx_is_relevant_detail(s: str) -> bool:
    """Evita jogar identificadores técnicos puros na descrição.

    Um complemento é considerado útil quando traz CPF/CNPJ ou expressão textual
    com cara de contraparte/beneficiário. Acrônimos curtos e sufixos técnicos
    como "Iof.ADic. -1" ficam fora.
    """
    s = fix_mojibake_ptbr(s)
    if not s or len(s) < 4:
        return False

    if ofx_contains_cpf_cnpj(s):
        return True

    if not ofx_has_letters(s):
        return False

    # Exclui hashes/ids técnicos sem espaçamento e sem cara de texto natural.
    if re.fullmatch(r"[A-Z0-9_.:/-]{4,}", s.upper()):
        return False

    alpha_tokens = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]{3,}", s)
    entity_terms = (
        "LTDA", "EIRELI", "COOPERATIVA", "COOP", "S/A", "SA", "ME", "EPP",
        "COMERCIO", "COMÉRCIO", "SERVICOS", "SERVIÇOS", "INDUSTRIA", "INDÚSTRIA",
        "HOSPITAL", "CLINICA", "CLÍNICA", "ESCOLA", "UNIMED"
    )
    up = s.upper()
    if len(alpha_tokens) >= 2 or any(term in up for term in entity_terms):
        return True

    return False


def ofx_is_good_nucleus(s: str) -> bool:
    """Reconhece o campo NAME como núcleo quando ele sintetiza a operação."""
    s = fix_mojibake_ptbr(s)
    if not s or not ofx_has_letters(s):
        return False
    up = s.upper()
    markers = (
        "PIX", "TED", "DOC", "TEF", "TRANSF", "TRANSFER", "TARIFA", "PAG", "PAGAMENTO",
        "BOLETO", "COBR", "COBRAN", "DÉBITO", "DEBITO", "CRÉDITO", "CREDITO", "CARTAO",
        "CARTÃO", "APLIC", "RESGATE", "REND", "ESTORNO", "COMPRA", "SAQUE", "DEPÓSITO",
        "DEPOSITO", "RECEBIDO", "ENVIADO"
    )
    return any(m in up for m in markers)


def ofx_is_code_token(tok: str) -> bool:
    tok = norm_space(tok)
    if not tok:
        return False
    up = tok.upper()
    digits = only_digits(up)
    # CPF/CNPJ deve ser preservado como detalhe quando houver nome/contraparte depois.
    if len(digits) in {11, 14}:
        return False
    if re.fullmatch(r"\d{1,10}", up):
        return True
    if re.fullmatch(r"[A-Z]{1,6}\d[A-Z0-9_.-]{0,12}", up):
        return True
    if re.fullmatch(r"[A-Z]+_[A-Z0-9_]+", up):
        return True
    if re.fullmatch(r"[A-Z0-9]{4,18}", up) and any(ch.isdigit() for ch in up):
        return True
    return False


def ofx_split_document_and_detail(s: str):
    """Separa 'códigos iniciais + contraparte' em (documento, detalhe).

    Exemplos:
      COB000001 241000013 UNIMED... -> doc=COB000001 241000013; detalhe=UNIMED...
      CX01551 13969248000107 LIFE STAR LTDA -> doc=CX01551; detalhe=13969248000107 LIFE STAR LTDA
      PIX_DEB 02597157008 JONATHAN... -> doc=PIX_DEB; detalhe=02597157008 JONATHAN...
    """
    s = fix_mojibake_ptbr(s)
    if not s:
        return "", ""
    parts = s.split()
    if len(parts) < 2:
        return s, ""

    doc_parts = []
    split_at = 0
    for idx, tok in enumerate(parts):
        remaining = parts[idx + 1:]
        digits = only_digits(tok)
        has_text_after = any(ofx_has_letters(x) for x in remaining)

        if len(digits) in {11, 14} and has_text_after:
            split_at = idx
            break

        if ofx_is_code_token(tok):
            doc_parts.append(tok)
            split_at = idx + 1
            continue

        # Primeira palavra que parece nome/razão social: a partir daqui é detalhe.
        split_at = idx
        break

    detail = norm_space(" ".join(parts[split_at:]))
    doc = norm_space(" ".join(doc_parts))

    # Não usa REFNUM/FITID puramente textual como complemento: para evitar
    # casos técnicos/acrônimos soltos, exige ao menos um prefixo documental claro.
    if not doc:
        return s, ""

    if not ofx_is_relevant_detail(detail):
        return s, ""
    return doc, detail


def ofx_add_detail(base: str, detail: str) -> str:
    base = fix_mojibake_ptbr(base)
    detail = fix_mojibake_ptbr(detail)
    if not ofx_is_relevant_detail(detail):
        return base
    kb = ofx_text_key(base)
    kd = ofx_text_key(detail)
    if not kd or kd == kb or kd in kb:
        return base
    if kb and kb in kd and len(kd) <= len(kb) + 8:
        return base
    return norm_space(base + " | " + detail)


def build_ofx_description_and_document(name: str, memo: str, trntype: str, checknum: str, refnum: str, fitid: str):
    """Monta descrição OFX com núcleo + detalhes relevantes, sem poluir com ids técnicos."""
    name = fix_mojibake_ptbr(name)
    memo = fix_mojibake_ptbr(memo)
    trntype = fix_mojibake_ptbr(trntype)
    checknum = fix_mojibake_ptbr(checknum)
    refnum = fix_mojibake_ptbr(refnum)
    fitid = fix_mojibake_ptbr(fitid)

    # Mantém a lógica tradicional quando não há NAME útil.
    if name and memo and ofx_is_good_nucleus(name):
        descricao = name
        descricao = ofx_add_detail(descricao, memo)
    else:
        descricao = memo or name or trntype or "Lançamento OFX"

    documento_original = checknum or refnum or fitid
    documento = documento_original

    # REFNUM costuma ser melhor para complemento textual do que FITID quando ambos existem.
    for candidate in [refnum, fitid]:
        if not candidate:
            continue
        if checknum and ofx_text_key(candidate) == ofx_text_key(checknum):
            continue
        doc_part, detail = ofx_split_document_and_detail(candidate)
        if detail:
            descricao = ofx_add_detail(descricao, detail)
            # Só encurta Documento quando o próprio campo usado como Documento contém texto descritivo.
            if documento_original == candidate and doc_part:
                documento = doc_part
            break

    return descricao, documento


def parse_ofx_file(ofx_path):
    raw_bytes = open(ofx_path, "rb").read()
    texto = _decode_ofx_bytes(raw_bytes)

    blocos = re.findall(r"<STMTTRN>(.*?)</STMTTRN>", texto, flags=re.S | re.I)
    rows = []

    def campo(raw: str, tag: str) -> str:
        m = re.search(rf"<{tag}>(.*?)(?:$|<)", raw, flags=re.I | re.S)
        return fix_mojibake_ptbr(m.group(1)) if m else ""

    debit_types = {
        "DEBIT", "PAYMENT", "DIRECTDEBIT", "REPEATPMT", "ATM",
        "POS", "CHECK", "FEE", "SRVCHG"
    }
    credit_types = {
        "CREDIT", "DEP", "DIRECTDEP", "DIV", "INT"
    }

    for raw in blocos:
        data = normalize_ofx_date(campo(raw, "DTPOSTED"))
        valor_txt = campo(raw, "TRNAMT")
        trntype = campo(raw, "TRNTYPE").upper()

        if not data or not valor_txt:
            continue

        valor, valor_tem_sufixo_cd = parse_ofx_amount(valor_txt, trntype)
        if valor is None:
            continue

        # Alguns OFX vêm com TRNAMT sem sinal, mas com TRNTYPE correto.
        # Nesses casos, o tipo da transação deve prevalecer.
        # Quando o próprio TRNAMT traz sufixo C/D, ele é mais específico e prevalece.
        if not valor_tem_sufixo_cd:
            if trntype in debit_types:
                valor = -abs(valor)
            elif trntype in credit_types:
                valor = abs(valor)

        name = campo(raw, "NAME")
        memo = campo(raw, "MEMO")
        checknum = campo(raw, "CHECKNUM")
        refnum = campo(raw, "REFNUM")
        fitid = campo(raw, "FITID")

        descricao, documento = build_ofx_description_and_document(
            name=name,
            memo=memo,
            trntype=trntype,
            checknum=checknum,
            refnum=refnum,
            fitid=fitid,
        )

        if is_ofx_balance_transaction(descricao):
            continue

        rows.append([data, descricao, documento, valor])

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )

# ---------------- C6 Bank ----------------


def parse_c6(lines, source_name=""):
    month_heading = re.compile(
        r"^(Janeiro|Fevereiro|Março|Marco|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro)\s+(\d{4})\b(?:\s*\(\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})\s*\))?",
        re.I
    )
    date_mmdd = re.compile(r"^\d{2}/\d{2}$")
    money_re = re.compile(r"^-?R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}$")

    month_to_num = {
        "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
        "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
    }

    def last_day_of_month(year, month):
        return calendar.monthrange(year, month)[1]

    def parse_block_headers(all_lines):
        headers = []
        for idx, line in enumerate(all_lines):
            m = month_heading.match(line)
            if not m:
                continue
            mon_name = m.group(1)
            year = int(m.group(2))
            month = month_to_num[mon_name.lower()]
            start_s = m.group(3)
            end_s = m.group(4)
            start_dt = datetime.strptime(start_s, "%d/%m/%Y") if start_s else None
            end_dt = datetime.strptime(end_s, "%d/%m/%Y") if end_s else None
            full_month = True
            if start_dt and end_dt:
                full_month = (start_dt.day == 1 and end_dt.day == last_day_of_month(year, month))
            headers.append({
                "idx": idx,
                "year": year,
                "month": month,
                "start_dt": start_dt,
                "end_dt": end_dt,
                "full_month": full_month,
            })
        return headers

    def compute_allowed_years(headers):
        if not headers:
            return None

        year_months = {}
        for h in headers:
            year_months.setdefault(h["year"], set()).add(h["month"])
        years_in_order = []
        for h in headers:
            if not years_in_order or years_in_order[-1] != h["year"]:
                years_in_order.append(h["year"])

        allowed_years = set(year_months.keys())
        if len(years_in_order) == 1:
            return allowed_years

        last_year = years_in_order[-1]
        prev_year = years_in_order[-2]
        last_headers = [h for h in headers if h["year"] == last_year]
        prev_headers = [h for h in headers if h["year"] == prev_year]
        last_month_count = len(year_months.get(last_year, set()))
        prev_month_count = len(year_months.get(prev_year, set()))
        last_all_partial = bool(last_headers) and all(not h["full_month"] for h in last_headers)

        # Regra 1: sobra clara do ano seguinte no final do PDF anual.
        if last_year > prev_year and prev_month_count >= 10 and last_month_count <= 2:
            allowed_years.discard(last_year)
            return allowed_years

        # Regra 2: bloco final nitidamente parcial, mesmo sem depender do nome do arquivo.
        if last_year > prev_year and prev_month_count >= 10 and last_all_partial:
            allowed_years.discard(last_year)
            return allowed_years

        return allowed_years

    headers = parse_block_headers(lines)
    allowed_years = compute_allowed_years(headers)

    rows = []
    current_year = None
    i = 0
    N = len(lines)

    while i < N:
        ln = lines[i]
        m = month_heading.match(ln)
        if m:
            current_year = int(m.group(2))
            i += 1
            continue

        low = ln.lower()
        if (
            low.startswith("sem lançamentos no mês") or low.startswith("sem lancamentos no mes") or
            low.startswith("saldo do dia ") or low.startswith("saldo do dia •") or
            low.startswith("entradas:") or low.startswith("saídas:") or low.startswith("saidas:")
        ):
            i += 1
            continue

        if low in {"data", "lançamento", "lancamento", "contábil", "contabil", "tipo", "descrição", "descricao", "valor"} or ln == "•":
            i += 1
            continue

        if current_year and allowed_years is not None and current_year not in allowed_years:
            i += 1
            continue

        if current_year and date_mmdd.match(ln) and i + 3 < N and date_mmdd.match(lines[i + 1]):
            data_lanc = ln
            tipo = lines[i + 2]
            desc = lines[i + 3]
            j = i + 4

            while j < N and not money_re.match(lines[j]):
                nxt = lines[j]
                if (
                    date_mmdd.match(nxt) or nxt.startswith("Saldo do dia ") or
                    nxt.startswith("Saldo do dia •") or month_heading.match(nxt)
                ):
                    break
                low_nxt = nxt.lower()
                if low_nxt in {"data", "lançamento", "lancamento", "contábil", "contabil", "tipo", "descrição", "descricao", "valor"} or nxt == "•":
                    j += 1
                    continue
                desc = norm_space(desc + " " + nxt)
                j += 1

            if j < N and money_re.match(lines[j]):
                try:
                    data = datetime.strptime(f"{data_lanc}/{current_year}", "%d/%m/%Y").strftime("%d/%m/%Y")
                    valor = money_to_float(lines[j])
                    descricao = norm_space(f"{tipo} {desc}")
                    rows.append([data, descricao, "", valor])
                    i = j + 1
                    continue
                except Exception:
                    pass

        i += 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


# ---------------- Caixa ----------------

def parse_caixa_layout1(lines):
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    doc_re = re.compile(r"^\d{4,}$")
    val_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}\s+[CD]$")
    saldo_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}\s+[CD]$")
    rows = []
    i = 0
    N = len(lines)
    while i < N:
        ln = lines[i]
        low = ln.lower()
        if not date_re.match(ln):
            i += 1
            continue
        # data + doc + histórico + valor [saldo opcional]
        if i + 3 >= N or not doc_re.match(lines[i+1]):
            i += 1
            continue
        data = ln
        doc = lines[i+1]
        j = i + 2
        hist_parts = []
        while j < N and not val_re.match(lines[j]):
            cur = lines[j]
            cur_low = cur.lower()
            if date_re.match(cur) or cur_low.startswith('saldo anterior') or cur_low.startswith('limite '):
                break
            if cur_low in {'data mov.','nr. doc.','histórico','historico','valor','saldo'}:
                j += 1
                continue
            if cur_low.startswith('firefox ') or 'about:blank' in cur_low:
                j += 1
                continue
            hist_parts.append(cur)
            j += 1
        if not hist_parts or j >= N or not val_re.match(lines[j]):
            i += 1
            continue
        hist = norm_space(' '.join(hist_parts))
        valtok = lines[j]
        m = re.match(r'^(\d{1,3}(?:\.\d{3})*,\d{2})\s+([CD])$', valtok)
        valor = money_to_float(m.group(1))
        valor = abs(valor) if m.group(2) == 'C' else -abs(valor)
        rows.append([data, hist, doc, valor])
        j += 1
        if j < N and saldo_re.match(lines[j]):
            j += 1
        i = j
    return standardize(pd.DataFrame(rows, columns=['Data','Descrição','Documento','Valor']), doc_cleaner=clean_document_token_flexible)


def parse_caixa_layout2(pdf_path, pdf_password=None):
    if pdfplumber is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    texto = ""
    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    linhas = [re.sub(r"\s{2,}", " ", l.strip()) for l in texto.splitlines() if l.strip()]

    re_tx = re.compile(
        r'^(?P<data>\d{2}/\d{2}/\d{4})\s+\d{2}/\d{2}\s+\d{2}:\d{2}\s+'
        r'(?P<doc>\d{6,})\s+(?P<hist>.*?)\s+'
        r'(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})(?P<dc>[CD])\s+'
        r'(?P<saldo>\d{1,3}(?:\.\d{3})*,\d{2})(?P<sdc>[CD])$'
    )
    re_periodo = re.compile(r'^\d{2}/\d{2}/\d{4}\s+-\s+\d{2}/\d{2}/\d{4}\b')

    skip_exact = {
        'extrato histórico da conta', 'extrato historico da conta',
        'titular', 'periodo', 'conta', 'nome do produto', 'data e hora',
        'data mov. data e hora nr.doc. histórico valor saldo',
        'data mov. data e hora nr.doc. historico valor saldo',
        'data mov. nr.doc. histórico valor saldo',
        'data mov. nr.doc. historico valor saldo'
    }
    skip_contains = [
        'consulta realizada em', 'página ', 'pagina ',
        'conta corrente pessoa juridica caixa', 'cpf/cnpj do titular',
        'nome da unidade', 'unidade', 'saldo anterior'
    ]
    rotativo_markers = [
        'informações do limite rotativo', 'informacoes do limite rotativo',
        'juros prov', 'mora prov', 'multa prov', 'iof aliq básica', 'iof aliq basica',
        'iof aliq adicional', 'custo efetivo total a.m', 'custo efetivo total a.a',
        'sublimite aval', 'limite disponibilizado', 'limite utilizado',
        'limite disponível', 'limite disponivel'
    ]

    rows = []
    last_idx = None
    in_rotativo = False

    for line in linhas:
        low = line.lower()

        m = re_tx.match(line)
        if m:
            in_rotativo = False
            data = m.group('data')
            doc = m.group('doc')
            hist = norm_space(m.group('hist'))
            val = money_to_float(m.group('val'))
            val = abs(val) if m.group('dc') == 'C' else -abs(val)
            if hist and val != 0:
                rows.append([data, hist, doc, val])
                last_idx = len(rows) - 1
            else:
                last_idx = None
            continue

        if 'informações do limite rotativo' in low or 'informacoes do limite rotativo' in low:
            in_rotativo = True
            last_idx = None
            continue

        if in_rotativo:
            if re_periodo.match(line):
                # novo bloco/período depois da seção de limite rotativo
                in_rotativo = False
            else:
                # ignora apenas a seção de limite; segue lendo o documento depois
                continue

        if low in skip_exact:
            last_idx = None
            continue
        if any(x in low for x in skip_contains):
            continue
        if any(marker in low for marker in rotativo_markers):
            last_idx = None
            continue
        if re_periodo.match(line):
            last_idx = None
            continue

        if last_idx is not None:
            if re.match(r'^\d{2}/\d{2}/\d{4}', line):
                last_idx = None
                continue
            if re.fullmatch(r'\d{1,3}(?:\.\d{3})*/\d{2}-\d', line):
                continue
            rows[last_idx][1] = norm_space(rows[last_idx][1] + ' ' + line)

    df = pd.DataFrame(rows, columns=['Data','Descrição','Documento','Valor'])
    return standardize(df, doc_cleaner=clean_document_token_flexible)

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


def is_bb_layout5_signature(txt: str, name: str = "") -> bool:
    """Reconhece o BB_layout5 (Extrato de Poupança).

    Estrutura observada:
      - cabeçalho "Extrato de Poupança" e "Dados da Conta";
      - seção "Lançamentos";
      - cada lançamento tem data completa e valor em R$, com sinal negativo
        anteposto ao símbolo ("-R$") nos débitos;
      - a descrição aparece logo abaixo da data;
      - rodapé "Saldos por data-base".
    """
    sample = norm_space(str(txt or "")).lower()
    nm = str(name or "").lower()
    return (
        (
            "extrato de poupança" in sample
            and "dados da conta" in sample
            and "lançamentos" in sample
            and "saldos por data-base" in sample
        )
        or nm.startswith("bb_layout5")
    )


def parse_bb_layout5(pdf_path, pdf_password=None):
    """Parser do BB_layout5 (Extrato de Poupança).

    A leitura usa a geometria do PDF, e não apenas a ordem linear do texto.
    Isso é necessário porque em algumas páginas (especialmente a última) os
    valores podem aparecer no fluxo textual depois de todas as descrições,
    embora visualmente estejam alinhados à respectiva data.

    Regras:
      - somente registros entre "Lançamentos" e "Saldos por data-base";
      - data no formato dd/mm/aaaa;
      - valor alinhado à mesma linha da data;
      - "-R$" = débito; "R$" = crédito;
      - descrição = texto entre a data atual e a próxima data, na coluna esquerda;
      - Documento permanece vazio neste layout.
    """
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    amount_re = re.compile(
        r"(?P<neg>-)?R\$\s*(?P<num>\d{1,3}(?:\.\d{3})*,\d{2})",
        re.I,
    )
    rows = []

    doc, _ = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            words = page.get_text("words") or []
            if not words:
                continue

            # Limites verticais da seção transacional.
            lanc_y = [
                float(w[1]) for w in words
                if norm_space(str(w[4])).lower() in {"lançamentos", "lancamentos"}
            ]
            if not lanc_y:
                continue
            start_y = min(lanc_y) + 2.0

            stop_candidates = []
            for w in words:
                if norm_space(str(w[4])).lower() != "saldos":
                    continue
                y = float(w[1])
                same_line = [
                    norm_space(str(x[4])).lower()
                    for x in words
                    if abs(float(x[1]) - y) <= 2.0
                ]
                if "por" in same_line and (
                    "data-base" in same_line
                    or ("data" in same_line and "base" in same_line)
                ):
                    stop_candidates.append(y)

            stop_y = min(
                [y for y in stop_candidates if y > start_y],
                default=float(page.rect.height),
            )

            # Datas relevantes da seção. A ordenação por coordenada preserva
            # a ordem visual mesmo quando a extração linear do PDF é irregular.
            date_words = []
            for w in words:
                x0, y0, x1, y1, token = w[:5]
                token = norm_space(str(token))
                if start_y < float(y0) < stop_y and date_re.fullmatch(token):
                    date_words.append((float(y0), token))
            date_words.sort(key=lambda t: t[0])

            for idx, (date_y, data) in enumerate(date_words):
                next_y = (
                    date_words[idx + 1][0]
                    if idx + 1 < len(date_words)
                    else stop_y
                )

                # O valor fica visualmente na mesma linha da data, na coluna direita.
                amount_words = [
                    w for w in words
                    if abs(float(w[1]) - date_y) <= 2.0
                    and float(w[0]) > float(page.rect.width) * 0.55
                ]
                amount_words.sort(key=lambda w: float(w[0]))
                amount_text = norm_space(" ".join(str(w[4]) for w in amount_words))
                m = amount_re.search(amount_text)
                if not m:
                    continue

                valor_txt = ("-" if m.group("neg") else "") + m.group("num")
                try:
                    valor = money_to_float(valor_txt)
                except Exception:
                    continue

                # A descrição aparece abaixo da data e antes da próxima data.
                # Limita-se à coluna esquerda para não incorporar valores/rodapé.
                desc_words = [
                    w for w in words
                    if date_y + 1.0 < float(w[1]) < next_y - 1.0
                    and float(w[0]) < float(page.rect.width) * 0.60
                ]
                desc_words.sort(key=lambda w: (float(w[1]), float(w[0])))

                desc_lines = []
                current_y = None
                current_words = []
                for w in desc_words:
                    y = float(w[1])
                    if current_y is None or abs(y - current_y) <= 2.0:
                        current_words.append(w)
                        if current_y is None:
                            current_y = y
                    else:
                        current_words.sort(key=lambda z: float(z[0]))
                        line = norm_space(" ".join(str(z[4]) for z in current_words))
                        if line:
                            desc_lines.append(line)
                        current_words = [w]
                        current_y = y

                if current_words:
                    current_words.sort(key=lambda z: float(z[0]))
                    line = norm_space(" ".join(str(z[4]) for z in current_words))
                    if line:
                        desc_lines.append(line)

                descricao = norm_space(" ".join(desc_lines))
                if not descricao or is_balance_or_summary_line(descricao):
                    continue

                rows.append([data, descricao, "", valor])
    finally:
        doc.close()

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: "",
    )


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

def parse_banrisul_layout1(lines):
    """Parser Banrisul do baseline v51.

    Mantém a compatibilidade com o layout já estabilizado, cuja competência
    vem em linha "PERIODO:" / "PERÍODO:".
    """
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


def parse_banrisul_layout2(lines):
    """Parser Banrisul layout II, incorporado sobre o baseline v51.

    Layout observado:
      B A N R I S U L - DEZEMBRO /2021 - PAG. 1
      DIA OP DOC SIS HISTORICO VALOR SALDO ORIG R*
      01 1058 000010 BEH OP.CREDITO C/PENHOR 1.809,38
         1325 211813 BDX DEBITO TRANSFERENCIA 171,90-
         CPF/CNPJ DESTINO 00090205960049
         AGENCIA DESTINO 0165
         CONTA DESTINO 0805952609

    A rotina é isolada: não altera OFX, não altera standardize() e não
    interfere nos demais bancos. Linhas de saldo, transporte, rodapé e taxa
    informativa de último dia são descartadas; linhas complementares de
    contraparte/destino são agregadas à descrição do lançamento imediatamente
    anterior, sem gerar novas transações.
    """
    month_re = re.compile(
        r"(?:B\s*A\s*N\s*R\s*I\s*S\s*U\s*L\s*-\s*)?"
        r"(JANEIRO|FEVEREIRO|MARCO|MARÇO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)"
        r"\s*/\s*(\d{4})",
        re.I,
    )

    # Alguns PDFs trazem mais de um lançamento na mesma linha, ou uma linha
    # complementar seguida de novo lançamento. Este marcador permite fatiar
    # essas ocorrências antes da interpretação.
    tx_start_re = re.compile(r"(?<!\S)(?:(?:\d{2})\s+)?\d{4}\s+\d{6}\s+[A-Z]{3}\s+")

    tx_day_re = re.compile(
        r"^(?P<day>\d{2})\s+(?P<op>\d{4})\s+(?P<doc>\d{6})\s+(?P<sis>[A-Z]{3})\s+"
        r"(?P<hist>.+?)\s+(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)(?P<trail>.*)$"
    )
    tx_noday_re = re.compile(
        r"^(?P<op>\d{4})\s+(?P<doc>\d{6})\s+(?P<sis>[A-Z]{3})\s+"
        r"(?P<hist>.+?)\s+(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)(?P<trail>.*)$"
    )

    skip_contains = [
        "SALDO", "TRANSPORTE", "DATA REFERENCIA", "DATA REFERÊNCIA", "LIM=", "TAXA=", "CET=",
        "DIA OP DOC", "BDPRVU", "CONTINUA", "FIM DE EXTRATO", "SAC -", "OUVIDORIA",
        "R* =", "TX JR", "TX ULT", "EXT ANT", "DATA ALT",
    ]

    complemento_prefixes = (
        "CPF/CNPJ DESTINO", "CPF/CNPJ REMETENTE", "CPF/CNPJ ",
        "CNPJ-CPF DO REMETENTE", "AGENCIA DESTINO", "AGÊNCIA DESTINO",
        "CONTA DESTINO", "AGENCIA CONTA DESTINO", "AGÊNCIA CONTA DESTINO",
    )

    def split_embedded_transactions(t: str):
        """Divide linhas que contêm complemento + novo lançamento ou dois lançamentos."""
        starts = [m.start() for m in tx_start_re.finditer(t)]
        if not starts:
            return [t]
        parts = []
        last = 0
        for pos in starts:
            if pos > last:
                pre = t[last:pos].strip()
                if pre:
                    parts.append(pre)
            last = pos
        tail = t[last:].strip()
        if tail:
            parts.append(tail)
        return parts or [t]

    def is_complement_line(t: str) -> bool:
        up = t.upper()
        return any(up.startswith(p) for p in complemento_prefixes)

    def clean_trailing_complement(trail: str) -> str:
        trail = norm_space(trail)
        if not trail:
            return ""
        # Sinalizador isolado observado em juros do Banrisul, não é complemento.
        if trail.upper() == "S":
            return ""
        return trail if is_complement_line(trail) else ""

    def append_complement(idx: int, comp: str):
        comp = norm_space(comp)
        if idx is None or idx < 0 or not comp:
            return
        if any(marker in comp.upper() for marker in skip_contains):
            return
        if not is_complement_line(comp):
            return
        # Evita duplicar a mesma informação quando o PDF repetir/mesclar linhas.
        atual = rows[idx][1]
        if comp not in atual:
            rows[idx][1] = norm_space(atual + " | " + comp)

    rows = []
    mes = None
    ano = None
    dia_atual = None
    last_idx = None

    # Expande as linhas antes de interpretar, para lidar com extrações do tipo:
    # "CPF/CNPJ DESTINO ... 1207 300321 BJR ...".
    expanded_lines = []
    for raw in lines:
        t0 = norm_space(raw)
        if not t0:
            continue
        expanded_lines.extend(split_embedded_transactions(t0))

    for t in expanded_lines:
        if not t:
            continue
        up = t.upper()

        mh = month_re.search(up)
        if mh:
            mes = MESES_BAN.get(mh.group(1).upper())
            try:
                ano = int(mh.group(2))
            except Exception:
                ano = None
            dia_atual = None
            last_idx = None
            continue

        if not (mes and ano):
            continue

        if any(marker in up for marker in skip_contains):
            continue

        if is_complement_line(t):
            append_complement(last_idx, t)
            continue

        m = tx_day_re.match(t)
        if m:
            try:
                dia_atual = int(m.group("day"))
            except Exception:
                dia_atual = None
                last_idx = None
                continue
            dia = dia_atual
        else:
            m = tx_noday_re.match(t)
            if not m or dia_atual is None:
                continue
            dia = dia_atual

        try:
            data = datetime(ano, mes, dia).strftime("%d/%m/%Y")
            valor = money_to_float(m.group("val"))
        except Exception:
            last_idx = None
            continue

        descricao = norm_space(f"{m.group('sis')} {m.group('hist')}")
        trailing = clean_trailing_complement(m.group("trail") or "")
        if trailing:
            descricao = norm_space(descricao + " | " + trailing)

        documento = norm_space(m.group("doc"))
        rows.append([data, descricao, documento, valor])
        last_idx = len(rows) - 1

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )


def parse_banrisul_layout3(lines):
    """Parser Banrisul layout III (Portal Banrisul / Extrato meses anteriores).

    Estrutura observada:
      OPERACAO DOCUMENTO VALOR SALDO SIST. AG.ORIGEM
      MOVIMENTOS DA CONTA CORRENTE
      01/09/2023 1166 TED - SPB 000018 20.000,00 BPB
      CPF/CNPJ: 07526557000100
      1005 PAGAMENTO TITULO 741294 116,22- BDX

    Regras:
      - a data pode aparecer apenas na primeira operação do dia;
      - linhas de saldo/cabeçalho/rodapé são ignoradas;
      - o código de operação fica na descrição;
      - o documento vai para a coluna Documento;
      - SIST. e eventual AG.ORIGEM são anexados ao fim da descrição;
      - CPF/CNPJ complementar é anexado à descrição da operação anterior.
    """
    txt = " ".join(lines[:120]).upper()
    if not (
        "MOVIMENTOS DA CONTA CORRENTE" in txt and
        "OPERACAO" in txt and "DOCUMENTO" in txt and "SIST" in txt
    ):
        return standardize(pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor"]))

    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_tok_re = r"\d{1,3}(?:\.\d{3})*,\d{2}-?"
    tx_re = re.compile(
        rf"^(?:(?P<data>\d{{2}}/\d{{2}}/\d{{4}})\s+)?"
        rf"(?P<op>\d{{4}})\s+"
        rf"(?P<hist>.+?)\s+"
        rf"(?P<doc>[A-Z0-9]{{1,12}})\s+"
        rf"(?P<valor>{money_tok_re})\s+"
        rf"(?P<sist>[A-Z]{{3}})(?:\s+(?P<agorig>\d{{3,6}}))?$",
        re.I
    )

    rows = []
    current_date = None
    in_movements = False

    skip_prefixes = (
        "BANCO DO ESTADO", "BVP-PORTAL", "EXTRATO MESES", "PÁGINA", "PAGINA",
        "DATA REF", "DATA PROC", "HORA PROC", "DADOS SELECIONADOS", "AGÊNCIA", "AGENCIA",
        "NOME", "DATA ABERTURA", "TIPO DE EXTRATO", "EXTRA-CONTÁBIL", "EXTRA-CONTABIL",
        "LIMITE", "PERIODO", "OPERACAO DOCUMENTO", "+---", "----", "++ MOVIMENTOS", "U=",
        "MOVIMENTOS DA CONTA CORRENTE"
    )

    for raw_ln in lines:
        ln = norm_space(raw_ln)
        if not ln:
            continue
        up = ln.upper()

        if "MOVIMENTOS DA CONTA CORRENTE" in up:
            in_movements = True
            continue
        if not in_movements:
            continue

        if up.startswith(skip_prefixes):
            continue
        if up.startswith("SALDO") or "SALDO NA DATA" in up or "SALDO ANT" in up:
            continue

        # Linha complementar de CPF/CNPJ não é transação; incorpora à descrição
        # da operação imediatamente anterior. Usamos "CPF/CNPJ " sem dois-pontos
        # para evitar colisão com filtros globais de cabeçalho em standardize().
        if up.startswith("CPF/CNPJ"):
            if rows:
                comp = re.sub(r"^CPF/CNPJ\s*[:.-]?\s*", "", ln, flags=re.I).strip()
                if comp:
                    desc_atual = rows[-1][1]
                    cpf_txt = f"CPF/CNPJ {comp}"
                    if cpf_txt not in desc_atual:
                        # Insere o complemento ainda no núcleo da descrição, antes
                        # dos sufixos padronizados "; Sistema Banrisul ..." e
                        # "; Ag. Origem ...".
                        if "; Sistema Banrisul" in desc_atual:
                            base, resto = desc_atual.split("; Sistema Banrisul", 1)
                            rows[-1][1] = norm_space(base + " " + cpf_txt + "; Sistema Banrisul" + resto)
                        else:
                            rows[-1][1] = norm_space(desc_atual + " " + cpf_txt)
            continue

        m = tx_re.match(ln)
        if not m:
            continue

        dt = m.group("data")
        if dt:
            current_date = dt
        if not current_date or not date_re.match(current_date):
            continue

        op = norm_space(m.group("op"))
        hist = norm_space(m.group("hist"))
        doc = norm_space(m.group("doc"))
        sist = norm_space(m.group("sist") or "")
        agorig = norm_space(m.group("agorig") or "")
        valor = money_to_float(m.group("valor"))

        desc_parts = [norm_space(f"{op} {hist}".strip())]
        if sist:
            desc_parts.append(f"Sistema Banrisul {sist}")
        if agorig:
            desc_parts.append(f"Ag. Origem {agorig}")
        descricao = norm_space("; ".join([p for p in desc_parts if p]))

        if descricao and not is_balance_or_summary_line(descricao):
            rows.append([current_date, descricao, doc, valor])

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )

def parse_banrisul(lines):
    # Primeiro tenta o layout III por assinatura estrutural própria.
    # Ele também contém linha "PERIODO:", mas sua tabela não é a do layout I.
    df_layout3 = parse_banrisul_layout3(lines)
    if not df_layout3.empty:
        return df_layout3

    # Depois tenta o layout Banrisul já estabilizado na v51.
    df_layout1 = parse_banrisul_layout1(lines)
    if not df_layout1.empty:
        return df_layout1

    # Se o layout antigo não se aplica, tenta o layout II.
    return parse_banrisul_layout2(lines)



# ---------------- Sicredi layout 3 - CCPI ----------------


def is_sicredi_layout3_ccpi_signature(txt: str, name: str = "") -> bool:
    """Reconhece o extrato anual da CCPI Sul Riograndense.

    Assinatura observada:
      - "CCPI SUL RIOGRANDENSE EXTRATO DE CONTA CORRENTE";
      - tabela DATA | DOCUMENTO | HISTORICO | DEBITO | CREDITO | SALDO;
      - linhas de transporte entre páginas.

    O nome de layout usado no dispatcher é ``sicredi_layout3``.
    """
    t = normalize_classification_text(txt)
    nm = normalize_classification_text(name)
    return (
        (
            "CCPI SUL RIOGRANDENSE" in t
            and "EXTRATO DE CONTA CORRENTE" in t
            and "DATA DOCUMENTO HISTORICO DEBITO CREDITO SALDO" in t
        )
        or (
            "SICREDI_LAYOUT3" in nm
            and "CCPI" in nm
            and "DATA DOCUMENTO HISTORICO DEBITO CREDITO SALDO" in t
        )
    )


def _ccpi_collapse_doubled_token(value: str) -> str:
    """Corrige glifos duplicados que aparecem em algumas linhas desse PDF.

    Exemplos reais do arquivo:
      0022//0033//22002222 -> 02/03/2022
      PPIIXX__CCRREEDD     -> PIX_CRED
      22..333333,,0000     -> 2.333,00

    Só colapsa quando TODO o token é composto por pares idênticos, evitando
    alterações em textos/números normais.
    """
    s = str(value or "")
    # Em linhas integralmente duplicadas, palavras de uma só letra aparecem como
    # "VV", "AA" etc. Colapsa também esse caso alfabético específico.
    if len(s) == 2 and s.isalpha() and s[0] == s[1]:
        return s[0]
    if len(s) >= 4 and len(s) % 2 == 0:
        if all(s[i] == s[i + 1] for i in range(0, len(s), 2)):
            return "".join(s[i] for i in range(0, len(s), 2))
    return s


def parse_sicredi_layout3_ccpi(pdf_path, pdf_password=None):
    """Parser dedicado ao Sicredi_layout3 / CCPI.

    Estrutura da origem:
        DATA | DOCUMENTO | HISTORICO | DEBITO | CREDITO | SALDO

    Regras:
      - cada data representa uma transação;
      - o sinal é inferido pela coluna física: DÉBITO negativo, CRÉDITO positivo;
      - SALDO é usado apenas para validação visual e não vira transação;
      - SALDO ANTERIOR e DE TRANSPORTE não têm data válida e são ignorados;
      - cabeçalhos/rodapés repetidos são naturalmente excluídos;
      - Documento e Histórico são preservados em seus respectivos campos.
    """
    if pdfplumber is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []

    def group_words_by_line(words, tol=1.2):
        ordered = sorted(words, key=lambda w: (float(w["top"]), float(w["x0"])))
        groups = []
        current = []
        current_top = None
        for w in ordered:
            top = float(w["top"])
            if current_top is None or abs(top - current_top) <= tol:
                current.append(w)
                current_top = top if current_top is None else (current_top + top) / 2.0
            else:
                groups.append(sorted(current, key=lambda z: float(z["x0"])))
                current = [w]
                current_top = top
        if current:
            groups.append(sorted(current, key=lambda z: float(z["x0"])))
        return groups

    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=1,
                y_tolerance=2,
                keep_blank_chars=False,
            ) or []
            if not words:
                continue

            # Algumas linhas possuem a camada textual duplicada caractere a caractere.
            for w in words:
                w["_ccpi_text"] = _ccpi_collapse_doubled_token(w.get("text", ""))

            groups = group_words_by_line(words)

            # Obtém as coordenadas das colunas a partir do cabeçalho da própria página.
            header = None
            for group in groups:
                coords = {
                    str(w["_ccpi_text"]).upper(): float(w["x0"])
                    for w in group
                }
                required = {"DATA", "DOCUMENTO", "HISTORICO", "DEBITO", "CREDITO", "SALDO"}
                if required.issubset(coords):
                    header = coords
                    break
            if not header:
                continue

            doc_x = header["DOCUMENTO"]
            hist_x = header["HISTORICO"]
            debit_x = header["DEBITO"]
            credit_x = header["CREDITO"]
            saldo_x = header["SALDO"]

            # Fronteiras entre as colunas monetárias. Como há pequenas variações de x
            # entre linhas, usa os pontos médios dos títulos em vez de valores fixos.
            debit_credit_mid = (debit_x + credit_x) / 2.0
            credit_saldo_mid = (credit_x + saldo_x) / 2.0
            money_scan_start = (hist_x + debit_x) / 2.0

            for group in groups:
                first = str(group[0].get("_ccpi_text", "")) if group else ""
                if not date_re.match(first):
                    continue

                money_words = [
                    w for w in group
                    if float(w["x0"]) >= money_scan_start
                    and money_re.match(str(w.get("_ccpi_text", "")))
                ]
                if not money_words:
                    continue

                debit = 0.0
                credit = 0.0
                has_debit = False
                has_credit = False
                first_money_x = min(float(w["x0"]) for w in money_words)

                for w in sorted(money_words, key=lambda z: float(z["x0"])):
                    token = str(w["_ccpi_text"])
                    x0 = float(w["x0"])
                    amount = abs(money_to_float(token))
                    if x0 < debit_credit_mid:
                        debit += amount
                        has_debit = True
                    elif x0 < credit_saldo_mid:
                        credit += amount
                        has_credit = True
                    else:
                        # Coluna SALDO: deliberadamente ignorada na transação.
                        pass

                if not has_debit and not has_credit:
                    continue

                # Em uma linha normal existe apenas débito OU crédito. A expressão
                # abaixo também é estável se surgir uma linha atípica com ambos.
                valor = credit - debit
                if valor == 0:
                    continue

                doc_parts = [
                    str(w["_ccpi_text"])
                    for w in group
                    if doc_x - 3 <= float(w["x0"]) < hist_x - 2
                ]
                hist_parts = [
                    str(w["_ccpi_text"])
                    for w in group
                    if hist_x - 2 <= float(w["x0"]) < first_money_x - 2
                    and not money_re.match(str(w["_ccpi_text"]))
                ]

                documento = norm_space(" ".join(doc_parts))
                historico = norm_space(" ".join(hist_parts))
                if not historico:
                    continue

                rows.append([first, historico, documento, valor])

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        # O campo DOCUMENTO do CCPI também usa códigos curtos como "1", PIX_DEB,
        # PIX_CRED e DAS; por isso não se aplica o filtro documental genérico.
        doc_cleaner=lambda x: norm_space(str(x))[:80],
    )


# ---------------- Sicredi ----------------


def parse_sicredi_table(pdf_path, pdf_password=None):
    if pdfplumber is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    def group_words_by_line(words, tol=1.2):
        words = sorted(words, key=lambda w: (w["top"], w["x0"]))
        groups = []
        current = []
        current_top = None
        for w in words:
            top = float(w["top"])
            if current_top is None or abs(top - current_top) <= tol:
                current.append(w)
                current_top = top if current_top is None else (current_top + top) / 2.0
            else:
                groups.append(sorted(current, key=lambda z: z["x0"]))
                current = [w]
                current_top = top
        if current:
            groups.append(sorted(current, key=lambda z: z["x0"]))
        return groups

    rows = []
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")

    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=2, keep_blank_chars=False) or []
            for group in group_words_by_line(words):
                texts = [w["text"] for w in group]
                if not texts:
                    continue

                first = texts[0]
                if first == "**/**/****":
                    continue
                if not date_re.match(first):
                    continue

                date = first
                doc_parts = [w["text"] for w in group if 78 <= w["x0"] < 125]
                hist_parts = [w["text"] for w in group if 125 <= w["x0"] < 365]
                money_words = [w for w in group if w["x0"] >= 365 and money_re.match(w["text"])]
                money_words = sorted(money_words, key=lambda z: z["x0"])

                doc = norm_space(" ".join(doc_parts))
                hist = norm_space(" ".join(hist_parts))

                if doc and not clean_document_token_sicredi(doc):
                    hist = norm_space((doc + " " + hist).strip())
                    doc = ""

                if not hist:
                    continue

                hist_low = hist.lower()
                if any(x in hist_low for x in ["saldo anterior", "d e t r a n s p o r t e", "saldo total", "saldo do dia"]):
                    continue
                if doc.lower() in {"data", "documento", "historico", "debito", "credito", "saldo"}:
                    continue

                if not money_words:
                    continue

                amount_word = money_words[0]
                # quando houver duas quantias, a segunda costuma ser o saldo
                if len(money_words) >= 2:
                    amount_word = money_words[0]

                amount_txt = amount_word["text"]
                if amount_word["x0"] < 450:
                    val = -abs(money_to_float(amount_txt))
                else:
                    val = abs(money_to_float(amount_txt))

                rows.append([date, hist, doc, val])

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
                      doc_cleaner=clean_document_token_sicredi)


def parse_sicredi(lines, pdf_path=None, pdf_password=None):
    joined = " ".join(lines[:120]).lower()
    if pdf_path and "data documento historico debito credito saldo" in joined and "extrato de conta corrente" in joined:
        return parse_sicredi_table(pdf_path, pdf_password=pdf_password)

    return parse_sicredi_generic(lines)

def parse_sicredi_generic(lines):
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


def _extract_trailing_doc_from_desc(desc: str, doc: str = ""):
    desc = norm_space(desc)
    doc = norm_space(doc)
    if doc:
        return desc, doc
    m = re.match(r"^(?P<desc>.+?)\s+(?P<doc>\d{4,8})$", desc)
    if m:
        cand_doc = m.group("doc")
        cand_desc = norm_space(m.group("desc"))
        # Evita mover ano de competência para Documento em expressões como "OUTUBRO / 2021".
        if re.fullmatch(r"(?:19|20)\d{2}", cand_doc) and cand_desc.rstrip().endswith("/"):
            return desc, doc
        if cand_desc.upper().endswith("FINAL"):
            return desc, doc
        return cand_desc, clean_document_token_flexible(cand_doc)
    return desc, doc


def parse_santander_layout1_from_lines(lines):
    """Parser rápido para Santander layout1/3, inclusive PDFs consolidados com vários meses.

    Usa as linhas já extraídas por PyMuPDF. O PDF costuma vir em fluxo vertical:
    Data, descrição, documento ou '-', valor, saldo opcional. O parser percorre todos
    os blocos "Movimentação" de conta corrente/contamax e ignora seções posteriores
    como Saldos por Período, investimentos, cartões etc.
    """
    rows = []
    re_resumo = re.compile(r"resumo\s*-\s*([A-Za-zçÇ]+)/\s*(\d{4})", re.I)
    re_mes_ano = re.compile(r"^([A-Za-zçÇ]+)/(\d{4})$", re.I)
    re_date = re.compile(r"^\d{2}/\d{2}$")
    money_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}-?$")
    # documento pode ser número, CPF/CNPJ, cheque, código etc.; '-' é placeholder da coluna.
    doc_re = re.compile(r"^(?:-|\d{3,}|\d{1,3}(?:\.\d{3}){1,2}-?\d{0,2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})$")

    def clean_doc_sant(doc):
        doc = norm_space(doc)
        if not doc or doc == "-":
            return ""
        return doc

    def is_header_or_noise(line):
        low = line.lower()
        if not line:
            return True
        if low.startswith("extrato_pj_") or low.startswith("balp_uy_") or low.startswith("pagina:"):
            return True
        if low in {"data", "descrição", "descricao", "nº documento", "no documento", "movimentos (r$)", "saldo (r$)", "créditos", "creditos", "débitos", "debitos"}:
            return True
        if low in {"conta corrente", "contamax", "movimentação", "movimentacao"}:
            return True
        if line.startswith('“') or line.startswith('"'):
            return True
        if any(x in low for x in [
            "este demonstrativo", "central de atendimento", "sac", "ouvidoria", "redes sociais", "fale conosco",
            "adiantamento a depositantes", "saldos por período", "saldos por periodo",
            "relação de cheques", "relacao de cheques", "débito automático", "debito automatico",
            "compras com cartão", "compras com cartao", "créditos contratados", "creditos contratados",
            "cdb / rdb", "aplicação n°", "aplicacao n°", "valor principal", "valor bruto", "valor líquido", "valor liquido",
            "movimentação mensal", "movimentacao mensal"
        ]):
            return True
        return False

    def parse_inline_transaction(text_line):
        # Caso em que PDF traga tudo na mesma linha: descrição + doc opcional + valor + saldo opcional.
        m = re.match(
            r"^(?P<desc>.+?)\s+(?:(?P<doc>\d{4,}|-)\s+)?(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)(?:\s+\d{1,3}(?:\.\d{3})*,\d{2}-?)?$",
            text_line
        )
        if not m:
            return None
        desc = norm_space(m.group('desc').strip(' -'))
        if not re.search(r"[A-Za-zÀ-ÿ]", desc):
            return None
        doc = clean_doc_sant(m.group('doc') or "")
        desc, doc = _extract_trailing_doc_from_desc(desc, doc)
        val = money_to_float(m.group('val'))
        return desc, doc, val

    # fallback de ano: primeiro padrão mês/ano encontrado no arquivo
    fallback_year = "2024"
    for line in lines[:300]:
        m = re.search(r"[A-Za-zçÇ]+/(\d{4})", line, re.I)
        if m:
            fallback_year = m.group(1)
            break

    year_ctx = None
    current_date = None
    in_mov = False
    i = 0
    N = len(lines)

    while i < N:
        line = norm_space(lines[i])
        low = line.lower()

        mr = re_resumo.search(line)
        if mr:
            year_ctx = mr.group(2)
            in_mov = False
            current_date = None
            i += 1
            continue
        mm = re_mes_ano.match(line)
        if mm:
            year_ctx = mm.group(2)
            i += 1
            continue

        if low in {"movimentação", "movimentacao"}:
            nxt = " ".join(lines[i + 1:i + 12]).lower()
            # entra apenas no bloco transacional da conta corrente/contamax
            if "nº documento" in nxt and "movimentos (r$)" in nxt and "saldo (r$)" in nxt:
                in_mov = True
                current_date = None
            i += 1
            continue

        if "saldos por período" in low or "saldos por periodo" in low:
            in_mov = False
            current_date = None
            i += 1
            continue

        if not in_mov:
            i += 1
            continue

        if is_header_or_noise(line):
            i += 1
            continue

        if low.startswith("saldo em "):
            # linha de saldo inicial/final; pula também o valor seguinte, se vier sozinho
            i += 1
            if i < N and money_re.match(norm_space(lines[i])):
                i += 1
            continue

        if re_date.match(line):
            current_date = f"{line}/{year_ctx or fallback_year}"
            i += 1
            continue

        if current_date is None:
            i += 1
            continue

        inline = parse_inline_transaction(line)
        if inline:
            desc, doc, val = inline
            if desc and not is_balance_or_summary_line(desc):
                rows.append([current_date, desc, doc, val])
            i += 1
            continue

        # Fluxo vertical: descrição em uma ou mais linhas, doc/placeholder, valor, saldo opcional.
        desc_parts = []
        doc = ""
        while i < N:
            cur = norm_space(lines[i])
            cur_low = cur.lower()
            if re_date.match(cur) or "saldos por período" in cur_low or "saldos por periodo" in cur_low:
                break
            if is_header_or_noise(cur):
                i += 1
                continue
            if cur_low.startswith("saldo em "):
                break

            # doc/placeholder imediatamente antes do valor.
            # Exceção: anos isolados em complemento de histórico, como "OUTUBRO / 2021".
            if doc_re.match(cur) and i + 1 < N and money_re.match(norm_space(lines[i + 1])):
                if re.fullmatch(r"(?:19|20)\d{2}", cur) and desc_parts and desc_parts[-1].rstrip().endswith("/"):
                    desc_parts.append(cur)
                    i += 1
                    continue
                doc = clean_doc_sant(cur)
                i += 1
                continue

            if money_re.match(cur):
                desc = norm_space(" ".join(desc_parts))
                desc, doc = _extract_trailing_doc_from_desc(desc, doc)
                val = money_to_float(cur)
                if desc and not is_balance_or_summary_line(desc):
                    rows.append([current_date, desc, doc, val])
                i += 1
                # saldo opcional após o valor da transação
                if i < N and money_re.match(norm_space(lines[i])):
                    i += 1
                break

            desc_parts.append(cur)
            i += 1
        else:
            i += 1
            continue

        # se não consumiu nada por alguma quebra, avança para evitar loop
        if i < N and (re_date.match(norm_space(lines[i])) or "saldos por período" in norm_space(lines[i]).lower() or "saldos por periodo" in norm_space(lines[i]).lower()):
            continue

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
                      doc_cleaner=clean_document_token_flexible)


def parse_santander_layout1_multiblock_fitz(pdf_path, pdf_password=None):
    """Parser complementar para PDFs Santander layout1/3 muito longos,
    nos quais cada página pode iniciar com continuação de movimentação sem repetir a data.
    Usa PyMuPDF por desempenho em documentos extensos.
    """
    re_resumo = re.compile(r"resumo\s*-\s*([A-Za-zçÇ]+)/\s*(\d{4})", re.I)
    re_mes_ano = re.compile(r"^([A-Za-zçÇ]+)/(\d{4})$", re.I)
    re_date = re.compile(r'^(?P<data>\d{2}/\d{2})(?:\s+(?P<rest>.+))?$')
    money_pat = r'\d{1,3}(?:\.\d{3})*,\d{2}-?'
    re_money = re.compile(rf'^{money_pat}$')
    re_doc = re.compile(r'^(?:\d{4,}|-|[0-9]{3}/[0-9]{4,})$')
    re_tx = re.compile(rf'^(?P<desc>.+?)\s+(?P<doc>\d{{4,}}|-)\s+(?P<val>{money_pat})(?:\s+(?P<saldo>{money_pat}))?$')
    re_tx_nodoc = re.compile(rf'^(?P<desc>.+?)\s+(?P<val>{money_pat})(?:\s+(?P<saldo>{money_pat}))?$')
    tx_prefix = re.compile(
        r"^(PIX|TRANSF|TRANSFER|TARIFA|TED|DOC|PAGAMENTO|PGTO|A CR|CR COB|RESGATE|APLICACAO|APLICAÇÃO|"
        r"ANTECIPACAO|ANTECIPAÇÃO|COMPRA|DEBITO|DÉBITO|MENSALIDADE|PREST|IOF|JUROS|OPERACAO|OPERAÇÃO|"
        r"CHEQUE|ESTORNO)",
        re.I
    )

    def noise(line):
        low = line.lower()
        if not line:
            return True
        if low.startswith(("extrato_pj", "balp_uy", "pagina:")):
            return True
        if low in {
            "extrato consolidado inteligente", "data", "descrição", "descricao", "nº documento", "no documento",
            "movimentos (r$)", "saldo (r$)", "créditos", "creditos", "débitos", "debitos",
            "conta corrente", "contamax", "movimentação", "movimentacao"
        }:
            return True
        if line.startswith(("“", '"')):
            return True
        if any(x in low for x in [
            "este demonstrativo", "central de atendimento", "sac", "ouvidoria", "redes sociais",
            "fale conosco", "relação de cheques", "relacao de cheques", "débito automático",
            "debito automatico", "compras com cartão", "compras com cartao",
            "créditos contratados", "creditos contratados", "limite santander master"
        ]):
            return True
        return False

    rows = []
    in_mov = False
    year_ctx = None
    current_date = None
    pending = None

    def finish(desc, doc, val, date):
        desc = norm_space(desc)
        doc = norm_space(doc)
        if doc == "-":
            doc = ""
        desc, doc = _extract_trailing_doc_from_desc(desc, doc)
        if desc and not is_balance_or_summary_line(desc):
            rows.append([date, desc, doc, val])

    def parse_complete(line, date):
        m = re_tx.match(line)
        if m:
            finish(m.group("desc"), m.group("doc") or "", money_to_float(m.group("val")), date)
            return True
        m = re_tx_nodoc.match(line)
        if m and re.search(r"[A-Za-zÀ-ÿ]", m.group("desc")):
            finish(m.group("desc"), "", money_to_float(m.group("val")), date)
            return True
        return False

    def looks_start(rest):
        return bool(rest and (tx_prefix.search(rest) or re.search(rf'\s{money_pat}(?:\s+{money_pat})?$', rest)))

    doc, _pw = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            lines = [norm_space(l) for l in (page.get_text("text") or "").splitlines()]
            lines = [l for l in lines if l]
            for line in lines:
                low = line.lower()

                mr = re_resumo.search(line)
                if mr:
                    year_ctx = mr.group(2)
                    in_mov = False
                    current_date = None
                    pending = None
                    continue

                mm = re_mes_ano.match(line)
                if mm:
                    year_ctx = mm.group(2)
                    continue

                if low in {"movimentação", "movimentacao"}:
                    in_mov = True
                    pending = None
                    continue

                if "saldos por período" in low or "saldos por periodo" in low:
                    in_mov = False
                    current_date = None
                    pending = None
                    continue

                if not in_mov:
                    continue

                if noise(line):
                    continue

                if low.startswith("saldo em "):
                    pending = None
                    continue

                md = re_date.match(line)
                if md and year_ctx:
                    candidate_date = f"{md.group('data')}/{year_ctx}"
                    rest = (md.group("rest") or "").strip()
                    if not rest:
                        current_date = candidate_date
                        pending = None
                        continue
                    if looks_start(rest):
                        current_date = candidate_date
                        pending = None
                        if not parse_complete(rest, current_date):
                            pending = {"date": current_date, "parts": [rest], "doc": ""}
                        continue
                    # Linha do tipo "16/09 BOURBON COUNTRY" costuma ser complemento de compra,
                    # não nova data de lançamento. Não altera current_date.
                    line = rest

                if current_date is None:
                    continue

                if pending is not None:
                    if re_money.match(line):
                        finish(" ".join(pending["parts"]), pending.get("doc", ""), money_to_float(line), pending["date"])
                        pending = None
                        continue
                    if re_doc.match(line) and not pending.get("doc"):
                        pending["doc"] = line
                        continue
                    if not noise(line):
                        pending["parts"].append(line)
                        continue

                if parse_complete(line, current_date):
                    pending = None
                    continue

                if re.search(r"[A-Za-zÀ-ÿ]", line):
                    pending = {"date": current_date, "parts": [line], "doc": ""}

    finally:
        doc.close()

    df = pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"])
    return standardize(df, doc_cleaner=clean_document_token_flexible)


def parse_santander_layout1_from_pdf(pdf_path, pdf_password=None):
    import pdfplumber, re, pandas as pd

    # Documentos Santander muito longos podem trazer várias competências e páginas
    # que começam com continuação de movimentação sem repetir a data.
    # Para esses casos, usa parser page-aware em PyMuPDF.
    try:
        _doc_tmp, _pw_tmp = open_pdf_with_password(pdf_path)
        _page_count = len(_doc_tmp)
        _doc_tmp.close()
        if _page_count > 80:
            return parse_santander_layout1_multiblock_fitz(pdf_path, pdf_password=pdf_password)
    except Exception:
        pass

    texto = ""
    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    linhas = [re.sub(r"\s{2,}", " ", l.strip()) for l in texto.splitlines() if l.strip()]

    # Este layout pode aparecer como extrato de mês único ou como PDF consolidado
    # com vários meses/anos. A versão antiga parava no primeiro "Saldos por Período";
    # agora varremos todos os blocos "Movimentação" do documento.
    re_resumo = re.compile(r"resumo\s*-\s*([A-Za-zçÇ]+)/\s*(\d{4})", re.I)
    re_mes_ano = re.compile(r"^([A-Za-zçÇ]+)/(\d{4})$", re.I)
    re_date = re.compile(r'^(?P<data>\d{2}/\d{2})(?:\s+(?P<rest>.+))?$')
    re_tx = re.compile(
        r'^(?P<desc>.+?)'
        r'(?:\s+(?P<doc>\d{4,}|-))?'
        r'\s+(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)'
        r'(?:\s+(?P<saldo>\d{1,3}(?:\.\d{3})*,\d{2}-?))?$'
    )
    re_tail_value = re.compile(
        r'^(?:(?P<doc>\d{4,}|-)\s+)?'
        r'(?P<val>\d{1,3}(?:\.\d{3})*,\d{2}-?)'
        r'(?:\s+(?P<saldo>\d{1,3}(?:\.\d{3})*,\d{2}-?))?$'
    )

    def is_noise_line(line: str) -> bool:
        low = line.lower()
        if not line:
            return True
        if low.startswith("extrato_pj_") or low.startswith("balp_uy_") or low.startswith("pagina:"):
            return True
        if low == "extrato consolidado inteligente":
            return True
        if low in {
            "data descrição nº documento movimentos (r$) saldo (r$)",
            "data descricao nº documento movimentos (r$) saldo (r$)",
            "data descrição no documento movimentos (r$) saldo (r$)",
            "créditos débitos", "creditos debitos", "créditos", "creditos", "débitos", "debitos",
            "conta corrente", "contamax", "movimentação", "movimentacao"
        }:
            return True
        if line.startswith('“') or line.startswith('"'):
            return True
        if any(x in low for x in [
            "este demonstrativo", "central de atendimento", "sac", "ouvidoria", "redes sociais",
            "adiantamento a depositantes", "jurosmoratórios", "jurosmoratorios", "produtocontratado",
            "saldo devedor", "sujeito à cobrança", "sujeito a cobranca", "fale conosco",
            "relação de cheques", "relacao de cheques", "débito automático", "debito automatico",
            "compras com cartão", "compras com cartao", "créditos contratados", "creditos contratados"
        ]):
            return True
        return False

    def finish_pending(pending, doc, val):
        if not pending:
            return
        desc = norm_space(" ".join(pending["parts"]))
        doc = norm_space(doc or pending.get("doc", ""))
        if doc == "-":
            doc = ""
        desc, doc = _extract_trailing_doc_from_desc(desc, doc)
        if desc and not is_balance_or_summary_line(desc):
            rows.append([pending["date"], desc, doc, val])

    rows = []
    in_mov = False
    year_ctx = None
    current_date = None
    pending = None

    # fallback: primeiro mês/ano presente no documento
    mhead = re.search(r'([A-Za-zçÇ]+)/(\d{4})', texto, re.I)
    fallback_year = mhead.group(2) if mhead else "2024"

    for idx_line, raw_line in enumerate(linhas):
        line = norm_space(raw_line)
        low = line.lower()

        mr = re_resumo.search(line)
        if mr:
            year_ctx = mr.group(2)
            in_mov = False
            current_date = None
            pending = None
            continue
        mm = re_mes_ano.match(line)
        if mm:
            year_ctx = mm.group(2)
            continue

        if low in {"movimentação", "movimentacao"}:
            # Só entra no bloco transacional de conta corrente/contamax.
            # O PDF também contém seções de investimentos com o título "Movimentação",
            # mas nelas o cabeçalho é outro (Valor Principal, Valor Bruto etc.).
            next_block = " ".join(linhas[idx_line + 1: idx_line + 8]).lower()
            if "data descrição nº documento movimentos" in next_block or "data descricao nº documento movimentos" in next_block:
                in_mov = True
                current_date = None
                pending = None
            continue

        if "saldos por período" in low or "saldos por periodo" in low:
            in_mov = False
            current_date = None
            pending = None
            continue

        if not in_mov:
            continue

        if is_noise_line(line):
            continue
        if low.startswith("saldo em "):
            pending = None
            continue

        mdate = re_date.match(line)
        if mdate:
            current_date = f"{mdate.group('data')}/{year_ctx or fallback_year}"
            rest = (mdate.group('rest') or "").strip()
            pending = None
            if not rest:
                continue

            mt = re_tx.match(rest)
            if mt:
                desc = mt.group('desc').strip(" -")
                doc = mt.group('doc') or ""
                if doc == "-":
                    doc = ""
                desc, doc = _extract_trailing_doc_from_desc(desc, doc)
                val = money_to_float(mt.group('val'))
                rows.append([current_date, desc, doc, val])
            else:
                pending = {"date": current_date, "parts": [rest], "doc": ""}
            continue

        if current_date is None:
            continue

        mt = re_tx.match(line)
        if mt and re.search(r'[A-Za-zÀ-ÿ]', mt.group('desc') or ''):
            pending = None
            desc = mt.group('desc').strip(" -")
            doc = mt.group('doc') or ""
            if doc == "-":
                doc = ""
            desc, doc = _extract_trailing_doc_from_desc(desc, doc)
            val = money_to_float(mt.group('val'))
            rows.append([current_date, desc, doc, val])
            continue

        # continuação de uma descrição iniciada em linha com data, ou em linha anterior
        if pending is not None:
            mv = re_tail_value.match(line)
            if mv:
                doc = mv.group("doc") or ""
                val = money_to_float(mv.group("val"))
                finish_pending(pending, doc, val)
                pending = None
                continue
            # linha de complemento de histórico/documento
            if not is_noise_line(line) and not re.fullmatch(r'\d{1,3}(?:\.\d{3})*,\d{2}-?', line):
                pending["parts"].append(line)
            continue

        # linha de continuação sem data, comum após a primeira transação do dia
        # Ex.: descrição em uma linha, documento em outra e valor na seguinte.
        if re.search(r'[A-Za-zÀ-ÿ]', line):
            pending = {"date": current_date, "parts": [line], "doc": ""}
            continue

    df = pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"])
    df = df[df["Valor"].notna()].copy()
    return standardize(df, doc_cleaner=clean_document_token_flexible)

def is_santander_layout4_signature(txt: str, name: str = "") -> bool:
    """Reconhece o Santander layout4 (Internet Banking Empresarial / tabela Data-Histórico-Documento-Valor-Saldo)."""
    sample = norm_space(str(txt or "")).lower()
    nm = str(name or "").lower()
    has_header = (
        "internet banking empresarial" in sample
        and "conta corrente > extratos" in sample
        and ("opção de pesquisa" in sample or "opcao de pesquisa" in sample)
    )
    has_columns = (
        "data histórico documento valor (r$) saldo (r$)" in sample
        or "data historico documento valor (r$) saldo (r$)" in sample
    )
    return (has_header and has_columns) or (nm.startswith("santander_layout4") and has_columns)


def parse_santander_layout4(lines):
    """Parser do Santander layout4.

    Estrutura observada:
      Data | Histórico | Documento | Valor (R$) | Saldo (R$)

    Regras:
      - cada transação começa com data completa dd/mm/aaaa;
      - o sinal do Valor identifica débito/crédito;
      - o primeiro registro "SALDO ANTERIOR" é ignorado;
      - a coluna Saldo, quando presente, é ignorada;
      - o quadro final de composição de saldo/limite é ignorado;
      - linhas de saldo informativo entre páginas não são tratadas como transações.
    """
    rows = []
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    doc_re = re.compile(r"^(?:\d{4,}|[A-Za-z0-9][A-Za-z0-9./_-]{2,30})$")

    # A tabela começa somente após o cabeçalho completo. Isso evita interpretar
    # datas do período/horário do cabeçalho como lançamentos.
    table_start = None
    for i in range(max(0, len(lines) - 4)):
        seq = [norm_space(x).lower() for x in lines[i:i + 5]]
        if seq == ["data", "histórico", "documento", "valor (r$)", "saldo (r$)"] or seq == ["data", "historico", "documento", "valor (r$)", "saldo (r$)"]:
            table_start = i + 5
            break

    if table_start is None:
        return standardize(pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor"]))

    def is_final_balance_section(line: str) -> bool:
        low = norm_space(line).lower()
        return (
            low.startswith("a = bloqueio dia")
            or low.startswith("entenda a composição do seu saldo")
            or low.startswith("entenda a composicao do seu saldo")
            or low.startswith("posição em:")
            or low.startswith("posicao em:")
        )

    i = table_start
    N = len(lines)
    while i < N:
        line = norm_space(lines[i])
        if is_final_balance_section(line):
            break

        if not date_re.match(line):
            i += 1
            continue

        data = line
        j = i + 1
        block = []
        while j < N:
            cur = norm_space(lines[j])
            if date_re.match(cur) or is_final_balance_section(cur):
                break
            block.append(cur)
            j += 1

        # O primeiro valor monetário do bloco é o Valor da transação; eventual
        # segundo valor é o Saldo (R$) e deve ser desprezado.
        amount_idx = None
        for k, item in enumerate(block):
            if money_re.match(item):
                amount_idx = k
                break

        if amount_idx is not None:
            before_value = [x for x in block[:amount_idx] if norm_space(x)]
            valor_txt = block[amount_idx]

            documento = ""
            desc_parts = before_value
            if len(before_value) >= 2 and doc_re.match(before_value[-1]):
                documento = norm_space(before_value[-1])
                desc_parts = before_value[:-1]

            descricao = norm_space(" ".join(desc_parts))
            if descricao and descricao.upper() != "SALDO ANTERIOR" and not is_balance_or_summary_line(descricao):
                try:
                    valor = money_to_float(valor_txt)
                except Exception:
                    valor = None
                if valor is not None and valor != 0:
                    rows.append([data, descricao, documento, valor])

        i = j

    df = pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"])

    # standardize() possui um filtro global de cabeçalhos com a expressão "período:".
    # Neste layout, porém, ela pode integrar legitimamente o histórico de uma transação
    # (ex.: "IOF ADICIONAL - AUTOMATICO PERIODO: 01/12 A 31/12/20").
    # Usa-se um marcador apenas durante a padronização e restaura-se o texto original
    # em seguida, sem alterar a lógica global compartilhada pelos demais parsers.
    sentinel_periodo = "__SANT4_PERIODO_COLON__"
    if not df.empty:
        df["Descrição"] = df["Descrição"].astype(str).str.replace(
            r"(?i)PER[IÍ]ODO:", sentinel_periodo, regex=True
        )

    out = standardize(
        df,
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )
    if not out.empty:
        out["Descrição"] = out["Descrição"].astype(str).str.replace(
            sentinel_periodo, "PERIODO:", regex=False
        )
    return out


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


def parse_santander(pdf_path, lines, pdf_password=None):
    re_sant_day_local = re.compile(
        r"^(Segunda|Terça|Terca|Quarta|Quinta|Sexta|Sábado|Sabado|Domingo),\s+(\d{1,2})\s+de\s+([A-Za-zç]+)\s+de\s+(\d{4})$",
        re.I
    )
    if any(re_sant_day_local.match(ln) for ln in lines[:200]):
        return parse_santander_layout2(lines)

    # PDFs Santander layout1/3 com mês único continuam no parser antigo, já estabilizado.
    # PDFs consolidados com vários meses/anos usam o parser por linhas, que percorre todos
    # os blocos de Movimentação e não para no primeiro "Saldos por Período".
    resumo_count = sum(1 for ln in lines if re.search(r"resumo\s*-", ln, re.I))
    mov_tx_count = 0
    for idx, ln in enumerate(lines):
        if norm_space(ln).lower() in {"movimentação", "movimentacao"}:
            nxt = " ".join(lines[idx + 1: idx + 12]).lower()
            if "nº documento" in nxt and "movimentos (r$)" in nxt and "saldo (r$)" in nxt:
                mov_tx_count += 1
    if resumo_count > 1 or mov_tx_count > 1:
        return parse_santander_layout1_multiblock_fitz(pdf_path, pdf_password=pdf_password)
    return parse_santander_layout1_from_pdf(pdf_path, pdf_password=pdf_password)

def detect_year(lines):
    sample = " ".join(lines[:400]).lower()
    m = re.search(r"\b(20\d{2})\b", sample)
    return int(m.group(1)) if m else datetime.now().year



def detect_itau_statement_year(lines):
    """Infere o ano de competência de extratos Itaú.

    Prioriza o período do extrato, pois alguns arquivos trazem no cabeçalho
    a data/hora de emissão em ano posterior ao período movimentado.
    """
    sample = " ".join(lines[:160])

    m = re.search(
        r"Extrato\s+de\s+\d{2}/\d{2}/(20\d{2})\s+(?:até|ate)\s+\d{2}/\d{2}/(20\d{2})",
        sample,
        re.I,
    )
    if m:
        return int(m.group(1))

    for ln in lines[:80]:
        m = re.search(r"\b([A-Za-zçÇ]+)/\s*(20\d{2})\b", ln)
        if m:
            return int(m.group(2))

    return detect_year(lines)


def is_itau_layout4_signature(txt: str) -> bool:
    """Reconhece o layout ItaúEmpresas com Data/Lançamento/Ag. Origem/Valor/Saldo.

    A identificação é estrutural para evitar dependência exclusiva da palavra Itaú,
    que em alguns PDFs pode estar apenas no logotipo.
    """
    t = norm_space(txt).lower()
    return (
        "extrato de " in t
        and "ag./origem" in t
        and "valor (r$)" in t
        and "saldo (r$)" in t
        and ("lançamento" in t or "lancamento" in t)
    )




def is_itau_layout5_signature(txt: str) -> bool:
    """Reconhece o layout Itaú BBA com datas no formato 'dd / mmm'.

    Layout observado:
      - período no formato "lançamentos período: dd/mm/aaaa até dd/mm/aaaa";
      - colunas: data | lançamentos | ag/origem | valor (R$) | saldo (R$);
      - as datas das operações aparecem como "03 / fev";
      - a coluna ag/origem e a coluna saldo são desprezadas.
    """
    t = norm_space(txt).lower()
    return (
        ("lançamentos período:" in t or "lancamentos periodo:" in t or "lançamentos periodo:" in t or "lancamentos período:" in t)
        and "ag/origem" in t
        and "valor (r$)" in t
        and "saldo (r$)" in t
    )




# ---------------- Itaú layout7 - PDF-imagem (OCR localizado) ----------------

def _configure_tesseract_if_available():
    """Localiza o executável do Tesseract apenas quando o layout-imagem é usado."""
    if pytesseract is None:
        return False

    current = getattr(pytesseract.pytesseract, "tesseract_cmd", "tesseract")
    if current and (shutil.which(current) or os.path.isfile(current)):
        return True

    candidates = [
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for cand in candidates:
        if cand and os.path.isfile(cand):
            pytesseract.pytesseract.tesseract_cmd = cand
            return True
    return False


def _itau_layout7_ocr_page(page, scale=2.30):
    """OCR de uma página rasterizada, retornando (largura, altura, dataframe TSV)."""
    if pytesseract is None or Image is None or ImageOps is None or not _configure_tesseract_if_available():
        raise RuntimeError(
            "Itaú layout7 em formato de imagem exige OCR. Instale 'pytesseract', 'Pillow' "
            "e o executável Tesseract-OCR; os demais layouts continuam sem essa dependência."
        )

    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    gray = ImageOps.autocontrast(img.convert("L"))
    data = pytesseract.image_to_data(
        gray,
        lang="eng",
        config="--psm 6",
        output_type=pytesseract.Output.DATAFRAME,
    )
    if data is None or data.empty:
        return pix.width, pix.height, pd.DataFrame()

    data = data.copy()
    data["conf"] = pd.to_numeric(data.get("conf"), errors="coerce")
    data = data[data["text"].notna() & data["conf"].ge(20)].copy()
    return pix.width, pix.height, data


def _itau_layout7_ocr_signature(data) -> bool:
    if data is None or data.empty:
        return False
    text = norm_space(" ".join(data["text"].astype(str))).upper()
    # NFKD para tolerar OCR sem acentos e erros usuais de 'LANÇAMENTO'.
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return (
        "EXTRATO DE CONTA CORRENTE" in text
        and ("ITAUEMPRESAS" in text or "ITAU EMPRESAS" in text or "ITAU" in text)
        and "VALOR" in text
        and "SALDO" in text
        and ("LANCAMENTO" in text or "LANGAMENTO" in text)
    )


def _itau_layout7_money_from_ocr(token):
    """Converte valor OCR como 18.106,00-; sinal final negativo = débito."""
    t = norm_space(str(token or "")).replace("−", "-").replace("–", "-")
    t = t.replace(" ", "")
    # Correção conservadora de O->0 somente quando o token já tem dígitos.
    if re.search(r"\d", t):
        t = t.replace("O", "0").replace("o", "0")
    if not re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}-?", t):
        return None
    try:
        return money_to_float(t)
    except Exception:
        return None


def _itau_layout7_group_ocr_lines(data):
    """Agrupa palavras do TSV do Tesseract preservando a geometria das linhas."""
    if data is None or data.empty:
        return []
    out = []
    keys = ["block_num", "par_num", "line_num"]
    for _, g in data.groupby(keys, sort=False):
        g = g.sort_values("left")
        top = float(pd.to_numeric(g["top"], errors="coerce").median())
        out.append((top, g))
    out.sort(key=lambda item: item[0])
    return out


def parse_itau_layout7_image(pdf_path, pdf_password=None):
    """Parser do Itaú layout7 composto por páginas-imagem.

    Estrutura visual observada:
      Data | Lançamento | Valor (R$) | Saldo (R$)

    Regras:
      - cada lançamento possui dd/mm na própria linha;
      - sinal '-' após o valor indica débito; ausência de sinal indica crédito;
      - marcador isolado 'D' após a data é ignorado;
      - códigos da coluna Lançamento (ex.: 445) permanecem agregados à descrição;
      - Documento fica vazio, pois o layout não possui coluna documental própria;
      - linhas que trazem somente valor na coluna Saldo (ex.: SALDO ANTERIOR / S A L D O)
        são descartadas automaticamente por não possuírem valor na coluna Valor;
      - 'SALDO APLIC AUT MAIS' / 'SDO APLIC AUT MAIS AP' são preservados quando o
        próprio PDF os posiciona na coluna Valor, conforme o layout fornecido.
    """
    doc = fitz.open(pdf_path)
    try:
        if doc.needs_pass:
            if not pdf_password or not doc.authenticate(pdf_password):
                raise RuntimeError("PDF Itaú layout7 protegido por senha e não autenticado")

        if len(doc) == 0:
            return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

        w0, h0, first_data = _itau_layout7_ocr_page(doc[0])
        if not _itau_layout7_ocr_signature(first_data):
            return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

        first_text = norm_space(" ".join(first_data["text"].astype(str)))
        ref_match = re.search(r"\b(\d{2}/\d{2}/20\d{2})\b", first_text)
        ref_date = None
        if ref_match:
            try:
                ref_date = datetime.strptime(ref_match.group(1), "%d/%m/%Y")
            except Exception:
                ref_date = None
        if ref_date is None:
            y = detect_year([first_text])
            ref_date = datetime(y, 12, 31) if y else None

        rows = []
        page_cache = {0: (w0, h0, first_data)}
        for pageno, page in enumerate(doc):
            if pageno in page_cache:
                width, height, data = page_cache[pageno]
            else:
                width, height, data = _itau_layout7_ocr_page(page)
            if data is None or data.empty:
                continue

            # Zonas relativas tornam o parser independente da resolução usada no OCR.
            date_xmax = 0.16 * width
            desc_xmin = 0.15 * width
            desc_xmax = 0.47 * width
            value_xmin = 0.47 * width
            value_xmax = 0.64 * width

            for _, line in _itau_layout7_group_ocr_lines(data):
                # Data deve estar na coluna esquerda.
                date_token = None
                left_part = line[pd.to_numeric(line["left"], errors="coerce") < date_xmax]
                for tok in left_part["text"].astype(str):
                    if re.fullmatch(r"\d{2}/\d{2}", norm_space(tok)):
                        date_token = norm_space(tok)
                        break
                if not date_token:
                    continue

                # Somente a coluna VALOR define transação. Números apenas em SALDO são ignorados.
                left_num = pd.to_numeric(line["left"], errors="coerce")
                value_part = line[(left_num >= value_xmin) & (left_num < value_xmax)]
                valor = None
                for tok in value_part["text"].astype(str):
                    parsed = _itau_layout7_money_from_ocr(tok)
                    if parsed is not None:
                        valor = parsed
                if valor is None or valor == 0:
                    continue

                desc_part = line[(left_num >= desc_xmin) & (left_num < desc_xmax)]
                desc_tokens = []
                for tok in desc_part["text"].astype(str):
                    t = norm_space(tok)
                    if not t or t.upper() == "D" or t in {"—", "–", "-", "|", "="}:
                        continue
                    if _itau_layout7_money_from_ocr(t) is not None:
                        continue
                    # Remove pontuação terminal espúria do OCR em palavras, sem tocar
                    # em códigos como 033.1002PLUGAR ou números de agência/conta.
                    if re.fullmatch(r"[A-Za-zÀ-ÿ]+[.,;:]", t):
                        t = t.rstrip(".,;:")
                    desc_tokens.append(t.strip("|"))
                descricao = norm_space(" ".join(desc_tokens)).strip(" -|=")
                if not descricao:
                    continue

                try:
                    day, month = [int(x) for x in date_token.split("/")]
                    if ref_date is None:
                        continue
                    year = ref_date.year - 1 if month > ref_date.month else ref_date.year
                    dt = datetime(year, month, day).strftime("%d/%m/%Y")
                except Exception:
                    continue

                rows.append([dt, descricao, "", valor])

        return standardize(
            pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
            doc_cleaner=lambda x: "",
        )
    finally:
        doc.close()


def is_itau_layout6_signature(txt: str) -> bool:
    """Reconhece o Itaú layout6.

    Layout observado:
      - período no formato "Lançamentos do período: dd/mm/aaaa até dd/mm/aaaa";
      - colunas: Data | Lançamentos | Razão Social | CNPJ/CPF | Valor (R$) | Saldo (R$);
      - todas as transações trazem data completa na própria linha;
      - Razão Social e CNPJ/CPF, quando presentes, são agregados à descrição;
      - campo Documento permanece vazio.
    """
    t = norm_space(txt).lower()
    return (
        ("lançamentos do período:" in t or "lancamentos do periodo:" in t or "lançamentos do periodo:" in t or "lancamentos do período:" in t)
        and ("razão social" in t or "razao social" in t)
        and "cnpj/cpf" in t
        and "valor (r$)" in t
        and "saldo (r$)" in t
        and "saldo total" in t
        and "limite da conta" in t
    )


def parse_itau_layout6(lines):
    """Parser para Itaú layout6.

    O layout é semelhante ao layout3, mas foi isolado para preservar a lógica
    anterior e lidar expressamente com os blocos finais de saldo/limite da conta.
    O parser usa a extração linear do PyMuPDF: Data, descrição/razão social/CNPJ
    e Valor aparecem em linhas sucessivas. A coluna Saldo e linhas de saldo são
    descartadas por descrição.
    """
    full_date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []
    i = 0
    N = len(lines)

    header_noise = {
        "Data", "Lançamentos", "Lancamentos", "Razão Social", "Razao Social",
        "CNPJ/CPF", "Valor (R$)", "Saldo (R$)", "Saldo total", "Limite da conta",
        "Utilizado", "Disponível", "Disponivel"
    }

    stop_prefixes = (
        "saldo da conta corrente", "aviso:", "aviso ", "atualizado em ",
        "em caso de dúvidas", "em caso de duvidas", "reclamações, informações",
        "reclamacoes, informacoes", "deficiente auditivo"
    )

    while i < N:
        ln = norm_space(lines[i])
        low = ln.lower()

        if low.startswith(stop_prefixes):
            break

        if not full_date_re.match(ln):
            i += 1
            continue

        dt = ln
        j = i + 1
        desc_parts = []

        while j < N:
            cur = norm_space(lines[j])
            cur_low = cur.lower()
            if full_date_re.match(cur):
                break
            if cur_low.startswith(stop_prefixes):
                break
            if cur in header_noise:
                j += 1
                continue
            if money_re.match(cur):
                break
            # Evita carregar ruídos de cabeçalho/rodapé quando houver quebra visual.
            if cur_low in {"agência", "agencia", "conta", "conta corrente"}:
                j += 1
                continue
            desc_parts.append(cur)
            j += 1

        if j >= N or not money_re.match(norm_space(lines[j])):
            i += 1
            continue

        desc = norm_space(" ".join(desc_parts))
        val = money_to_float(norm_space(lines[j]))

        if desc and not is_balance_or_summary_line(desc) and not is_itau_balance_description(desc):
            rows.append([dt, desc, "", val])

        i = j + 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


MESES_ITAU_ABREV = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def parse_itau_layout5_pdf(pdf_path, lines=None, pdf_password=None):
    """Parser para Itaú layout5.

    Usa coordenadas do PyMuPDF porque a extração textual linear desse layout
    pode separar/embaralhar as colunas. A rotina lê apenas as colunas:
    data, lançamentos e valor; ignora ag/origem e saldo.

    Robustez multibloco: quando o PDF trouxer vários extratos mensais no mesmo
    arquivo, a rotina atualiza o ano de referência a cada nova linha do tipo
    "lançamentos período: dd/mm/aaaa até dd/mm/aaaa". Isso evita que um bloco
    posterior, inclusive em outro ano, herde indevidamente o ano do primeiro
    período encontrado.
    """
    lines = lines or []

    date_txt_re = re.compile(r"^\d{1,2}\s*/\s*[A-Za-zçÇ]{3}$", re.I)
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    periodo_re = re.compile(
        r"lan[çc]amentos\s+per[ií]odo:\s*"
        r"(?P<d1>\d{2})/(?P<m1>\d{2})/(?P<y1>20\d{2})\s+(?:at[eé])\s+"
        r"(?P<d2>\d{2})/(?P<m2>\d{2})/(?P<y2>20\d{2})",
        re.I,
    )

    fallback_year = detect_year(lines)
    current_period = None

    def group_words_by_line(words, tol=2.0):
        words = sorted(words, key=lambda w: (w[1], w[0]))
        groups = []
        current = []
        current_top = None
        for w in words:
            top = float(w[1])
            if current_top is None or abs(top - current_top) <= tol:
                current.append(w)
                current_top = top if current_top is None else (current_top + top) / 2.0
            else:
                groups.append(sorted(current, key=lambda z: z[0]))
                current = [w]
                current_top = top
        if current:
            groups.append(sorted(current, key=lambda z: z[0]))
        return groups

    def update_period_from_text(text_line: str):
        nonlocal current_period
        t = norm_space(text_line).lower()
        m = periodo_re.search(t)
        if not m:
            return
        try:
            current_period = {
                "start_month": int(m.group("m1")),
                "start_year": int(m.group("y1")),
                "end_month": int(m.group("m2")),
                "end_year": int(m.group("y2")),
            }
        except Exception:
            current_period = None

    def infer_year_for_month(month: int) -> int:
        if current_period:
            sy = current_period["start_year"]
            ey = current_period["end_year"]
            sm = current_period["start_month"]
            em = current_period["end_month"]
            if sy == ey:
                return sy
            # Bloco que cruza o ano, por exemplo 01/12/2021 até 31/01/2022.
            if month >= sm:
                return sy
            if month <= em:
                return ey
            return sy
        return fallback_year

    rows = []
    doc, _password_used = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            words = page.get_text("words") or []
            for group in group_words_by_line(words):
                line_txt = norm_space(" ".join(w[4] for w in group))
                if line_txt:
                    update_period_from_text(line_txt)

                date_txt = norm_space(" ".join(w[4] for w in group if w[0] < 70))
                if not date_txt_re.match(date_txt):
                    continue

                m_dt = re.match(r"^(\d{1,2})\s*/\s*([A-Za-zçÇ]{3})$", date_txt, re.I)
                if not m_dt:
                    continue
                mes_txt = m_dt.group(2).lower().replace("ç", "c")
                mes = MESES_ITAU_ABREV.get(mes_txt)
                if not mes:
                    continue

                desc = norm_space(" ".join(w[4] for w in group if 70 <= w[0] < 300))
                if not desc:
                    continue

                # Valor da transação fica na coluna valor (R$). Valores de saldo
                # aparecem mais à direita e são ignorados.
                value_tokens = [
                    w[4] for w in group
                    if 360 <= w[0] < 470 and money_re.fullmatch(w[4])
                ]
                if not value_tokens:
                    continue

                if is_balance_or_summary_line(desc) or is_itau_balance_description(desc):
                    continue

                try:
                    data = datetime(infer_year_for_month(mes), mes, int(m_dt.group(1))).strftime("%d/%m/%Y")
                    valor = money_to_float(value_tokens[-1])
                except Exception:
                    continue

                rows.append([data, desc, "", valor])
    finally:
        doc.close()

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))

def is_itau_layout3_signature(txt: str) -> bool:
    """Reconhece o layout Itaú com colunas Data/Lançamentos/Razão Social/CNPJ/CPF/Valor/Saldo.

    A identificação é estrutural porque alguns PDFs trazem o logotipo Itaú como imagem,
    sem a palavra ITAU/ITAÚ no texto extraído pelo PyMuPDF.
    """
    t = norm_space(txt).lower()
    return (
        "lançamentos do período:" in t or "lancamentos do periodo:" in t
    ) and (
        "razão social" in t or "razao social" in t
    ) and (
        "cnpj/cpf" in t
    ) and (
        "valor (r$)" in t
    ) and (
        "saldo (r$)" in t
    )


def is_itau_balance_description(desc: str) -> bool:
    low = norm_space(desc).lower()
    return any(x in low for x in [
        "saldo anterior",
        "saldo inicial",
        "saldo final",
        "saldo parcial",
        "saldo total",
        "saldo total disponível dia",
        "saldo total disponivel dia",
        "saldo movimentação conta",
        "saldo movimentacao conta",
        "saldo aplic",
        "saldo aplicação",
        "saldo aplicacao",
        "saldo a liberar",
        "saldo em conta corrente",
    ])




def parse_itau_layout2(lines):
    year = None
    for ln in lines[:40]:
        m = re.search(r"\b([A-Za-zçÇ]+)/(\d{4})\b", ln)
        if m:
            try:
                year = int(m.group(2))
                break
            except Exception:
                pass
    if year is None:
        year = detect_year(lines)

    date_re = re.compile(r"^\d{2}/\d{2}$")
    money_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}(?:\s*-\s*)?$")
    rows = []
    i = 0
    N = len(lines)

    skip_desc_starts = (
        "saldo inicial", "saldo anterior", "saldo final", "saldo parcial",
        "saldo aplic", "saldo em "
    )

    while i < N:
        ln = lines[i]
        if not date_re.match(ln):
            i += 1
            continue

        dt = datetime.strptime(f"{ln}/{year}", "%d/%m/%Y").strftime("%d/%m/%Y")
        j = i + 1
        if j >= N:
            break

        desc = norm_space(lines[j])
        if not desc or desc.lower().startswith(skip_desc_starts):
            i = j + 1
            continue

        j += 1
        # coluna Orig opcional
        if j < N and re.fullmatch(r"\d{3,6}", lines[j]):
            j += 1

        if j >= N or not money_re.match(lines[j]):
            i += 1
            continue

        val = money_to_float(lines[j])
        rows.append([dt, desc, "", val])

        # consumir saldo subsequente, se houver
        j += 1
        if j < N and money_re.match(lines[j]):
            j += 1

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))



def parse_itau_layout3(lines):
    full_date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []
    i = 0
    N = len(lines)

    header_noise = {
        "Data", "Lançamentos", "Lancamentos", "Razão Social", "Razao Social",
        "CNPJ/CPF", "Valor (R$)", "Saldo (R$)"
    }

    while i < N:
        ln = lines[i]
        if not full_date_re.match(ln):
            i += 1
            continue

        dt = ln
        j = i + 1
        desc_parts = []

        while j < N:
            cur = lines[j]
            if full_date_re.match(cur):
                break
            if cur in header_noise:
                j += 1
                continue
            if money_re.match(cur):
                break
            desc_parts.append(cur)
            j += 1

        if j >= N or not money_re.match(lines[j]):
            i += 1
            continue

        desc = norm_space(" ".join(desc_parts))
        val = money_to_float(lines[j])

        if desc and not is_balance_or_summary_line(desc) and not is_itau_balance_description(desc):
            rows.append([dt, desc, "", val])

        i = j + 1

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))


def parse_itau_layout4(lines):
    """Parser para Itaú layout4.

    Layout observado:
      - cabeçalho ItaúEmpresas;
      - período no formato "Extrato de dd/mm/aaaa até dd/mm/aaaa";
      - colunas: Data | Lançamento | Ag./Origem | Valor (R$) | Saldo (R$);
      - datas das operações no formato dd/mm;
      - Ag./Origem e Saldo (R$) são desprezados;
      - linhas de saldo não devem virar transação.
    """
    year = detect_itau_statement_year(lines)
    date_re = re.compile(r"^\d{2}/\d{2}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    rows = []
    i = 0
    N = len(lines)

    header_noise = {
        "ItaúEmpresas", "ItauEmpresas", "Data", "Lançamento", "Lancamento",
        "Ag./Origem", "Valor (R$)", "Saldo (R$)"
    }

    while i < N:
        ln = norm_space(lines[i])
        if not date_re.match(ln):
            i += 1
            continue

        try:
            dt = datetime.strptime(f"{ln}/{year}", "%d/%m/%Y").strftime("%d/%m/%Y")
        except Exception:
            i += 1
            continue

        j = i + 1
        desc_parts = []
        while j < N:
            cur = norm_space(lines[j])
            if date_re.match(cur):
                break
            if cur in header_noise or re.fullmatch(r"\d+", cur):
                j += 1
                continue
            if money_re.match(cur):
                break
            desc_parts.append(cur)
            j += 1

        if j >= N or not money_re.match(norm_space(lines[j])):
            i += 1
            continue

        desc = norm_space(" ".join(desc_parts))

        if desc and not is_balance_or_summary_line(desc) and not is_itau_balance_description(desc):
            val = money_to_float(norm_space(lines[j]))
            rows.append([dt, desc, "", val])

        # Consome o valor da transação e, se existir, o saldo imediatamente subsequente.
        j += 1
        if j < N and money_re.match(norm_space(lines[j])):
            j += 1

        i = j

    return standardize(pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]))

def parse_itau(lines, pdf_path=None, pdf_password=None):
    sample = " ".join(lines[:180]).lower()
    if ("data lançamentos ag/origem valor (r$) saldo (r$)" in sample or
        "data lancamentos ag/origem valor (r$) saldo (r$)" in sample or
        is_itau_layout5_signature(sample)) and pdf_path:
        return parse_itau_layout5_pdf(pdf_path, lines=lines, pdf_password=pdf_password)
    if is_itau_layout6_signature(sample):
        return parse_itau_layout6(lines)
    if ("data lançamentos razão social cnpj/cpf valor (r$) saldo (r$)" in sample or
        "data lancamentos razao social cnpj/cpf valor (r$) saldo (r$)" in sample or
        is_itau_layout3_signature(sample)):
        return parse_itau_layout3(lines)
    if ("data lançamento ag./origem valor (r$) saldo (r$)" in sample or
        "data lancamento ag./origem valor (r$) saldo (r$)" in sample or
        is_itau_layout4_signature(sample)):
        return parse_itau_layout4(lines)
    if "data histórico de lançamentos orig valor (r$) saldo (r$)".lower() in sample or "data historico de lancamentos orig valor (r$) saldo (r$)" in sample:
        return parse_itau_layout2(lines)

    year = detect_year(lines)
    rows = []
    current = None
    in_mov = False
    i = 0
    N = len(lines)

    header_noise = (
        "data", "descrição", "descricao", "entradas r$", "saídas r$", "saidas r$", "saldo r$",
        "(créditos)", "(creditos)", "(debitos)", "(débitos)", "a = agendamento", "b = ações movimentadas",
        "c = crédito a compensar", "c = credito a compensar", "d = débito a compensar", "d = debito a compensar",
        "g = aplicação programada", "g = aplicacao programada", "p = poupança automática", "p = poupanca automatica",
        "pela bolsa de valores", "para demais siglas, consulte as notas", "explicativas no final do extrato"
    )

    mes_ano_re = re.compile(
        r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(20\d{2})\b",
        re.I,
    )

    def update_year_from_month_header(text_line: str):
        nonlocal year
        m = mes_ano_re.search(norm_space(text_line).lower())
        if m:
            try:
                year = int(m.group(2))
            except Exception:
                pass

    def is_itau_layout1_mov_header(text_line: str) -> bool:
        t = norm_space(text_line).lower()
        return bool(re.fullmatch(r"conta corrente\s*\|\s*movimenta[cç][aã]o", t, flags=re.I))

    while i < N:
        ln = norm_space(lines[i])
        low = ln.lower()

        # Em arquivos consolidados, cada mês pode reiniciar com novo cabeçalho
        # "jan 2021", "fev 2021" etc.; atualiza o ano de referência sem
        # depender somente do primeiro mês do PDF.
        update_year_from_month_header(ln)

        if is_itau_layout1_mov_header(ln):
            in_mov = True
            current = None
            i += 1
            continue

        if not in_mov:
            i += 1
            continue

        # Fim apenas do bloco de movimentação corrente. Não encerra o parser,
        # pois PDFs consolidados podem trazer outro bloco Conta Corrente |
        # Movimentação alguns páginas depois.
        if re.search(r"^Conta\s+Corrente\s*\|\s*Aplica", ln, re.I):
            in_mov = False
            current = None
            i += 1
            continue
        if re.search(r"^Conta\s+Corrente\s*\|\s*Cheque", ln, re.I):
            in_mov = False
            current = None
            i += 1
            continue

        # datas
        if re.fullmatch(r"\d{2}/\d{2}", ln):
            try:
                current = datetime.strptime(f"{ln}/{year}", "%d/%m/%Y")
            except Exception:
                current = None
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
            "totalizador de aplicações automáticas", "totalizador de aplicacoes automaticas",
            "os valores referentes ao totalizador", "este material está disponível", "este material esta disponivel"
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
        if i + 1 < N and re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}-?", norm_space(lines[i + 1])):
            val = money_to_float(norm_space(lines[i + 1]))
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

def parse_unicred(pdf_path, pdf_password=None):
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
    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
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


# ---------------- Unicred layout 2 ----------------

def is_unicred_layout2_signature(txt: str, name: str = "") -> bool:
    """Assinatura estrutural do Unicred_layout2.

    Layout observado:
      Data | Lançamentos | Valor (R$) | Saldo (R$)

    Evita depender do nome do arquivo. Exige elementos típicos do extrato
    Unicred/Cooperativa para não confundir com layouts Itaú/Bradesco.
    """
    sample = norm_space(str(txt or "")).lower()
    has_columns = (
        ("lançamentos" in sample or "lancamentos" in sample)
        and "valor (r$)" in sample
        and "saldo (r$)" in sample
        and "data" in sample
    )
    has_unicred_context = (
        "coop:" in sample and "conta:" in sample and "período de" in sample
    ) or (
        "coop:" in sample and "solicitado por" in sample
    ) or name.lower().startswith("unicred")
    # Distingue dos layouts já mapeados que também podem possuir Data/Lançamentos.
    not_itau = "itau" not in sample and "itaú" not in sample and "ag/origem" not in sample
    not_bradesco = "bradesco" not in sample and "dcto." not in sample and "dcto " not in sample
    return has_columns and has_unicred_context and not_itau and not_bradesco


def parse_unicred_layout2(pdf_path, pdf_password=None):
    """Parser para Unicred_layout2.

    Estrutura:
      Data | Lançamentos | Valor (R$) | Saldo (R$)

    A extração por pdfplumber pode alternar a ordem textual das colunas:
    em vários registros a descrição vem antes da linha de data/valor, e a
    continuação do campo Doc. pode vir na linha seguinte. Por isso o parser
    monta cada registro a partir de:
      descrição pendente + trecho da linha de data antes do valor + continuação
      necessária para fechar o parêntese do Doc.
    """
    if pdfplumber is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    date_prefix_re = re.compile(r"^(\d{2}/\d{2}/\d{4})\b")
    money_re = re.compile(r"(?P<neg>-)?\s*R\$\s*(?P<num>\d{1,3}(?:\.\d{3})*,\d{2})")
    doc_re = re.compile(r"\(\s*Doc\.?\s*:\s*(.*?)\s*\)", re.I)

    def parse_money_match(m):
        val = float(m.group("num").replace(".", "").replace(",", "."))
        return -val if m.group("neg") else val

    def is_noise_line(line: str) -> bool:
        low = norm_space(line).lower()
        if not low:
            return True
        if low in {"data lançamentos valor (r$) saldo (r$)", "data lancamentos valor (r$) saldo (r$)", "data", "lançamentos", "lancamentos", "valor (r$)", "saldo (r$)"}:
            return True
        return any(x in low for x in [
            "central de relacionamento", "capitais e regiões metropolitanas", "capitais e regioes metropolitanas",
            "0800 200 7302", "ouvidoria", "pág.", "pag.", "/8", "extrato",
            "período de ", "periodo de ", "coop:", "solicitado por", "saldo em ",
            "saldo atual", "total disponível", "total disponivel", "limite de cheque especial",
            "saldo bloqueado", "tarifas pendentes", "saldo bloqueado judicialmente", "saldo bloqueado cautelar"
        ])

    def should_stop(line: str) -> bool:
        low = norm_space(line).lower()
        return low.startswith("saldo no final do período") or low.startswith("saldo no final do periodo") or low.startswith("lançamentos futuros") or low.startswith("lancamentos futuros")

    def needs_continuation(desc_parts):
        txt = norm_space(" ".join(desc_parts))
        # Quase todas as quebras relevantes do layout estão dentro de ( Doc.: ... ).
        return ("( Doc" in txt or "(Doc" in txt or "( doc" in txt.lower()) and txt.count("(") > txt.count(")")

    def make_record(data: str, desc_parts, valor: float):
        body_desc = norm_space(" ".join([x for x in desc_parts if norm_space(x)]))
        if not body_desc or is_balance_or_summary_line(body_desc):
            return None

        # Refinamento v58.1:
        # O conteúdo entre parênteses "( Doc.: ... )" pertence à coluna Lançamentos.
        # Só deve ser deslocado para a coluna Documento quando for um único número,
        # por exemplo: "( Doc.: 197793 )" ou "( Doc.: 0 )".
        # Quando houver texto, barra, beneficiário, convênio etc., preserva-se integralmente
        # na Descrição, pois é informação material da operação.
        doc = ""
        desc = body_desc
        mdoc = doc_re.search(body_desc)
        if mdoc:
            doc_candidate = norm_space(mdoc.group(1))
            if re.fullmatch(r"\d+", doc_candidate):
                doc = doc_candidate[:80]
                desc = norm_space(doc_re.sub("", body_desc))

        if not desc or is_balance_or_summary_line(desc):
            return None
        return [data, desc, doc, valor]

    raw_lines = []
    with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines.extend([norm_space(x) for x in text.splitlines() if norm_space(x)])

    rows = []
    pending_desc = []
    in_table = False
    i = 0
    N = len(raw_lines)
    while i < N:
        line = raw_lines[i]
        low_line = norm_space(line).lower()

        if should_stop(line):
            break
        if ("data lançamentos" in low_line or "data lancamentos" in low_line) and "valor (r$)" in low_line and "saldo (r$)" in low_line:
            in_table = True
            pending_desc = []
            i += 1
            continue
        if not in_table:
            i += 1
            continue
        if is_noise_line(line):
            pending_desc = []
            i += 1
            continue

        dm = date_prefix_re.match(line)
        if not dm:
            # Linha de descrição, geralmente antes da linha que contém data/valor.
            if not money_re.search(line):
                pending_desc.append(line)
            i += 1
            continue

        # Linhas de data sem dois valores monetários são cabeçalho/ruído, não transação.
        matches = list(money_re.finditer(line))
        if len(matches) < 2:
            i += 1
            continue

        data = dm.group(1)
        tail = norm_space(line[dm.end():])
        valor = parse_money_match(matches[-2])
        # O trecho entre a data e o valor da transação também pode conter descrição.
        inline_desc = norm_space(line[dm.end():matches[-2].start()])
        desc_parts = pending_desc + ([inline_desc] if inline_desc else [])
        pending_desc = []

        j = i + 1
        while j < N and needs_continuation(desc_parts):
            nxt = raw_lines[j]
            if should_stop(nxt) or is_noise_line(nxt) or date_prefix_re.match(nxt):
                break
            # Continuação do campo Doc.; não deve capturar início de outra transação.
            if money_re.search(nxt):
                break
            desc_parts.append(nxt)
            j += 1

        rec = make_record(data, desc_parts, valor)
        if rec:
            rows.append(rec)
        i = j

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )




# ---------------- Bradesco layout 3 (Extrato Mensal / Por Período) ----------------

def is_bradesco_layout3_signature(txt: str, name: str = "") -> bool:
    sample = norm_space(str(txt or "")).lower()
    has_columns = (
        ("data lançamento dcto" in sample or "data lancamento dcto" in sample)
        and ("crédito (r$)" in sample or "credito (r$)" in sample)
        and ("débito (r$)" in sample or "debito (r$)" in sample)
        and "saldo (r$)" in sample
    )
    has_header = (
        "extrato mensal / por período" in sample or
        "extrato mensal / por periodo" in sample or
        "extrato de: ag:" in sample or
        "total disponível (r$)" in sample or
        "total disponivel (r$)" in sample
    )
    # Distingue do Bradesco_layout2, cujo cabeçalho usa HISTÓRICO/DOCTO e
    # contém o bloco "MOVIMENTAÇÃO DO PERÍODO - CONTA CORRENTE".
    not_layout2 = "movimentação do período - conta corrente" not in sample and "movimentacao do periodo - conta corrente" not in sample
    return has_columns and has_header and not_layout2


def parse_bradesco_layout3(pdf_path, pdf_password=None):
    """Parser para Bradesco_layout3.

    Layout observado:
      Data | Lançamento | Dcto. | Crédito (R$) | Débito (R$) | Saldo (R$)

    A extração usa a ordem textual do PDF, porque esse layout textual preserva
    adequadamente a sequência: descrição(s) -> documento -> valor -> saldo.

    Refinamento v57.9:
      - após o bloco principal, o parser também analisa "Últimos Lançamentos";
      - considera somente transações do(s) mesmo(s) mês(es) do bloco principal;
      - elimina duplicidades contra o bloco principal, priorizando Data+Dcto+Valor;
      - continua ignorando o bloco "Saldos Invest Fácil / Plus" e notas manuais.
    """
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")

    def bradesco_amount_to_float(tok: str) -> float:
        t = norm_space(str(tok or "")).replace("−", "-")
        neg = t.startswith("-") or t.endswith("-")
        t = t.strip("-")
        val = float(t.replace(".", "").replace(",", "."))
        return -val if neg else val

    def is_header_line(line: str) -> bool:
        low = norm_space(line).lower()
        if not low:
            return True
        return low in {
            "data", "lançamento", "lancamento", "dcto.", "dcto", "crédito (r$)", "credito (r$)",
            "débito (r$)", "debito (r$)", "saldo (r$)", "histórico", "historico", "valor (r$)",
        } or low.startswith((
            "extrato de:", "agência | conta", "agencia | conta", "total disponível", "total disponivel", "total (r$)",
            "extrato mensal / por período", "extrato mensal / por periodo", "antonio jose", "nome do usuário", "nome do usuario",
            "data da operação", "data da operacao", "os dados acima", "saldos invest", "data histórico valor", "data historico valor"
        ))

    def is_balance_desc(desc: str) -> bool:
        up = norm_space(desc).upper()
        return up.startswith("SALDO") or up.startswith("TOTAL")

    def month_key(data: str):
        try:
            dt = datetime.strptime(data, "%d/%m/%Y")
            return (dt.year, dt.month)
        except Exception:
            return None

    def norm_key_text(s: str) -> str:
        return normalize_text_for_dedupe(s)

    rows_main = []
    rows_ultimos_raw = []
    current_date = None
    pending_desc = []
    mode = "seek_main"  # seek_main | main | after_main | seek_ultimos_header | ultimos | done

    def clear_if_balance_pending():
        nonlocal pending_desc
        if pending_desc and is_balance_desc(" ".join(pending_desc)):
            pending_desc = []

    def append_row(target_rows, doc_txt: str, valor_txt: str):
        nonlocal pending_desc
        if not current_date or not pending_desc:
            pending_desc = []
            return
        desc = norm_space(" ".join(pending_desc))
        pending_desc = []
        if not desc or is_balance_desc(desc):
            return
        try:
            valor = bradesco_amount_to_float(valor_txt)
        except Exception:
            return
        target_rows.append([current_date, desc, norm_space(doc_txt), valor])

    def duplicate_against_main(row) -> bool:
        data, desc, doc_txt, valor = row
        valor_r = round(float(valor), 2)
        doc_key = norm_space(doc_txt)
        desc_key = norm_key_text(desc)
        for m_data, m_desc, m_doc, m_val in rows_main:
            if data != m_data:
                continue
            if round(float(m_val), 2) != valor_r:
                continue
            if doc_key and norm_space(m_doc) == doc_key:
                return True
            if not doc_key and norm_key_text(m_desc) == desc_key:
                return True
        return False

    doc, _ = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            raw_lines = [norm_space(x) for x in (page.get_text("text") or "").splitlines()]
            lines = [x for x in raw_lines if x]
            i = 0
            while i < len(lines):
                line = lines[i]
                low = norm_space(line).lower()

                # As notas manuais do arquivo amostral aparecem após o rodapé
                # "Folha X/Y". No bloco principal, interrompe apenas a página,
                # preservando descrição pendente que continue na página seguinte.
                if re.match(r"^folha\s+\d+\s*/\s*\d+", low):
                    break

                if mode == "seek_main":
                    if low == "saldo (r$)":
                        mode = "main"
                    i += 1
                    continue

                if mode == "after_main":
                    pending_desc = []
                    if low.startswith("últimos lançamentos") or low.startswith("ultimos lancamentos"):
                        mode = "seek_ultimos_header"
                    i += 1
                    continue

                if mode == "seek_ultimos_header":
                    if low == "saldo (r$)":
                        current_date = None
                        pending_desc = []
                        mode = "ultimos"
                    elif low.startswith("saldos invest"):
                        mode = "done"
                    i += 1
                    continue

                if mode == "done":
                    i += 1
                    continue

                # Aqui restam os modos transacionais: main ou ultimos.
                target = rows_main if mode == "main" else rows_ultimos_raw

                # No bloco principal, "Total" encerra o período principal, mas não
                # o arquivo inteiro: ainda pode haver "Últimos Lançamentos" úteis.
                # No bloco de últimos lançamentos, "Total" encerra esse bloco.
                if low == "total":
                    pending_desc = []
                    mode = "after_main" if mode == "main" else "done"
                    i += 1
                    continue

                # Segurança: se o bloco de saldos começar, nada mais é transação.
                if low.startswith("saldos invest"):
                    pending_desc = []
                    mode = "done"
                    i += 1
                    continue

                if is_header_line(line):
                    i += 1
                    continue

                # Data isolada: inaugura novo grupo de lançamentos.
                if date_re.fullmatch(line):
                    clear_if_balance_pending()
                    current_date = line
                    pending_desc = []
                    i += 1
                    continue

                # Caso defensivo para linhas com data + texto na mesma linha.
                m_inline_date = re.match(r"^(\d{2}/\d{2}/\d{4})\s+(.+)$", line)
                if m_inline_date:
                    clear_if_balance_pending()
                    current_date = m_inline_date.group(1)
                    rest = norm_space(m_inline_date.group(2))
                    pending_desc = []
                    if rest and not is_balance_desc(rest):
                        pending_desc.append(rest)
                    i += 1
                    continue

                # SALDO ANTERIOR / SALDO INVEST etc. não são transação.
                if is_balance_desc(line):
                    pending_desc = []
                    i += 1
                    continue

                # Quando há descrição pendente e encontramos documento + valor + saldo,
                # fechamos o lançamento. O saldo é consumido e descartado.
                if pending_desc and i + 2 < len(lines) and money_re.fullmatch(lines[i + 1]) and money_re.fullmatch(lines[i + 2]):
                    append_row(target, line, lines[i + 1])
                    i += 3
                    continue

                # Restos numéricos sem descrição pendente normalmente são saldos ou totais.
                if money_re.fullmatch(line):
                    i += 1
                    continue

                # Linha textual da coluna Lançamento, incluindo complementos como REMET:/DEST:/INTERNET.
                pending_desc.append(line)
                i += 1
    finally:
        doc.close()

    main_months = {month_key(r[0]) for r in rows_main if month_key(r[0])}
    rows_ultimos = []
    for row in rows_ultimos_raw:
        if month_key(row[0]) not in main_months:
            continue
        if duplicate_against_main(row):
            continue
        rows_ultimos.append(row)

    rows = rows_main + rows_ultimos

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )

# ---------------- Bradesco layout 2 (Movimentação do Período - Conta Corrente) ----------------

def is_bradesco_layout2_signature(txt: str, name: str = "") -> bool:
    sample = norm_space(str(txt or "")).lower()
    # Não depender do nome do arquivo: alguns downloads do Bradesco vêm como
    # EXTRATO_*.pdf ou outro nome genérico. A assinatura estrutural do layout
    # é suficiente e mais segura para identificação.
    has_mov_periodo = (
        "movimentação do período - conta corrente" in sample or
        "movimentacao do periodo - conta corrente" in sample
    )
    has_columns = (
        "data histórico docto" in sample or
        "data historico docto" in sample or
        ("docto" in sample and ("crédito" in sample or "credito" in sample) and ("débito" in sample or "debito" in sample))
    )
    return has_mov_periodo and has_columns


def parse_bradesco_layout2(pdf_path, pdf_password=None):
    """Parser para extrato Bradesco textual.

    Layout observado:
      MOVIMENTAÇÃO DO PERÍODO - CONTA CORRENTE
      DATA | HISTÓRICO | DOCTO | CRÉDITO | DÉBITO | SALDO

    A extração usa coordenadas das palavras para separar Crédito, Débito e Saldo.
    O campo Saldo e linhas de saldo/totalização são ignorados. A data é mantida
    para lançamentos subsequentes até nova data aparecer. Linhas complementares
    do histórico são anexadas ao mesmo lançamento até surgir novo registro com
    documento/valor.
    """

    date_re = re.compile(r"^\d{2}/\d{2}/\d{2}$")
    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}-?$")

    def bradesco_amount_to_float(tok: str) -> float:
        t = norm_space(str(tok or "")).replace("−", "-")
        neg = t.endswith("-") or t.startswith("-")
        t = t.strip("-")
        val = float(t.replace(".", "").replace(",", "."))
        return -val if neg else val

    def group_words_by_row(words, tol=2.5):
        words = sorted(words, key=lambda w: (float(w[1]), float(w[0])))
        rows = []
        cur = []
        cur_y = None
        for w in words:
            y = (float(w[1]) + float(w[3])) / 2.0
            if cur_y is None or abs(y - cur_y) <= tol:
                cur.append(w)
                cur_y = y if cur_y is None else (cur_y + y) / 2.0
            else:
                rows.append(sorted(cur, key=lambda z: float(z[0])))
                cur = [w]
                cur_y = y
        if cur:
            rows.append(sorted(cur, key=lambda z: float(z[0])))
        return rows

    def is_header_or_footer(text: str) -> bool:
        up = norm_space(text).upper()
        if not up:
            return True
        if up.startswith("DATA HIST") or up in {"(CRÉDITOS) (DÉBITOS)", "(CREDITOS) (DEBITOS)"}:
            return True
        if (
            "EXTRATO PARA SIMPLES" in up or
            up.startswith("SEGUNDA,") or
            up.startswith("AGÊNCIA:") or up.startswith("AGENCIA:")
        ):
            return True
        return False

    def is_balance_or_total(hist: str) -> bool:
        up = norm_space(hist).upper()
        return (
            up.startswith("SALDO") or
            up.startswith("TOTAL DA MOVIMENTAÇÃO") or
            up.startswith("TOTAL DA MOVIMENTACAO")
        )

    def make_date_yyyy(date_txt: str) -> str:
        dd, mm, yy = date_txt.split("/")
        yyyy = "20" + yy if int(yy) < 70 else "19" + yy
        return f"{dd}/{mm}/{yyyy}"

    rows = []
    current = None
    current_date = None
    in_period = False

    def commit_current():
        nonlocal current
        if current and current.get("Valor") is not None and current.get("Descrição"):
            rows.append([
                current["Data"],
                norm_space(" ".join(current["Descrição"])),
                norm_space(current.get("Documento", "")),
                current["Valor"],
            ])
        current = None

    doc, _ = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            words = page.get_text("words") or []
            for wr in group_words_by_row(words):
                row_text = norm_space(" ".join(str(w[4]) for w in wr))
                up = row_text.upper()

                if "MOVIMENTAÇÃO DO PERÍODO" in up or "MOVIMENTACAO DO PERIODO" in up:
                    in_period = True
                    continue

                if not in_period:
                    continue

                if up.startswith("TOTAL DA MOVIMENTAÇÃO") or up.startswith("TOTAL DA MOVIMENTACAO"):
                    commit_current()
                    in_period = False
                    break

                if is_header_or_footer(row_text):
                    continue

                cols = {"date": [], "hist": [], "doc": [], "cred": [], "deb": [], "saldo": []}
                for w in wr:
                    x0, x1 = float(w[0]), float(w[2])
                    x = (x0 + x1) / 2.0
                    tok = str(w[4])
                    if x < 60:
                        cols["date"].append(tok)
                    elif x < 220:
                        cols["hist"].append(tok)
                    elif x < 285:
                        cols["doc"].append(tok)
                    elif x < 380:
                        cols["cred"].append(tok)
                    elif x < 480:
                        cols["deb"].append(tok)
                    else:
                        cols["saldo"].append(tok)

                date_txt = norm_space(" ".join(cols["date"]))
                hist = norm_space(" ".join(cols["hist"]))
                doc_txt = norm_space(" ".join(cols["doc"]))
                cred_txt = norm_space(" ".join(cols["cred"]))
                deb_txt = norm_space(" ".join(cols["deb"]))

                if date_re.fullmatch(date_txt):
                    current_date = make_date_yyyy(date_txt)

                if not current_date:
                    continue

                valor = None
                if money_re.fullmatch(cred_txt):
                    valor = abs(bradesco_amount_to_float(cred_txt))
                if money_re.fullmatch(deb_txt):
                    valor = -abs(bradesco_amount_to_float(deb_txt))

                if hist and is_balance_or_total(hist):
                    continue

                if valor is not None:
                    if current:
                        commit_current()
                    current = {
                        "Data": current_date,
                        "Descrição": [hist] if hist else [],
                        "Documento": doc_txt,
                        "Valor": valor,
                    }
                elif hist:
                    if current:
                        current["Descrição"].append(hist)
                    else:
                        current = {
                            "Data": current_date,
                            "Descrição": [hist],
                            "Documento": doc_txt,
                            "Valor": None,
                        }

            if not in_period:
                # Sai quando o totalizador final do bloco de movimentação é encontrado.
                break

        commit_current()
    finally:
        doc.close()

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )

# ---------------- Bradesco layout 1 (duas tabelas por página) ----------------

def is_bradesco_layout1_signature(txt: str, name: str = "") -> bool:
    sample = norm_space(str(txt or "")).lower()
    nm = str(name or "").lower()
    return (
        ("extrato para" in sample and ("simples confer" in sample or "simples conferencia" in sample or "simples conferência" in sample)
         and ("débito/crédito/saldo" in sample or "debito/credito/saldo" in sample or "dèbito/crèdito/saldo" in sample))
        or ("bradesco" in sample and "débito/crédito/saldo" in sample)
        or nm.startswith("bradesco")
    )


def parse_bradesco_layout1(pdf_path, pdf_password=None):
    """Parser para extrato Bradesco com duas tabelas por página.

    Cada página contém duas tabelas lado a lado, equivalentes a uma sequência
    cronológica contínua: esquerda -> direita -> próxima página esquerda.
    O parser usa a posição textual fixa gerada pelo PDF para separar as duas
    metades da página, ignora SALDO/TRANSPORTE e mantém continuidades de
    histórico até surgir nova data.
    """

    date_re = re.compile(
        r"^(?P<data>\d{2}/\d{2}/\d{2})\s+"
        r"(?P<body>.+?)\s+"
        r"(?P<valor>\d{1,3}(?:\.\d{3})*,\d{2}-?)\s*$"
    )
    money_status_re = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}(?:CR|DV)\s*$", re.I)
    # Linha de cabeçalho operacional do Bradesco, no padrão:
    # "NOME DO TITULAR 0326-3 20.917-1". Em páginas com duas tabelas,
    # esse trecho pode aparecer entre o início e a continuação de uma descrição.
    # Não é dado transacional e não deve ser incorporado ao histórico.
    header_ag_conta_re = re.compile(r"\b\d{4}-\d\s+\d{1,3}\.\d{3}-\d\b")

    def bradesco_amount_to_float(tok: str) -> float:
        t = norm_space(str(tok or "")).replace("−", "-")
        neg = t.endswith("-")
        if neg:
            t = t[:-1]
        val = float(t.replace(".", "").replace(",", "."))
        return -val if neg else val

    def is_saldo_transporte(seg: str) -> bool:
        up = norm_space(str(seg or "")).upper()
        return up.startswith("SALDO ") or up.startswith("SALDO EM") or up.startswith("TRANSPORTE")

    def is_header_footer(seg: str) -> bool:
        up = norm_space(str(seg or "")).upper()
        if not up:
            return True
        fixed = {
            "CONTA", "NOME", "AGENCIA", "AGÊNCIA", "FOLHA", "EMISSÁO", "EMISSAO",
            "DATA", "HISTÓRICO", "HISTORICO", "DOCUMENTO", "DÈBITO/CRÈDITO/SALDO",
            "DEBITO/CREDITO/SALDO", "CONTA CORRENTE", "SCEX59"
        }
        if up in fixed:
            return True
        starts = (
            "EXTRATO PARA", "SIMPLES CONFER", "EMISS", "DATA HIST", "DADOS DE EMISS",
            "EXTRATO SEGUNDA", "SEGUNDA VIA", "NDA VIA"
        )
        if up.startswith(starts):
            return True
        if re.fullmatch(r"(?:\d{7}\s*){1,4}", up):
            return True
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}\s+\d{2}(?:\s+\d{2}/\d{2}/\d{4}\s+\d{2})?", up):
            return True
        if "CONTA CORRENTE" in up:
            return True
        # Cabeçalho com titular + agência + conta, por exemplo:
        # "LB VERT ESTETICA LTDA 0326-3 20.917-1" ou
        # "LIFE EXPRESS LTDA 0268-2 320.954-7".
        if header_ag_conta_re.search(up):
            return True
        if "320.954" in up or "LIFE EXPRESS" in up:
            return True
        return False

    def parse_segment(seg: str, current: dict, rows: list):
        raw = str(seg or "").rstrip()
        ns = norm_space(raw)
        if not ns or is_header_footer(ns):
            return current

        m = date_re.match(ns)
        if m:
            if current:
                rows.append(current)

            data = m.group("data")
            body = norm_space(m.group("body"))
            valor_txt = m.group("valor")

            if is_saldo_transporte(body):
                return None

            parts = body.split()
            documento = ""
            if parts and re.fullmatch(r"\d{3,}", parts[-1]):
                documento = parts[-1]
                descricao = " ".join(parts[:-1])
            else:
                descricao = body

            dd, mm, yy = data.split("/")
            yyyy = "20" + yy if int(yy) < 70 else "19" + yy
            return {
                "Data": f"{dd}/{mm}/{yyyy}",
                "Descrição": descricao,
                "Documento": documento,
                "Valor": bradesco_amount_to_float(valor_txt),
            }

        # Sem data: é continuação do histórico da transação anterior, salvo
        # linhas técnicas/meramente informativas do extrato.
        if is_saldo_transporte(ns) or money_status_re.fullmatch(ns):
            return current
        if current:
            current["Descrição"] = norm_space(current["Descrição"] + " " + ns)
        return current

    doc, _ = open_pdf_with_password(pdf_path)
    try:
        rows = []
        current = None
        split_at = 103  # ponto fixo que separa as duas tabelas no texto extraído deste layout
        for page in doc:
            raw_lines = (page.get_text("text") or "").splitlines()
            left_segments = []
            right_segments = []
            for ln in raw_lines:
                left_segments.append(ln[:split_at].rstrip())
                right_segments.append(ln[split_at:].rstrip())

            # Ordem cronológica: tabela da esquerda, depois tabela da direita.
            for seg in left_segments:
                current = parse_segment(seg, current, rows)
            for seg in right_segments:
                current = parse_segment(seg, current, rows)

        if current:
            rows.append(current)
    finally:
        doc.close()

    df = pd.DataFrame(rows)
    if df.empty:
        return standardize(df)

    def bradesco_doc_cleaner(x):
        # Defesa adicional: se algum resíduo de agência/conta do cabeçalho
        # escapar da filtragem por linha, remove-o apenas na saída Bradesco.
        return norm_space(header_ag_conta_re.sub("", str(x or "")))

    return standardize(df, doc_cleaner=bradesco_doc_cleaner)




# ---------------- CrediSIS ----------------

def is_credisis_signature(txt: str, name: str = "") -> bool:
    sample = norm_space(str(txt or "")).lower()
    nm = str(name or "").lower()
    return (
        "credisis" in sample
        or "crd_0027" in sample
        or ("extrato de conta corrente" in sample and "cheques ordem num" in sample and "n.doc" in sample)
        or nm.startswith("credisis")
    )


def parse_credisis(pdf_path, pdf_password=None):
    """Parser CrediSIS.

    Layout observado:
      Data | N.Doc. | Histórico | Débitos | Créditos | Saldos | CHEQUES ORDEM NUMÉRICA

    A rotina lê as palavras por coordenadas para considerar apenas a tabela
    principal, da coluna Data até Créditos. A coluna Saldos e a seção Cheques
    são ignoradas. A data é mantida para os lançamentos subsequentes até que
    nova data seja encontrada.
    """

    money_re = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    def group_words_by_row(words, tol=2.8):
        words = sorted(words, key=lambda w: (float(w[1]), float(w[0])))
        rows = []
        cur = []
        cur_y = None
        for w in words:
            y = float(w[1])
            if cur_y is None or abs(y - cur_y) <= tol:
                cur.append(w)
                cur_y = y if cur_y is None else (cur_y + y) / 2.0
            else:
                rows.append(sorted(cur, key=lambda z: float(z[0])))
                cur = [w]
                cur_y = y
        if cur:
            rows.append(sorted(cur, key=lambda z: float(z[0])))
        return rows

    def join_words(words):
        return norm_space(" ".join(w[4] for w in sorted(words, key=lambda z: float(z[0]))))

    def amount_from_words(words):
        vals = [w[4] for w in words if money_re.match(str(w[4]))]
        return vals[0] if vals else ""

    def safe_money(tok):
        return money_to_float(tok)

    def is_noise_row(*parts):
        up = norm_space(" ".join(str(p or "") for p in parts)).upper()
        if not up:
            return True
        if up.startswith("SALDO ") or up == "SALDO ANTERIOR":
            return True
        if any(k in up for k in [
            "DATA N.DOC", "N.DOC. HIST", "DÉBITOS", "DEBITOS", "CRÉDITOS", "CREDITOS",
            "SALDOS", "CHEQUES ORDEM", "NUM. DATA VALOR", "PÁGINA:", "PAGINA:",
            "EXTRATO DE CONTA CORRENTE", "CREDISIS", "COOPESA", "CPF/CNPJ", "PERÍODO:",
            "PERIODO:", "BANCO:", "CONTA:", "OUVIDORIA", "POSIÇÃO DE SALDOS", "POSICAO DE SALDOS",
            "SALDO ATUAL", "CHEQUE ESPECIAL", "CRÉDITO ROTATIVO", "CREDITO ROTATIVO",
            "LANÇAMENTOS FUTURO", "LANCAMENTOS FUTURO", "APLICAÇÕES", "APLICACOES",
            "EMPRÉSTIMOS", "EMPRESTIMOS", "LIMITE PRÉ-APROVADO", "LIMITE PRE-APROVADO"
        ]):
            return True
        return False

    rows = []
    current_date = ""
    doc, _ = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            words = page.get_text("words") or []
            for group in group_words_by_row(words):
                # Despreza apenas a margem superior extrema e o rodapé real.
                # Nas páginas seguintes à primeira, o CrediSIS inicia a tabela logo no topo
                # (após repetir apenas o cabeçalho das colunas). Portanto, a faixa útil
                # precisa começar antes de 115; caso contrário, a primeira data da página
                # é perdida e os lançamentos recebem indevidamente a data anterior.
                y0 = min(float(w[1]) for w in group)
                if y0 < 45 or y0 > 795:
                    continue

                data_txt = join_words([w for w in group if float(w[0]) < 75])
                doc_txt = join_words([w for w in group if 75 <= float(w[0]) < 110])
                hist_txt = join_words([w for w in group if 110 <= float(w[0]) < 255])
                deb_txt = amount_from_words([w for w in group if 255 <= float(w[0]) < 310])
                cred_txt = amount_from_words([w for w in group if 310 <= float(w[0]) < 375])

                if date_re.match(data_txt):
                    current_date = data_txt
                elif data_txt and not doc_txt and current_date:
                    # Em algumas páginas o N.Doc. de registros sem data explícita fica
                    # levemente deslocado para a esquerda e sai na faixa da coluna Data.
                    # Nesses casos, trata-se como documento do lançamento corrente.
                    doc_txt = data_txt
                    data_txt = ""

                if is_noise_row(data_txt, doc_txt, hist_txt):
                    continue
                if not current_date or not doc_txt or not hist_txt:
                    continue
                if not deb_txt and not cred_txt:
                    continue

                if deb_txt and cred_txt:
                    # Situação inesperada; preserva a primeira movimentação da esquerda.
                    valor = -abs(safe_money(deb_txt))
                elif deb_txt:
                    valor = -abs(safe_money(deb_txt))
                else:
                    valor = abs(safe_money(cred_txt))

                if hist_txt.upper().startswith("SALDO"):
                    continue

                rows.append([current_date, hist_txt, doc_txt, valor])
    finally:
        doc.close()

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )


# ---------------- Safra layout2 ----------------

SAFRA_MONTHS = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}


def is_safra_layout2_signature(txt: str, name: str) -> bool:
    low = norm_space(txt).lower()
    # Assinatura estrutural, sem depender do nome do arquivo.
    has_safra = "banco safra" in low or "safrapay" in low or "safra sa" in low
    has_ref_month = bool(re.search(r"\b(?:jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{4}\b", low, re.I))
    has_cols = (
        "num.docto" in low or "num.docto." in low or
        "déb/créd" in low or "deb/cred" in low or
        ("demonstrativo consolidado" in low and "reais" in low)
    )
    has_safrapay_rows = "pix recebido safrapay" in low or "antecipacao rv" in low or "resumo vendas cartao deb" in low
    return bool(has_safra and has_ref_month and (has_cols or has_safrapay_rows))


def _safra_layout2_lines_by_tables(page, split_y=430.0):
    """Extrai linhas de uma página Safra separando a tabela esquerda da direita.

    O PDF vem girado internamente: no PyMuPDF, x representa a posição vertical
    e y representa a posição horizontal. Por isso a divisão das duas tabelas é
    feita pela coordenada y. A ordem cronológica é: tabela esquerda -> tabela direita.
    """
    words = page.get_text("words") or []
    out_segments = []
    for side in ["left", "right"]:
        selected = []
        for w in words:
            x0, y0, x1, y1, txt = w[:5]
            if side == "left" and y0 >= split_y:
                selected.append((float(x0), float(y0), str(txt)))
            elif side == "right" and y0 < split_y:
                selected.append((float(x0), float(y0), str(txt)))
        if not selected:
            out_segments.append([])
            continue
        selected.sort(key=lambda t: (t[0], -t[1]))
        rows = []
        current = []
        current_x = None
        tol = 2.0
        for x, y, txt in selected:
            if current_x is None or abs(x - current_x) <= tol:
                current.append((x, y, txt))
                if current_x is None:
                    current_x = x
            else:
                current.sort(key=lambda t: -t[1])
                line = norm_space(" ".join(t[2] for t in current))
                if line:
                    rows.append(line)
                current = [(x, y, txt)]
                current_x = x
        if current:
            current.sort(key=lambda t: -t[1])
            line = norm_space(" ".join(t[2] for t in current))
            if line:
                rows.append(line)
        out_segments.append(rows)
    return out_segments


def parse_safra_layout2(pdf_path, pdf_password=None):
    """Parser do Safra em demonstrativo consolidado com duas tabelas por página.

    Estrutura observada:
      Data | Descrição | Num.Docto. | Déb/Créd | Saldos

    A leitura é feita por segmentos de tabela, separando esquerda e direita por
    coordenadas. Isso evita que, quando ambas as tabelas da mesma página possuem
    movimentos, uma linha da esquerda seja misturada com outra da direita.
    """
    month_header_re = re.compile(r"\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{4})\b", re.I)
    page_date_re = re.compile(r"^\s*(\d{4})/(\d{2})/(\d{2})\s*$")
    date_re = re.compile(r"^\s*(\d{2}/\d{2})\s+(.+?)\s*$")
    money_re = re.compile(r"^(?:-)?\d{1,3}(?:\.\d{3})*,\d{2}-?$")
    cpf_cnpj_re = re.compile(r"\b(?:\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\b")

    def parse_amount_token(tok: str):
        tok = norm_space(tok)
        if not money_re.match(tok):
            return None
        return money_to_float(tok)

    def is_control_line(line: str) -> bool:
        raw = str(line or "")
        low = norm_space(raw).lower()
        if not low:
            return True
        if low.startswith("continua"):
            return True
        if page_date_re.match(raw):
            return True
        if month_header_re.search(raw):
            return True
        if re.fullmatch(r"\d{3}/\d{3}", low):
            return True
        if "casa ludica brinquedoteca ltda" in low and not cpf_cnpj_re.search(raw):
            return True
        if any(x in low for x in [
            "banco safra", "demonstrativo consolidado", "cnpj:", "conta nº", "conta n",
            "empresa:", "cliente:", "limite:", "agência:", "agencia:", "data descrição",
            "data descricao", "num.docto", "déb/créd", "deb/cred", "saldos",
            "previsão para débito", "previsao para debito", "valor r$", "folha ",
            "extrato mensal", "por período", "por periodo", "legenda:", "sac -", "ouvidoria",
            "(p)pessoal", "(e)eletronico", "(i)internet", "(tar)tarifa", "correspondente no pais",
            "correspondente no país", "data previsao", "previsao para"
        ]):
            return True
        return False

    def is_saldo_row(desc: str) -> bool:
        up = norm_space(desc).upper()
        return up in {"CONTA CORRENTE", "SALDO", "SALDO ANTERIOR", "SALDO FINAL"} or up.startswith("SALDO ")

    def split_body_amount_doc(body: str):
        parts = norm_space(body).split()
        if not parts:
            return "", "", None
        amount_idx = None
        for idx in range(len(parts) - 1, -1, -1):
            if money_re.match(parts[idx]):
                amount_idx = idx
                break
        if amount_idx is None:
            return norm_space(body), "", None
        valor = parse_amount_token(parts[amount_idx])
        before = parts[:amount_idx]
        if not before:
            return "", "", valor
        doc = ""
        if re.fullmatch(r"[A-Z0-9./-]{3,}", before[-1], flags=re.I) and any(ch.isdigit() for ch in before[-1]):
            if not re.fullmatch(r"\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", before[-1]):
                doc = before[-1]
                before = before[:-1]
        desc = norm_space(" ".join(before))
        return desc, doc, valor

    def add_continuation(rows, idx, cont: str):
        cont = norm_space(cont)
        if idx is None or idx < 0 or idx >= len(rows) or not cont:
            return idx
        if is_control_line(cont):
            return idx
        up = cont.upper()
        if up.startswith("CONTA CORRENTE") or up.startswith("SALDO") or up.startswith("TOTAL"):
            return idx
        rows[idx][1] = norm_space(rows[idx][1] + " " + cont)
        return idx

    rows = []
    pending_idx = None
    current_year = None

    doc, _pw = open_pdf_with_password(pdf_path)
    try:
        if pdf_password is not None:
            try:
                doc.authenticate(pdf_password)
            except Exception:
                pass
        for page in doc:
            for segment_lines in _safra_layout2_lines_by_tables(page):
                for raw_line in segment_lines:
                    line = norm_space(raw_line)
                    if not line:
                        continue

                    mp = page_date_re.match(line)
                    if mp:
                        current_year = int(mp.group(1))
                        continue

                    mh = month_header_re.search(line)
                    if mh:
                        current_year = int(mh.group(2))
                        continue

                    m = date_re.match(line)
                    if m:
                        dm = m.group(1)
                        body = m.group(2)
                        desc, docnum, valor = split_body_amount_doc(body)
                        if valor is None:
                            pending_idx = None
                            continue
                        if is_saldo_row(desc):
                            pending_idx = None
                            continue
                        if not current_year:
                            pending_idx = None
                            continue
                        try:
                            data = datetime.strptime(f"{dm}/{current_year}", "%d/%m/%Y").strftime("%d/%m/%Y")
                        except Exception:
                            pending_idx = None
                            continue
                        if desc and valor != 0:
                            rows.append([data, desc, docnum, valor])
                            pending_idx = len(rows) - 1
                        else:
                            pending_idx = None
                        continue

                    if not is_control_line(line):
                        pending_idx = add_continuation(rows, pending_idx, line)
    finally:
        doc.close()

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )


# ---------------- Sicoob layout 1 (SISBR - Histórico de Movimentação) ----------------

def is_sicoob_layout1_signature(txt: str, name: str = "") -> bool:
    """Reconhece o extrato Sicoob/SISBR com tabela DATA | HISTÓRICO | VALOR.

    Assinatura observada:
      - PLATAFORMA DE SERVIÇOS FINANCEIROS DO SICOOB - SISBR;
      - EXTRATO CONTA CORRENTE;
      - COOP. / CONTA / PERÍODO;
      - quadro HISTÓRICO DE MOVIMENTAÇÃO com colunas DATA, HISTÓRICO e VALOR.
    """
    sample = norm_space(str(txt or "")).lower()
    name_low = norm_space(str(name or "")).lower()
    has_bank = (
        "plataforma de serviços financeiros do sicoob - sisbr" in sample or
        "plataforma de servicos financeiros do sicoob - sisbr" in sample
    )
    has_table = (
        ("histórico de movimentação" in sample or "historico de movimentacao" in sample)
        and "extrato conta corrente" in sample
        and "coop.:" in sample
        and "conta:" in sample
        and ("período:" in sample or "periodo:" in sample)
    )
    # Nome serve apenas como reforço; a assinatura estrutural continua obrigatória.
    name_hint = name_low.startswith("sicoob") or name_low.startswith("sicob")
    return has_table and (has_bank or (name_hint and "sicoob" in sample))


def parse_sicoob_layout1(lines):
    """Parser para Sicoob/SISBR - Histórico de Movimentação.

    Regras específicas do layout:
      - cada lançamento inicia por data DD/MM, ainda que repita a mesma data;
      - o ano é inferido a partir do PERÍODO declarado no cabeçalho;
      - a descrição pode ocupar várias linhas e é recomposta até o próximo lançamento;
      - o valor pode vir como ``1.234,56D``/``1.234,56C`` ou com C/D na linha seguinte;
      - ``DOC.:`` é preservado na coluna Documento;
      - SALDO ANTERIOR, SALDO BLOQ.ANTERIOR e SALDO DO DIA não são transações;
      - a leitura encerra no quadro RESUMO.
    """
    date_re = re.compile(r"^(\d{2})/(\d{2})$")
    amount_cd_re = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})\s*(?P<dc>[CD])$", re.I)
    amount_only_re = re.compile(r"^(?P<val>\d{1,3}(?:\.\d{3})*,\d{2})$")
    doc_re = re.compile(r"^DOC\.\s*:\s*(.*)$", re.I)

    # Período é a referência de ano porque as linhas da tabela trazem apenas DD/MM.
    period_start = None
    period_end = None
    period_re = re.compile(
        r"(?:PERÍODO|PERIODO)\s*:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})",
        re.I,
    )
    for ln in lines[:120]:
        m = period_re.search(norm_space(ln))
        if not m:
            continue
        try:
            period_start = datetime.strptime(m.group(1), "%d/%m/%Y")
            period_end = datetime.strptime(m.group(2), "%d/%m/%Y")
        except Exception:
            period_start = period_end = None
        break

    if period_start is None or period_end is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    # Localiza somente o quadro de movimentação. Cabeçalho e resumo ficam fora do parser.
    start_idx = None
    for idx, ln in enumerate(lines):
        low = norm_space(ln).lower()
        if low in {"histórico de movimentação", "historico de movimentacao"}:
            start_idx = idx + 1
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"])

    def infer_full_date(ddmm: str):
        """Resolve DD/MM para o ano compatível com o período declarado."""
        try:
            day, month = map(int, ddmm.split("/"))
        except Exception:
            return None

        candidates = []
        for year in range(period_start.year - 1, period_end.year + 2):
            try:
                dt = datetime(year, month, day)
            except Exception:
                continue
            if period_start <= dt <= period_end:
                return dt
            candidates.append(dt)

        # Linhas de saldo anterior podem estar imediatamente fora da janela.
        # Elas serão descartadas, mas esta aproximação mantém o parser determinístico.
        if not candidates:
            return None
        def distance_to_period(dt):
            if dt < period_start:
                return (period_start - dt).days
            if dt > period_end:
                return (dt - period_end).days
            return 0
        return min(candidates, key=distance_to_period)

    def is_balance_desc(desc: str) -> bool:
        low = norm_space(desc).lower()
        return (
            "saldo anterior" in low or
            "saldo bloq.anterior" in low or
            "saldo bloq. anterior" in low or
            "saldo bloqueado anterior" in low or
            "saldo do dia" in low
        )

    rows = []
    i = start_idx
    N = len(lines)

    # Cabeçalho linear DATA / HISTÓRICO / VALOR pode aparecer logo após o título.
    header_noise = {"data", "histórico", "historico", "valor"}

    while i < N:
        ln = norm_space(lines[i])
        low = ln.lower()
        if low == "resumo":
            break
        if low in header_noise:
            i += 1
            continue

        dm = date_re.match(ln)
        if not dm:
            i += 1
            continue

        # O bloco pertence a uma única transação: vai até a próxima data ou RESUMO.
        j = i + 1
        block = []
        while j < N:
            cur = norm_space(lines[j])
            if cur.lower() == "resumo" or date_re.match(cur):
                break
            if cur:
                block.append(cur)
            j += 1

        if not block:
            i = j
            continue

        primary_desc = block[0]
        if is_balance_desc(primary_desc):
            i = j
            continue

        amount_idx = None
        sign_idx = None
        valor = None

        for k, item in enumerate(block):
            m = amount_cd_re.fullmatch(item)
            if m:
                raw_val = money_to_float(m.group("val"))
                valor = abs(raw_val) if m.group("dc").upper() == "C" else -abs(raw_val)
                amount_idx = k
                break

            m = amount_only_re.fullmatch(item)
            if m and k + 1 < len(block) and block[k + 1].upper() in {"C", "D"}:
                raw_val = money_to_float(m.group("val"))
                dc = block[k + 1].upper()
                valor = abs(raw_val) if dc == "C" else -abs(raw_val)
                amount_idx = k
                sign_idx = k + 1
                break

        if valor is None or amount_idx is None:
            # Sem valor assinado não há lançamento confiável neste layout.
            i = j
            continue

        documento = ""
        desc_parts = []
        for k, item in enumerate(block):
            if k == amount_idx or k == sign_idx:
                continue
            if item.upper() in {"C", "D"}:
                continue
            md = doc_re.match(item)
            if md:
                cand = norm_space(md.group(1))
                if cand and not documento:
                    documento = cand
                continue
            if item.lower() in header_noise:
                continue
            desc_parts.append(item)

        descricao = norm_space(" ".join(desc_parts))
        dt = infer_full_date(ln)
        if dt is not None and descricao and valor != 0:
            rows.append([dt.strftime("%d/%m/%Y"), descricao, documento, valor])

        i = j

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: norm_space(str(x))[:80]
    )



# ---------------- Arbi layout 1 ----------------

def is_arbi_layout1_signature(txt: str, name: str = "") -> bool:
    """Reconhece extrato de conta vinculada do Banco Arbi (código 213)."""
    sample = norm_space(str(txt or "")).lower()
    nm = str(name or "").lower()
    return (
        (
            "extrato de conta vinculada" in sample
            and ("banco: 213 - arbi" in sample or "banco: 213" in sample)
            and "data descrição valor" in sample
        )
        or nm.startswith("arbi_layout1")
    )


def _parse_arbi_layout1_by_geometry(page):
    """Extrai as linhas usando a geometria da tabela.

    As linhas horizontais da grade delimitam cada transação; Data e Valor funcionam
    como âncoras, e todo o texto da coluna central é concatenado como Descrição.
    """
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    amount_re = re.compile(r"^-?\d+(?:[.,]\d+)*,\d{2}$")

    words = page.get_text("words") or []
    if not words:
        return []

    date_words = []
    amount_words = []
    for w in words:
        x0, y0, x1, y1, token = w[:5]
        token = norm_space(token)
        yc = (float(y0) + float(y1)) / 2.0
        if float(x0) < 180 and date_re.fullmatch(token):
            date_words.append((yc, token))
        elif float(x0) > 430 and amount_re.fullmatch(token):
            amount_words.append((yc, token))

    # Pareia Data e Valor que ocupam a mesma linha/célula vertical.
    anchors = []
    used_amounts = set()
    for dy, dt in date_words:
        candidates = [
            (abs(ay - dy), idx, amt)
            for idx, (ay, amt) in enumerate(amount_words)
            if idx not in used_amounts and abs(ay - dy) <= 4.0
        ]
        if not candidates:
            continue
        _, idx, amount_txt = min(candidates)
        used_amounts.add(idx)
        anchors.append((dy, dt, amount_txt))
    anchors.sort(key=lambda x: x[0])
    if not anchors:
        return []

    # Limites horizontais reais da grade: evitam incorporar a descrição da linha seguinte.
    boundaries = []
    try:
        for drawing in page.get_drawings() or []:
            for item in drawing.get("items", []):
                if not item or item[0] != "l":
                    continue
                p1, p2 = item[1], item[2]
                if abs(float(p1.y) - float(p2.y)) <= 0.7 and abs(float(p2.x) - float(p1.x)) >= 150 and float(p1.y) > 130:
                    boundaries.append(round(float(p1.y), 1))
    except Exception:
        boundaries = []
    boundaries = sorted(set(boundaries))

    desc_words = []
    for w in words:
        x0, y0, x1, y1, token = w[:5]
        if 190 <= float(x0) < 420:
            desc_words.append(((float(y0) + float(y1)) / 2.0, float(x0), norm_space(token)))

    rows = []
    for i, (yc, dt, amount_txt) in enumerate(anchors):
        if boundaries:
            prev_y = [y for y in boundaries if y < yc]
            next_y = [y for y in boundaries if y > yc]
            upper = max(prev_y) if prev_y else (137.0 if i == 0 else (anchors[i - 1][0] + yc) / 2.0)
            lower = min(next_y) if next_y else (810.0 if i == len(anchors) - 1 else (yc + anchors[i + 1][0]) / 2.0)
        else:
            upper = 137.0 if i == 0 else (anchors[i - 1][0] + yc) / 2.0
            lower = 810.0 if i == len(anchors) - 1 else (yc + anchors[i + 1][0]) / 2.0

        selected = [(y, x, tok) for y, x, tok in desc_words if upper < y < lower]
        selected.sort(key=lambda t: (round(t[0], 1), t[1]))
        desc = norm_space(" ".join(tok for _, _, tok in selected))
        if not desc:
            continue
        try:
            valor = money_to_float(amount_txt)
        except Exception:
            continue
        if valor != 0:
            rows.append([dt, desc, "", valor])
    return rows

def parse_arbi_layout1(pdf_path, pdf_password=None):
    """Parser do Arbi layout1.

    Estrutura:
        Data | Descrição | Valor

    - Cada linha lógica da tabela representa uma transação.
    - Descrições com quebra de linha são concatenadas integralmente.
    - O sinal antes do valor define débito/crédito.
    - Não há campo Documento nesse layout; permanece vazio.
    """
    rows = []
    doc, _ = open_pdf_with_password(pdf_path)
    try:
        for page in doc:
            # A grade vetorial é suficiente e evita dependência de OCR ou de APIs novas.
            page_rows = _parse_arbi_layout1_by_geometry(page)

            # Fallback para PDFs Arbi equivalentes em que a grade não seja exposta por get_drawings().
            if not page_rows and hasattr(page, "find_tables"):
                try:
                    tables = page.find_tables()
                    for table in getattr(tables, "tables", []):
                        for row in (table.extract() or []):
                            if not row or len(row) < 3:
                                continue
                            dt = norm_space(row[0])
                            if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", dt):
                                continue
                            desc = norm_space(str(row[1] or "").replace("\n", " "))
                            amount_txt = norm_space(str(row[2] or "").replace("R$", ""))
                            if not desc or not amount_txt:
                                continue
                            try:
                                valor = money_to_float(amount_txt)
                            except Exception:
                                continue
                            if valor != 0:
                                page_rows.append([dt, desc, "", valor])
                except Exception:
                    page_rows = []

            rows.extend(page_rows)
    finally:
        doc.close()

    return standardize(
        pd.DataFrame(rows, columns=["Data", "Descrição", "Documento", "Valor"]),
        doc_cleaner=lambda x: ""
    )

# ---------------- Dispatcher ----------------

def parse_one_pdf(pdf_path):
    lines, pdf_password = extract_lines(pdf_path)
    name = os.path.basename(pdf_path).lower()

    if not lines:
        # Fallback OCR estritamente localizado para o Itaú layout7 em PDF-imagem.
        # Em PDFs escaneados de outros layouts, mantém o comportamento anterior.
        if name.startswith("itau") or pytesseract is not None:
            try:
                df_img = parse_itau_layout7_image(pdf_path, pdf_password=pdf_password)
                if df_img is not None and not df_img.empty:
                    return "itau_layout7", df_img
            except RuntimeError:
                # Se o nome indicar Itaú, propaga mensagem útil de dependência OCR.
                if name.startswith("itau"):
                    raise
        return "", pd.DataFrame()

    txt = " ".join(lines[:250]).lower()

    if "extrato exportado no dia" in txt and "data contábil" in txt and "saldo do dia" in txt:
        return "c6", parse_c6(lines, pdf_path)

    if "extrato histórico da conta" in txt and "data e hora" in txt and "data mov." in txt and ("nr.doc." in txt or "nr.doc" in txt):
        return "caixa_layout2", parse_caixa_layout2(pdf_path, pdf_password=pdf_password)

    if "sihex" in txt and ("sistema de histórico de extratos" in txt or "sistema de historico de extratos" in txt) and "data mov." in txt and "nr. doc." in txt:
        return "caixa_layout1", parse_caixa_layout1(lines)

    if is_bb_layout5_signature(txt, name):
        return "bb_layout5", parse_bb_layout5(pdf_path, pdf_password=pdf_password)

    if ("consultas - extrato de conta corrente" in txt) or ("sisbb" in txt) or name.startswith("bb_layout"):
        layout, df = parse_bb_auto(lines)
        return layout, df

    if "banco abc" in txt or name == "abc.pdf":
        return "abc", parse_abc(lines)

    if "banrisul" in txt or "b a n r i s u l" in txt or name.startswith("banrisul"):
        return "banrisul", parse_banrisul(lines)

    if is_credisis_signature(txt, name):
        return "credisis", parse_credisis(pdf_path, pdf_password=pdf_password)

    if is_sicoob_layout1_signature(txt, name):
        return "sicoob_layout1", parse_sicoob_layout1(lines)

    if is_arbi_layout1_signature(txt, name):
        return "arbi_layout1", parse_arbi_layout1(pdf_path, pdf_password=pdf_password)

    if is_sicredi_layout3_ccpi_signature(txt, name):
        return "sicredi_layout3", parse_sicredi_layout3_ccpi(pdf_path, pdf_password=pdf_password)

    if "sicredi" in txt or "cooperativa:" in txt:
        return "sicredi", parse_sicredi(lines, pdf_path=pdf_path, pdf_password=pdf_password)

    if is_bradesco_layout3_signature(txt, name):
        return "bradesco_layout3", parse_bradesco_layout3(pdf_path, pdf_password=pdf_password)

    if is_bradesco_layout2_signature(txt, name):
        return "bradesco_layout2", parse_bradesco_layout2(pdf_path, pdf_password=pdf_password)

    if is_bradesco_layout1_signature(txt, name):
        return "bradesco_layout1", parse_bradesco_layout1(pdf_path, pdf_password=pdf_password)

    if is_safra_layout2_signature(txt, name):
        return "safra_layout2", parse_safra_layout2(pdf_path, pdf_password=pdf_password)

    if is_santander_layout4_signature(txt, name):
        return "santander_layout4", parse_santander_layout4(lines)

    if "santander" in txt or "contamax" in txt or name.startswith("santander"):
        return "santander", parse_santander(pdf_path, lines, pdf_password=pdf_password)

    if ("banco inter" in txt and "santander" not in txt) or "saldo por transação" in txt or name.startswith("inter"):
        return "inter", parse_inter(lines)

    if ("itaú" in txt or "itau" in txt or "extrato mensal" in txt or
        is_itau_layout3_signature(txt) or is_itau_layout4_signature(txt) or
        is_itau_layout5_signature(txt) or is_itau_layout6_signature(txt) or
        name.startswith("itau") or name == "itau.pdf"):
        return "itau", parse_itau(lines, pdf_path=pdf_path, pdf_password=pdf_password)

    if is_unicred_layout2_signature(txt, name):
        return "unicred_layout2", parse_unicred_layout2(pdf_path, pdf_password=pdf_password)

    if "unicred" in txt or "instituição financeira:  136" in txt.lower() or name == "unicred.pdf":
        return "unicred", parse_unicred(pdf_path, pdf_password=pdf_password)

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


# ---------------- Classificação provável de créditos (Glossário Mestre v1) ----------------
#
# Módulo pós-processamento: NÃO participa da leitura dos PDFs/OFX, do dispatcher,
# dos parsers bancários nem de standardize(). Atua somente sobre os lançamentos
# já consolidados e apenas sobre créditos (Valor > 0).
#
# Fonte documental: Glossário Mestre de Classificação de Créditos Bancários — v1.
# Saída no Consolidado:
#   - Classificação provável
#   - Categoria
#
# A classificação é indiciária para triagem fiscal. As regras de exclusão/revisão
# têm precedência sobre as regras positivas de receita.

GLOSSARIO_VERSION = "v1"
COL_CLASSIFICACAO = "Classificação provável"
COL_CATEGORIA = "Categoria"
CLASSIFICACAO_RECEITA = "Presunção de receita tributável"
CLASSIFICACAO_SEM_PRESUNCAO = "Sem presunção de receita operacional"
CLASSIFICACAO_REVISAO = "Necessária revisão manual"


def _class_rule(priority, classification, foundation, pattern=None, kind="regex"):
    return {
        "priority": int(priority),
        "classification": classification,
        "foundation": foundation,
        "pattern": pattern,
        "kind": kind,
    }


# Os padrões são aplicados sobre uma cópia normalizada da descrição (maiúsculas,
# sem acentos, espaços normalizados). As descrições originais permanecem intactas.
CREDIT_CLASSIFICATION_RULES = [
    # Excludentes / revisão prioritária.
    _class_rule(1, CLASSIFICACAO_SEM_PRESUNCAO, "Mesma titularidade",
                r"\b(?:MESMA|MMA)\s*TIT|MESMA TITULARIDADE|ENTRE CONTAS PROPRIAS|PROPRIA TITULARIDADE|MESMO TITULAR"),
    _class_rule(2, CLASSIFICACAO_SEM_PRESUNCAO, "Estorno/devolução",
                r"ESTORN|\bEST\b|\bDEV\b|DEVOL|REJEIT|CANCEL|REVERS|RETORN.*PIX|PIX.*RETORN"),
    _class_rule(3, CLASSIFICACAO_SEM_PRESUNCAO, "Aplicação/resgate",
                r"RESG|RES APLIC|RESGATE|APLIC AUT|INVEST FACIL|RENDE FACIL|\bBB RF\b|RF CP|RF MAIS|RF AUTOM|RF REF|FUNDO|\bCDB\b|POUPANCA|CONTAMAX|SDO CTA APL|BAIXA AUTOMAT"),
    _class_rule(4, CLASSIFICACAO_SEM_PRESUNCAO, "Rendimento financeiro",
                r"REND(?:IMENTO|IMENTOS| PAGO|\b)|RENTAB|JUROS.*CRED|REMUNERACAO|ATUALIZACAO MONET|CORRECAO MONET"),
    _class_rule(5, CLASSIFICACAO_SEM_PRESUNCAO, "Empréstimo/financiamento",
                r"EMPREST|FINANC|OP[ .]?CREDITO|OPERACAO DE CREDITO|GIRO|PRONAMPE|PENHOR|CAPITAL GIRO|ROTATIVO|CHEQUE ESPECIAL|LIBERACAO EMPR|CRED EMPR|CRED CA/CL|BBH.*CRED|BBH.*RET"),
    _class_rule(6, CLASSIFICACAO_SEM_PRESUNCAO, "Consórcio/composição de dívida",
                r"CONSORCIO|CARTA DE CREDITO|COMPOSICAO DE DIVIDA"),
    _class_rule(7, CLASSIFICACAO_SEM_PRESUNCAO, "Capitalização",
                r"BRASILCAP|CAPITALIZAC"),
    _class_rule(8, CLASSIFICACAO_SEM_PRESUNCAO, "Sinistro/indenização",
                r"SINISTRO|INDENIZ"),
    _class_rule(9, CLASSIFICACAO_SEM_PRESUNCAO, "Restituição/ressarcimento",
                r"RESTIT|RESSARC|CASHBACK|REEMBOLSO|REEMB|CB IOF SUS|IOF/SUS"),
    _class_rule(10, CLASSIFICACAO_SEM_PRESUNCAO, "Saldo/ajuste contábil",
                r"^SALDO\b|^SDO\b|SALDO A LIBERAR|MOVIMENTO DO DIA|RECLASSIF.*SDO"),
    _class_rule(11, CLASSIFICACAO_SEM_PRESUNCAO, "Transferência possivelmente interna",
                r"TRANSFERENC\.?\s+DE AGENCIA"),
    _class_rule(12, CLASSIFICACAO_SEM_PRESUNCAO, "Ajuste/código bancário",
                r"^COD\.?\s*LANC\.?\s*\d+$"),
    _class_rule(13, CLASSIFICACAO_REVISAO, "Bloqueio/desbloqueio judicial",
                r"(?:BLOQ|DESBL).*JUDICIAL|BACEN JUD|DESBLOQ.*ORDEM JUDICIAL"),
    _class_rule(14, CLASSIFICACAO_REVISAO, "Crédito BACEN", r"CREDITO BACEN"),
    _class_rule(15, CLASSIFICACAO_REVISAO, "Crédito genérico/indeterminado",
                r"CRED AUTOM|CREDITO CONTABILIDADE"),
    _class_rule(16, CLASSIFICACAO_REVISAO, "Distribuição cooperativa/sobras", r"\bSOBRAS\b"),
    _class_rule(17, CLASSIFICACAO_REVISAO, "Descrição insuficiente", r"^0+$"),
    _class_rule(18, CLASSIFICACAO_REVISAO, "Possível artefato de extração",
                r"INTERNET.*BANKING|HTTP[S]?://|SIIBC|IMPRIME_EXT_PERIODO"),
    _class_rule(19, CLASSIFICACAO_REVISAO, "Possível artefato/cotação",
                r"^(?:DOLAR COMERCIAL|EURO|SALARIO MINIMO)$"),
    _class_rule(20, CLASSIFICACAO_REVISAO, "Possível artefato numérico", kind="numeric_artifact"),

    # Regras positivas de receita, aplicadas somente após as anteriores.
    _class_rule(30, CLASSIFICACAO_RECEITA, "Antecipação/desconto de recebíveis",
                r"ANTECIP|ANTEC RV|SOROCRED ANTEC|SOROCR VISA AT|DESCONTO.*(?:DUPLIC|TITUL|RECEB)|DESC.*DUPLIC|CESSAO.*RECEB|ADIANT.*RECEB"),
    _class_rule(31, CLASSIFICACAO_RECEITA, "Recebível/pagamento a fornecedor",
                r"AQUISICAO FORNECEDORES|PAGAMENTO\s+FORNECEDOR|PAGAMENTO A FORNECEDORES|PAG FORNEC"),
    _class_rule(32, CLASSIFICACAO_RECEITA, "Cartões/adquirentes",
                r"GETNET|STONE|STON\b|REDE\b|VERO|BANRICOMPRAS|PAGAMENTO CARTAO|CARTAO DE (?:CREDITO|DEBITO)|CRED LOJ.*CART|CR VD CART|ELO DB|VISA DB|MAST DB|MC CC|VS CC|MC CD|VS CD|EL CD|SAFRAPAY|CIELO|CIEL\b|PAGSEGURO|MERCADO PAGO.*CART|SUMUP"),
    _class_rule(33, CLASSIFICACAO_RECEITA, "Cobrança/boleto/título",
                r"COBRAN|COBINT|LIQ.*COB|LIQUIDACAO.*COB|MOV TIT COB|LIQ.*BOL|BOLETOS? RECEB|CRED.*COB|CONF RECEBIMENTO|BOL PAGO|CREDITO\s+TITULOS"),
    _class_rule(34, CLASSIFICACAO_RECEITA, "PIX recebido",
                r"\bPIX\b|CRED PIX|PIX RECEB|ENTRADA PIX|TRANSF.*PIX"),
    _class_rule(35, CLASSIFICACAO_RECEITA, "TED/DOC recebido",
                r"\bTED\b|CRED TED|TED RECEB|RECEBIMENTO DE TED|\bDOC\b|DOC E|DOC CREDITO|C DOC"),
    _class_rule(36, CLASSIFICACAO_RECEITA, "SISPAG/pagamento programado", r"SISPAG|CX PROGRAM"),
    _class_rule(37, CLASSIFICACAO_RECEITA, "Transferência recebida de terceiro",
                r"TRANSF.*RECEB|TRANSFERENCIA RECEB|TRANSFERENCIA A CREDITO|TRANSFERENCIA AGENDADA|CREDITO TRANSFERENCIA|CRED TEV|AG TEF|\bTEF\b|\bTBI\b|BDX CREDITO TRANSFERENCI"),
    _class_rule(38, CLASSIFICACAO_RECEITA, "Transferência genérica sem excludente",
                r"TRANSFERENCIA ENTRE CONTAS|TRANSFERENCIA VALORES|BNO TR\. VALOR|\bTRANSF\b|\bTR ONLINE\b"),
    _class_rule(39, CLASSIFICACAO_RECEITA, "Depósito",
                r"\bDEP\b|DEPOS|DP DIN|DEP DINHEIRO|RECICLADOR|DEP CORBAN|CREDITO DEPOSITO|CEI.*DINHEIRO"),
    _class_rule(40, CLASSIFICACAO_RECEITA, "Recebimento/pagamento identificado",
                r"RECEBIMENTO|RECEB POR FORNECIMENTO|PAGAMENTO RECEBIDO|CREDITO CLIENTE|PGTO CLIENTE|REPASSE|CREDITO VENDA|PAGTO EXPRESSO MERCURIO"),
    _class_rule(41, CLASSIFICACAO_RECEITA, "Saúde/convênios",
                r"CR P SAUDE|CAIXA DE ASSISTENCIA|CASSI|GEAP|CREDITO CASSI"),
    _class_rule(42, CLASSIFICACAO_RECEITA, "Órgão público/repasse institucional",
                r"GRUPAMENTO|MUNICIPIO|COMANDO DO COMANDO MIL|SESC"),
    _class_rule(43, CLASSIFICACAO_RECEITA, "Operação de câmbio/recebimento exterior",
                r"OPERACAO DE CAMBIO|LIBERACAO CAMBIO|OP\s*(?:REC|RECEBIDA).*EXT"),
    _class_rule(44, CLASSIFICACAO_RECEITA, "Liquidação de exportação — revisar tributabilidade",
                r"LIQ\s*EXPORT"),
    _class_rule(45, CLASSIFICACAO_RECEITA, "Liberação de garantia", r"TEG EX GAR"),
    _class_rule(46, CLASSIFICACAO_RECEITA, "Entrada com contraparte identificada", r"^ENTRADA\s+.+"),
    _class_rule(47, CLASSIFICACAO_RECEITA, "Contraparte identificada", kind="counterparty"),
    _class_rule(48, CLASSIFICACAO_RECEITA, "Pagamento genérico creditado",
                r"\b(?:PAGAMENTO|PAGTO|PGTO)\b"),
    _class_rule(49, CLASSIFICACAO_RECEITA, "Convênio/arrecadação repassada", r"CONVENIO ARREC"),
    _class_rule(50, CLASSIFICACAO_RECEITA, "Entrada eletrônica genérica", r"\bE ELECTRON\b"),
]


def normalize_classification_text(value: str) -> str:
    """Normaliza somente a cópia usada pelo glossário; não altera a descrição original."""
    s = unicodedata.normalize("NFKD", str(value or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper().replace("−", "-").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _mask_numeric_artifact(series: pd.Series) -> pd.Series:
    # Ex.: "7.384,27 C 00" ou linhas compostas quase só por valor/código.
    # Mantém esta regra em Revisão manual, sem excluir o lançamento.
    return series.str.fullmatch(
        r"[\d\s.,/\-]+(?:\s+C(?:\s+\d+)?)?",
        na=False,
    )


def _looks_like_counterparty_text(value: str) -> bool:
    """Heurística subsidiária para nome/razão social do remetente.

    Só é chamada depois de todas as regras específicas. Evita usar um simples
    'qualquer texto = receita'. Retorna True para nomes/razões sociais plausíveis
    sem vocabulário bancário genérico remanescente.
    """
    s = normalize_classification_text(value)
    if not s:
        return False

    blocked = (
        "CREDITO", "DEBITO", "TRANSFER", "TRANSF", "PAGAMENTO", "PAGTO", "PGTO",
        "DEPOS", "PIX", "TED", "DOC", "SALDO", "JUROS", "TARIFA", "APLIC", "RESG",
        "EMPREST", "FINANC", "CONSORC", "ESTORN", "DEVOL", "BACEN", "BLOQ", "DESBL",
        "CONVENIO", "CAMBIO", "COBRAN", "BOLETO", "LIQUID", "LANC", "AJUST", "SOBRAS",
    )
    if any(term in s for term in blocked):
        return False

    entity_markers = (
        " LTDA", " EIRELI", " S/A", " SA", " ME", " EPP", " CONDOMINIO", " CLINICA",
        " HOSPITAL", " ASSOCIACAO", " FUNDACAO", " INSTITUTO", " SERVICOS", " COMERCIO",
        " INDUSTRIA", " PREFEITURA", " MUNICIPIO", " SESC", " SENAC", " SENAI",
    )
    if any(marker in " " + s for marker in entity_markers):
        return True

    # Nome de pessoa física: ao menos duas palavras alfabéticas substanciais.
    words = re.findall(r"\b[A-Z]{3,}\b", s)
    if len(words) >= 2 and len(s) <= 100:
        return True

    return False


def classify_credit_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """Acrescenta as duas colunas do Glossário Mestre v1 ao Consolidado.

    - Classifica somente créditos (Valor > 0).
    - Mantém débitos com as duas colunas vazias.
    - Primeira regra compatível vence (ordem de prioridade).
    - Crédito sem regra compatível fica em Revisão manual.
    """
    if df is None:
        return df

    out = df.copy()
    class_col = COL_CLASSIFICACAO
    foundation_col = COL_CATEGORIA
    out[class_col] = ""
    out[foundation_col] = ""

    if out.empty:
        return out

    values = pd.to_numeric(out.get("Valor"), errors="coerce")
    credit_mask = values.gt(0)
    if not bool(credit_mask.any()):
        return out

    normalized = out["Descrição"].fillna("").astype(str).map(normalize_classification_text)
    remaining = credit_mask.copy()

    for rule in sorted(CREDIT_CLASSIFICATION_RULES, key=lambda x: x["priority"]):
        if not bool(remaining.any()):
            break

        kind = rule.get("kind", "regex")
        if kind == "regex":
            pattern = rule.get("pattern") or r"(?!)"
            hit = remaining & normalized.str.contains(pattern, regex=True, na=False)
        elif kind == "numeric_artifact":
            hit = remaining & _mask_numeric_artifact(normalized)
        elif kind == "counterparty":
            # Aplica a heurística apenas às linhas ainda não classificadas.
            hit = pd.Series(False, index=out.index)
            idxs = out.index[remaining]
            hit.loc[idxs] = out.loc[idxs, "Descrição"].map(_looks_like_counterparty_text).fillna(False).astype(bool)
        else:
            continue

        if not bool(hit.any()):
            continue

        out.loc[hit, class_col] = rule["classification"]
        out.loc[hit, foundation_col] = rule["foundation"]
        remaining = remaining & ~hit

    # Todo crédito remanescente é explicitamente levado a revisão, sem inferência forçada.
    if bool(remaining.any()):
        out.loc[remaining, class_col] = CLASSIFICACAO_REVISAO
        out.loc[remaining, foundation_col] = f"Não classificado pelo glossário {GLOSSARIO_VERSION}"

    return out


# ---------------- Análise leve de cobertura/lacunas ----------------

BANK_DISPLAY_BY_LAYOUT = {
    "abc": "Banco ABC",
    "c6": "C6 Bank",
    "caixa_layout1": "Caixa",
    "caixa_layout2": "Caixa",
    "banrisul": "Banrisul",
    "credisis": "CrediSIS",
    "sicredi": "Sicredi",
    "sicredi_layout3": "Sicredi",
    "sicoob_layout1": "Sicoob",
    "arbi_layout1": "Arbi",
    "santander_layout4": "Santander",
    "bradesco_layout1": "Bradesco",
    "bradesco_layout2": "Bradesco",
    "bradesco_layout3": "Bradesco",
    "safra_layout2": "Safra",
    "santander": "Santander",
    "inter": "Inter",
    "itau": "Itaú",
    "itau_layout7": "Itaú",
    "unicred": "Unicred",
    "unicred_layout2": "Unicred",
    "efi": "EFI",
    "bb_layout1": "Banco do Brasil",
    "bb_layout2": "Banco do Brasil",
    "bb_layout3": "Banco do Brasil",
    "bb_layout4": "Banco do Brasil",
    "bb_layout5": "Banco do Brasil",
}

BANK_ID_DISPLAY = {
    "001": "Banco do Brasil",
    "033": "Santander",
    "041": "Banrisul",
    "077": "Inter",
    "104": "Caixa",
    "213": "Arbi",
    "237": "Bradesco",
    "246": "Banco ABC",
    "336": "C6 Bank",
    "341": "Itaú",
    "422": "Safra",
    "748": "Sicredi",
    "756": "Sicoob",
}

MONTH_ABBR_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
}

MONTH_NAME_NUM = {
    "jan": 1, "janeiro": 1,
    "fev": 2, "fevereiro": 2,
    "mar": 3, "marco": 3, "março": 3,
    "abr": 4, "abril": 4,
    "mai": 5, "maio": 5,
    "jun": 6, "junho": 6,
    "jul": 7, "julho": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "setembro": 9,
    "out": 10, "outubro": 10,
    "nov": 11, "novembro": 11,
    "dez": 12, "dezembro": 12,
}


def _bank_from_name(name: str) -> str:
    """Fallback conservador para arquivos sem layout útil/arquivos com erro."""
    low = norm_space(os.path.basename(str(name or ""))).lower()
    aliases = [
        (("banco do brasil", "bb_layout", "bb-", "bb_"), "Banco do Brasil"),
        (("banrisul",), "Banrisul"),
        (("bradesco",), "Bradesco"),
        (("c6",), "C6 Bank"),
        (("caixa",), "Caixa"),
        (("credisis", "credi sis"), "CrediSIS"),
        (("efi",), "EFI"),
        (("inter",), "Inter"),
        (("itaú", "itau"), "Itaú"),
        (("safra",), "Safra"),
        (("santander",), "Santander"),
        (("sicredi",), "Sicredi"),
        (("sicoob", "sicob"), "Sicoob"),
        (("unicred",), "Unicred"),
        (("abc",), "Banco ABC"),
        (("arbi", "arbi_layout"), "Arbi"),
    ]
    for needles, display in aliases:
        if any(x in low for x in needles):
            return display
    return "Banco não identificado"


def identify_bank_for_analysis(file_path: str, layout: str = "") -> str:
    """Identifica banco apenas para Log_Lacunas; não participa do dispatcher/parser."""
    layout = norm_space(layout).lower()
    if layout in BANK_DISPLAY_BY_LAYOUT:
        return BANK_DISPLAY_BY_LAYOUT[layout]
    if layout.startswith("bb_layout"):
        return "Banco do Brasil"

    if os.path.splitext(file_path)[1].lower() == ".ofx":
        try:
            # Leitura curta e independente, apenas para metadados do relatório.
            raw = open(file_path, "rb").read(32768)
            text = _decode_ofx_bytes(raw)
            low = fix_mojibake_ptbr(text).lower()
            textual = [
                (("banco do brasil",), "Banco do Brasil"),
                (("sicredi", "coop cred, poup e invest"), "Sicredi"),
                (("sicoob",), "Sicoob"),
                (("banrisul",), "Banrisul"),
                (("bradesco",), "Bradesco"),
                (("santander",), "Santander"),
                (("banco inter",), "Inter"),
                (("itaú", "itau"), "Itaú"),
                (("safra",), "Safra"),
                (("caixa econômica", "caixa economica"), "Caixa"),
                (("c6 bank",), "C6 Bank"),
                (("unicred",), "Unicred"),
            ]
            for needles, display in textual:
                if any(x in low for x in needles):
                    return display
            m = re.search(r"<BANKID>\s*([0-9]{3})", text, flags=re.I)
            if m and m.group(1) in BANK_ID_DISPLAY:
                return BANK_ID_DISPLAY[m.group(1)]
        except Exception:
            pass

    return _bank_from_name(file_path)


def _period_key(year: int, month: int):
    try:
        return pd.Period(year=int(year), month=int(month), freq="M")
    except Exception:
        return None


def _months_between(start_p, end_p):
    if start_p is None or end_p is None or end_p < start_p:
        return set()
    return set(pd.period_range(start=start_p, end=end_p, freq="M"))


def _months_from_filename(name: str):
    """Reconhece competências explícitas no nome; não presume cobertura por ano isolado."""
    low = unicodedata.normalize("NFC", os.path.basename(str(name or ""))).lower()
    found = set()

    month_pat = "|".join(sorted((re.escape(k) for k in MONTH_NAME_NUM), key=len, reverse=True))
    for m in re.finditer(rf"(?<![a-zà-ÿ])({month_pat})(?![a-zà-ÿ])\s*[-_./ ]*\s*(20\d{{2}})", low, flags=re.I):
        mon = MONTH_NAME_NUM.get(m.group(1).lower())
        p = _period_key(int(m.group(2)), mon)
        if p is not None:
            found.add(p)

    for m in re.finditer(r"(?<!\d)(0?[1-9]|1[0-2])\s*[-_./]\s*(20\d{2})(?!\d)", low):
        p = _period_key(int(m.group(2)), int(m.group(1)))
        if p is not None:
            found.add(p)

    for m in re.finditer(r"(?<!\d)(20\d{2})\s*[-_./]\s*(0?[1-9]|1[0-2])(?!\d)", low):
        p = _period_key(int(m.group(1)), int(m.group(2)))
        if p is not None:
            found.add(p)

    return found


def _months_from_ofx_declared_period(file_path: str):
    """Aproveita DTSTART/DTEND quando presentes; apenas metadado, sem interferir na extração."""
    if os.path.splitext(file_path)[1].lower() != ".ofx":
        return set()
    try:
        text = _decode_ofx_bytes(open(file_path, "rb").read(65536))
        vals = {}
        for tag in ("DTSTART", "DTEND"):
            m = re.search(rf"<{tag}>(.*?)(?:$|<|\r|\n)", text, flags=re.I | re.S)
            vals[tag] = normalize_ofx_date(m.group(1)) if m else ""
        if vals.get("DTSTART") and vals.get("DTEND"):
            a = datetime.strptime(vals["DTSTART"], "%d/%m/%Y")
            b = datetime.strptime(vals["DTEND"], "%d/%m/%Y")
            return _months_between(_period_key(a.year, a.month), _period_key(b.year, b.month))
    except Exception:
        pass
    return set()


def _fmt_period(p) -> str:
    if p is None:
        return ""
    return f"{MONTH_ABBR_PT[int(p.month)]}/{int(p.year)}"


def _fmt_periods(periods, max_items=12) -> str:
    vals = sorted(periods)
    if not vals:
        return ""
    labels = [_fmt_period(p) for p in vals]
    if len(labels) <= max_items:
        return ", ".join(labels)
    return ", ".join(labels[:max_items]) + f" (+{len(labels) - max_items})"


def build_log_lacunas(df_all: pd.DataFrame, file_records):
    """Gera análise pós-processamento usando os dados já extraídos.

    Retorna dois DataFrames para a aba Log_Lacunas:
      1) resumo de cobertura mensal por banco;
      2) intervalos atípicos entre datas com transações.
    """
    records = []
    tx_months_by_file = {}

    if df_all is not None and not df_all.empty:
        aux = df_all[["Arquivo", "Data"]].copy()
        aux["_dt"] = pd.to_datetime(aux["Data"], format="%d/%m/%Y", errors="coerce")
        aux = aux[aux["_dt"].notna()]
        for arquivo, g in aux.groupby("Arquivo"):
            tx_months_by_file[str(arquivo)] = set(g["_dt"].dt.to_period("M").tolist())

    for r in file_records:
        file_path = r.get("file_path", "")
        arquivo = r.get("Arquivo", os.path.basename(file_path))
        layout = r.get("Layout", "")
        banco = r.get("Banco") or identify_bank_for_analysis(file_path, layout)
        status = r.get("Status", "ok")
        file_months = set(_months_from_filename(arquivo))
        file_months |= _months_from_ofx_declared_period(file_path)
        file_months |= tx_months_by_file.get(str(arquivo), set())
        records.append({
            "Arquivo": arquivo,
            "Banco": banco,
            "Layout": layout,
            "Status": status,
            "CompetenciasArquivo": file_months,
            "CompetenciasTransacoes": tx_months_by_file.get(str(arquivo), set()),
        })

    summary_rows = []
    bancos = sorted({r["Banco"] for r in records})
    for banco in bancos:
        br = [r for r in records if r["Banco"] == banco]
        months_file = set().union(*(r["CompetenciasArquivo"] for r in br)) if br else set()
        months_tx = set().union(*(r["CompetenciasTransacoes"] for r in br)) if br else set()
        errors = sum(1 for r in br if r["Status"] == "erro")

        observed = months_file | months_tx
        if observed:
            pmin, pmax = min(observed), max(observed)
            expected = _months_between(pmin, pmax)
            window = f"{_fmt_period(pmin)} a {_fmt_period(pmax)}"
        else:
            expected = set()
            window = "não identificada"

        missing = expected - months_file

        conclusion_parts = []
        if errors:
            conclusion_parts.append(
                f"Há {errors} arquivo(s) sem transações extraídas; verificar se constam transações no(s) arquivo(s)."
            )
        if missing:
            conclusion_parts.append(f"Lacuna(s) interna(s) de competência: {_fmt_periods(missing)}.")
        elif expected and not errors:
            conclusion_parts.append("Cobertura mensal completa na janela observada. Não foi identificada lacuna de competência.")
        elif not expected and not errors:
            conclusion_parts.append("Não foi possível identificar janela mensal para este banco.")
        elif expected and errors:
            conclusion_parts.append("Não foi identificada lacuna mensal adicional com os metadados disponíveis.")

        summary_rows.append([
            banco,
            window,
            len(expected),
            len(months_file),
            len(months_tx),
            errors,
            " ".join(conclusion_parts),
        ])

    df_summary = pd.DataFrame(summary_rows, columns=[
        "Banco", "Janela observada", "Competências esperadas", "Competências com arquivo",
        "Meses com transações extraídas", "Erros no Log", "Conclusão"
    ])

    # Intervalos atípicos: critério P99 por banco.
    # Usa apenas as datas já extraídas no Consolidado, sem reler PDF/OFX.
    # Para cada banco:
    #   1) reúne as datas únicas de transação de todos os seus arquivos;
    #   2) calcula os gaps em dias corridos entre datas consecutivas;
    #   3) calcula o percentil 99 (P99) desses gaps;
    #   4) sinaliza somente gaps estritamente superiores ao P99 do banco.
    # O intervalo, isoladamente, não prova ausência de extrato.
    interval_rows = []
    if df_all is not None and not df_all.empty:
        aux = df_all[["Arquivo", "Data"]].copy()
        aux["_dt"] = pd.to_datetime(aux["Data"], format="%d/%m/%Y", errors="coerce")
        aux = aux[aux["_dt"].notna()]
        bank_by_file = {r["Arquivo"]: r["Banco"] for r in records}
        aux["Banco"] = aux["Arquivo"].astype(str).map(bank_by_file).fillna("Banco não identificado")

        for banco, gb in aux.groupby("Banco"):
            dates = sorted(set(gb["_dt"].dt.normalize().tolist()))
            if len(dates) < 2:
                continue

            diffs = [(dates[i] - dates[i - 1]).days for i in range(1, len(dates))]
            positive = [d for d in diffs if d > 0]
            if not positive:
                continue

            # pandas usa interpolação linear por padrão, compatível com o critério
            # adotado na análise anterior (ex.: P99 não precisa ser inteiro).
            p99 = float(pd.Series(positive, dtype="float64").quantile(0.99, interpolation="linear"))

            for i, delta in enumerate(diffs, start=1):
                if float(delta) <= p99:
                    continue

                dt_prev = pd.Timestamp(dates[i - 1])
                dt_next = pd.Timestamp(dates[i])
                # Dias úteis estritamente entre as duas datas, sem calendário de feriados.
                # Ex.: 16/10 -> 24/10 conta 17, 20, 21, 22 e 23 = 5 dias úteis.
                if dt_next - dt_prev <= pd.Timedelta(days=1):
                    business_days_without = 0
                else:
                    business_days_without = len(pd.bdate_range(
                        start=dt_prev + pd.Timedelta(days=1),
                        end=dt_next - pd.Timedelta(days=1)
                    ))

                interval_rows.append([
                    banco,
                    dt_prev.strftime("%d/%m/%Y"),
                    dt_next.strftime("%d/%m/%Y"),
                    int(delta),
                    int(business_days_without),
                    round(p99, 2),
                ])

    df_intervals = pd.DataFrame(interval_rows, columns=[
        "Banco", "Data anterior", "Data seguinte", "Intervalo (dias)",
        "Dias úteis sem tx", "P99 do banco"
    ])
    if not df_intervals.empty:
        # Facilita a inspeção: agrupa por banco e mostra primeiro os maiores gaps.
        df_intervals = df_intervals.sort_values(
            ["Banco", "Intervalo (dias)", "Data anterior"],
            ascending=[True, False, True],
            kind="stable"
        ).reset_index(drop=True)

    return df_summary, df_intervals


# ---------------- XLSX e fluxo ----------------

def export_xlsx(out_path, df_all, df_logs, df_lacunas=None, df_intervalos=None):
    engine = "xlsxwriter"
    try:
        __import__("xlsxwriter")
    except Exception:
        engine = "openpyxl"

    if df_lacunas is None:
        df_lacunas = pd.DataFrame(columns=[
            "Banco", "Janela observada", "Competências esperadas", "Competências com arquivo",
            "Meses com transações extraídas", "Erros no Log", "Conclusão"
        ])
    if df_intervalos is None:
        df_intervalos = pd.DataFrame(columns=[
            "Banco", "Data anterior", "Data seguinte", "Intervalo (dias)",
            "Dias úteis sem tx", "P99 do banco"
        ])

    summary_startrow = 1
    interval_title_row = summary_startrow + len(df_lacunas) + 2
    interval_startrow = interval_title_row + 1

    with pd.ExcelWriter(out_path, engine=engine) as writer:
        df_all.to_excel(writer, index=False, sheet_name="Consolidado")
        df_logs.to_excel(writer, index=False, sheet_name="Logs")
        df_lacunas.to_excel(writer, index=False, sheet_name="Log_Lacunas", startrow=summary_startrow)
        df_intervalos.to_excel(writer, index=False, sheet_name="Log_Lacunas", startrow=interval_startrow)

        if engine == "xlsxwriter":
            wb = writer.book
            ws = writer.sheets["Consolidado"]
            ws_logs = writer.sheets["Logs"]
            ws_lac = writer.sheets["Log_Lacunas"]

            money_fmt = wb.add_format({"num_format": "R$ #,##0.00;[Red]-R$ #,##0.00"})
            date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})
            title_fmt = wb.add_format({"bold": True, "font_size": 12})
            header_fmt = wb.add_format({"bold": True, "text_wrap": True, "valign": "top", "border": 1})
            wrap_fmt = wb.add_format({"text_wrap": True, "valign": "top"})

            ws.set_column("A:A", 35)
            ws.set_column("B:B", 12, date_fmt)
            ws.set_column("C:C", 110)
            ws.set_column("D:D", 22)
            ws.set_column("E:E", 16, money_fmt)
            ws.set_column("F:F", 6)
            ws.set_column("G:H", 16, money_fmt)
            ws.set_column("I:I", 42, wrap_fmt)
            ws.set_column("J:J", 48, wrap_fmt)

            ws_logs.set_column("A:A", 35)
            ws_logs.set_column("B:B", 20)

            ws_lac.write(0, 0, "1. RESUMO POR BANCO", title_fmt)
            ws_lac.write(interval_title_row, 0, "2. Critério: intervalos superiores ao percentil 99 (P99) dos gaps entre datas únicas de transação de cada banco. O intervalo, isoladamente, não prova ausência de extrato.", title_fmt)
            ws_lac.set_row(summary_startrow, 32, header_fmt)
            ws_lac.set_row(interval_startrow, 32, header_fmt)
            ws_lac.set_column("A:A", 22)
            ws_lac.set_column("B:B", 36)
            ws_lac.set_column("C:F", 22)
            ws_lac.set_column("G:G", 85, wrap_fmt)
            ws_lac.freeze_panes(2, 0)

        else:
            ws = writer.sheets["Consolidado"]
            ws_logs = writer.sheets["Logs"]
            ws_lac = writer.sheets["Log_Lacunas"]
            widths = {"A": 35, "B": 12, "C": 110, "D": 22, "E": 16, "F": 6, "G": 16, "H": 16, "I": 42, "J": 48}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
            ws_logs.column_dimensions["A"].width = 35
            ws_logs.column_dimensions["B"].width = 20
            for row in ws.iter_rows(min_row=2):
                row[1].number_format = 'DD/MM/YYYY'
                for idx in (4, 6, 7):
                    row[idx].number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

            ws_lac.cell(row=1, column=1, value="1. RESUMO POR BANCO")
            ws_lac.cell(row=interval_title_row + 1, column=1, value="2. Critério: intervalos superiores ao percentil 99 (P99) dos gaps entre datas únicas de transação de cada banco. O intervalo, isoladamente, não prova ausência de extrato.")
            ws_lac.column_dimensions["A"].width = 22
            ws_lac.column_dimensions["B"].width = 36
            for col in ("C", "D", "E", "F"):
                ws_lac.column_dimensions[col].width = 22
            ws_lac.column_dimensions["G"].width = 85
            ws_lac.freeze_panes = "A3"

def escolher_pasta():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pasta = filedialog.askdirectory(title="Selecione a pasta com PDFs/OFX")
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
    file_records = []
    total = len(arquivos)

    for idx, file_path in enumerate(arquivos, start=1):
        nome = norm_space(os.path.basename(file_path))
        layout = ""
        try:
            layout, df = parse_one_file(file_path)
            banco = identify_bank_for_analysis(file_path, layout)
            if df.empty:
                logs.append([nome, "erro"])
                file_records.append({
                    "file_path": file_path, "Arquivo": nome, "Layout": layout,
                    "Banco": banco, "Status": "erro"
                })
                print(f"[{idx}/{total}] ERRO - {nome} | sem transações extraídas")
                continue

            df.insert(0, "Arquivo", nome)
            dados.append(df)
            logs.append([nome, int(len(df))])
            file_records.append({
                "file_path": file_path, "Arquivo": nome, "Layout": layout,
                "Banco": banco, "Status": "ok"
            })
            print(f"[{idx}/{total}] OK   - {nome} | {layout} | {len(df)} transação(ões)")

        except Exception as e:
            banco = identify_bank_for_analysis(file_path, layout)
            logs.append([nome, "erro"])
            file_records.append({
                "file_path": file_path, "Arquivo": nome, "Layout": layout,
                "Banco": banco, "Status": "erro"
            })
            print(f"[{idx}/{total}] ERRO - {nome} | {e}")

    base_columns = ["Arquivo", "Data", "Descrição", "Documento", "Valor", "Tipo", "Débito", "Crédito"]
    if dados:
        df_all = pd.concat(dados, ignore_index=True)
        df_all = df_all[base_columns]
    else:
        df_all = pd.DataFrame(columns=base_columns)

    # Classificação pós-processamento. Não interfere nos parsers nem nas 8 colunas-base.
    df_all = classify_credit_transactions(df_all)
    df_all = df_all[base_columns + [COL_CLASSIFICACAO, COL_CATEGORIA]]

    df_logs = pd.DataFrame(logs, columns=["Arquivo", "n_transações_obtidas"])
    df_lacunas, df_intervalos = build_log_lacunas(df_all, file_records)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(folder, f"consolidado_lancamentos_{stamp}.xlsx")
    export_xlsx(out_path, df_all, df_logs, df_lacunas, df_intervalos)
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
        if COL_CLASSIFICACAO in df_all.columns:
            creditos = pd.to_numeric(df_all["Valor"], errors="coerce").gt(0)
            revisao = creditos & df_all[COL_CLASSIFICACAO].eq(CLASSIFICACAO_REVISAO)
            print(f"Glossário de créditos: {GLOSSARIO_VERSION}")
            print(f"Créditos classificados: {int(creditos.sum() - revisao.sum())} | revisão manual: {int(revisao.sum())}")
    except Exception as e:
        print("\nERRO GERAL:")
        print(str(e))


if __name__ == "__main__":
    main()
