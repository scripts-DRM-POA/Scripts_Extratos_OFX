import os, glob, re
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd

from tkinter import Tk, filedialog


def selecionar_pasta_pdfs() -> str:
    """Abre uma janela para o usuário selecionar a pasta onde estão os PDFs."""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    pasta = filedialog.askdirectory(title="Selecione a pasta onde estão os PDFs do extrato")
    root.destroy()

    if not pasta:
        raise SystemExit("Nenhuma pasta selecionada. Execução cancelada.")
    return pasta


def extract_text_lines(pdf_path: str):
    """Extrai texto do PDF em linhas (PyMuPDF get_text('text'))."""
    doc = fitz.open(pdf_path)
    lines = []
    for pg in doc:
        txt = pg.get_text("text")
        page_lines = [ln.replace("\xa0", " ").strip() for ln in txt.split("\n")]
        lines.extend([ln for ln in page_lines if ln])
    return lines


# --- REGEX ---
RE_MONEY = re.compile(r"(?<!\d)(-?\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?!\d)")
RE_MONEY_TRAIL_MINUS = re.compile(r"(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})-(?!\d)")


def is_money_line(ln: str) -> bool:
    return bool(RE_MONEY.fullmatch(ln) or RE_MONEY_TRAIL_MINUS.fullmatch(ln))


def money_to_float_br(token: str) -> float:
    """
    Converte:
      - "-449,84" -> -449.84
      - "73,26-"  -> -73.26
      - "1.690,37"-> 1690.37
    """
    t = token.strip()
    neg = False
    if t.endswith("-"):
        neg = True
        t = t[:-1]
    if t.startswith("-"):
        neg = True
        t = t[1:]
    val = float(t.replace(".", "").replace(",", "."))
    return -val if neg else val


# 1) linhas de SALDO / POSIÇÃO / TOTALIZADORES (não são transação)
#    IMPORTANTE: NÃO filtrar "aplic aut mais" sozinho, para não perder "Rend Pago Aplic Aut Mais".
SKIP_DESC_RE = re.compile(
    r"(?:^|\b)("
    r"saldo\s+a\s+liberar|"
    r"saldo\s+anterior|"
    r"saldo\s+final(?:\s+(?:devedor|dispon[ií]vel))?|"
    r"saldo\s+total|"
    r"saldo\s+dispon[ií]vel|"
    r"saldo\s+total\s+dispon[ií]vel\s+dia|"
    r"saldo\s+aplic(?:\s+aut)?(?:\s+mais)?|"
    r"saldo\s+em\s+aplica(ç|c)[aã]o\s+autom[aá]tica|"
    r"saldo\s+dispon[ií]vel\s+sem\s+investimentos\s+autom[aá]ticos|"
    r"valor\s+total\s+em\s+aplica(ç|c)[aã]o(?:es)?\s+autom[aá]tica(?:s)?|"
    r"rendimentos\s+de\s+aplica(ç|c)[aã]o(?:es)?\s+autom[aá]tica(?:s)?|"
    r"saldo\s+da\s+conta\s+corrente|"
    r"limite\s+da\s+conta|"
    r"total\s+dispon[ií]vel(?:\s+para\s+uso)?|"
    r"totalizador|"
    r"resumo\s*-\s*m[eê]s|"
    r"utilizado|"
    r"dispon[ií]vel|"
    r"total\s+de\s+cr[eé]ditos|"
    r"total\s+de\s+d[eé]bitos|"
    # quadros-resumo do extrato mensal (não são lançamentos)
    r"entrada\s+R\$\s*sa[ií]da\s+R\$\s+na\s+conta\s+corrente|"
    r"rendimentos\s+resgates\s+antec|"
    r"valor\s+da\s+rendimento(?:s)?\b|"
    r"vencimentos\b|"
    r"principal\b"
    r")\b",
    re.IGNORECASE
)

# 2) “totais” que aparecem como linhas e acabam virando descrição (não são transação)
SKIP_TOTAL_LINE_RE = re.compile(
    r"^(?:total|bruto|l[ií]quido)\b|"
    r"^total\s+R\$\s*|"
    r"^TOTAL\s+DE\s+",
    re.IGNORECASE
)

# 3) Ruído/cabeçalho/rodapé (não é descrição de lançamento)
IGNORE_TEXT_LINES = re.compile(
    r"^(Agência|Conta|CNPJ|Saldo|Limite|Utilizado|Disponível|Lançamentos do período|Data$|Lançamentos$|Razão Social|"
    r"CNPJ/CPF|Valor\s*\(R\$\)|Saldo\s*\(R\$\)|extrato mensal|CTCE|ESC|Av |entradas|saídas|total entradas|total saídas|"
    r"\(créditos\)|\(débitos\)|A =|B =|C =|D =|G =|P =|Para demais|Este material|491029\b|"
    r"/9135|REM-C|Minha conta|Minha agência|saldo em|01\. Conta Corrente)",
    re.IGNORECASE
)


def detect_year(lines):
    """
    Infere o ano quando o PDF usa dd/mm (sem ano).
    PRIORIDADE: usar o ano do período "Extrato de dd/mm/AAAA até dd/mm/AAAA".
    """
    sample = " ".join(lines[:400]).lower()

    # 1) Prioridade máxima: período do extrato
    m = re.search(
        r"extrato\s+de\s+\d{2}/\d{2}/(20\d{2})\s+at[eé]\s+\d{2}/\d{2}/(20\d{2})",
        sample
    )
    if m:
        y1, y2 = int(m.group(1)), int(m.group(2))
        return y1 if y1 == y2 else y2

    # 2) Outras heurísticas
    m = re.search(r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\b\D{0,12}\b(20\d{2})\b", sample)
    if m:
        return int(m.group(2))

    m = re.search(r"\b(0?[1-9]|1[0-2])/(20\d{2})\b", sample)
    if m:
        return int(m.group(2))

    m = re.search(r"\b(20\d{2})\b", sample)
    if m:
        return int(m.group(1))

    return None


def parse_multi_layout_column_broken(lines):
    """
    Parser “stream” para PDFs com colunas quebradas.

    Ajustes pedidos:
      - NÃO remover "data grudada" na descrição (mantém como está no PDF).
      - NÃO interromper ao sair de "Conta Corrente | Movimentação": assim mantém Aplicações/Resgates/Rendimentos etc.
    """
    year_ctx = detect_year(lines) or datetime.now().year

    current_date = None
    pending_desc = None
    rows = []

    # detectar se é extrato mensal
    is_extrato_mensal = any(re.search(r"\bextrato\s+mensal\b", ln, re.IGNORECASE) for ln in lines[:150])
    in_mov_section = False

    START_MOV_RE = re.compile(r"Conta\s+Corrente\s*\|\s*Movimenta(ç|c)[aã]o", re.IGNORECASE)

    # gatilho “resumo/saldos” (layout 2023) — só usar quando NÃO for extrato mensal
    STOP_SECTION_RESUMO_2023 = re.compile(
        r"^Saldo\s+da\s+conta\s+corrente\b|"
        r"^(?:Descri(ç|c)[aã]o)\s*$|"
        r"^(?:Valor\s*\(R\$\))\s*$|"
        r"^(?:Saldo\s*\(R\$\))\s*$|"
        r"saldo\s+dispon[ií]vel\s+sem\s+investimentos\s+autom[aá]ticos|"
        r"total\s+dispon[ií]vel\s+para\s+uso",
        re.IGNORECASE
    )

    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue

        # Para layout 2023: encerra ao entrar no quadro “Saldo da conta corrente”
        if (not is_extrato_mensal) and rows and STOP_SECTION_RESUMO_2023.search(ln):
            break

        # Controle de seção: extrato mensal
        if is_extrato_mensal:
            # entra na seção de movimentação quando achar o cabeçalho;
            # a partir daí, mantém leitura para capturar também outras seções (aplicações/resgates etc.)
            if START_MOV_RE.search(ln):
                in_mov_section = True
                pending_desc = None
                current_date = None
                continue

            if not in_mov_section:
                continue

        # DATA com ano
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", ln):
            current_date = datetime.strptime(ln, "%d/%m/%Y")
            pending_desc = None
            continue

        # DATA sem ano (extrato mensal)
        if re.fullmatch(r"\d{2}/\d{2}", ln):
            current_date = datetime.strptime(f"{ln}/{year_ctx}", "%d/%m/%Y")
            pending_desc = None
            continue

        if current_date is None:
            continue

        # Se a linha inteira já é de saldo/totalizador, mata pendência e segue
        if SKIP_DESC_RE.search(ln) or SKIP_TOTAL_LINE_RE.search(ln):
            pending_desc = None
            continue

        # VALOR
        if is_money_line(ln):
            if not pending_desc:
                continue

            if SKIP_DESC_RE.search(pending_desc) or SKIP_TOTAL_LINE_RE.search(pending_desc):
                pending_desc = None
                continue

            val = money_to_float_br(ln)
            tipo = "C" if val > 0 else "D"
            rows.append([
                current_date.strftime("%d/%m/%Y"),
                current_date.year,
                current_date.month,
                pending_desc,  # mantém "data grudada" se vier no PDF
                val,
                tipo
            ])
            pending_desc = None
            continue

        # DESCRIÇÃO
        if IGNORE_TEXT_LINES.search(ln):
            continue

        ln2 = re.sub(r"^(D|C)\s+", "", ln, flags=re.IGNORECASE).strip()

        # ignora “linhas numéricas”
        if re.fullmatch(r"[\d\.\-%/\\]+", ln2):
            continue

        # ignora linhas de "total/bruto/líquido"
        if SKIP_TOTAL_LINE_RE.search(ln2):
            pending_desc = None
            continue

        # ignora se já “parece saldo/posição/resumo”
        if SKIP_DESC_RE.search(ln2):
            pending_desc = None
            continue

        # acumula descrição (multiline)
        if pending_desc:
            pending_desc = (pending_desc + " " + ln2).strip()
        else:
            pending_desc = ln2

    return pd.DataFrame(rows, columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])


def listar_pdfs(pasta):
    pdfs = glob.glob(os.path.join(pasta, "*.pdf"))
    return sorted({os.path.abspath(p) for p in pdfs})


def processar_pasta_multi(pasta_pdf: str, saida_xlsx: str = None):
    pdfs = listar_pdfs(pasta_pdf)
    if not pdfs:
        raise FileNotFoundError(f"Não encontrei PDFs em: {pasta_pdf}")

    print(f"Processando {len(pdfs)} PDF(s) em: {pasta_pdf}")

    dfs = []
    erros = []
    vazios = []

    for idx, pdf_path in enumerate(pdfs, start=1):
        nome = os.path.basename(pdf_path)
        try:
            print(f"[{idx}/{len(pdfs)}] Lendo: {nome}")
            lines = extract_text_lines(pdf_path)
            df = parse_multi_layout_column_broken(lines)

            if df.empty:
                vazios.append(nome)
                print("    -> AVISO: nenhum lançamento extraído")
                continue

            df.insert(0, "Arquivo", nome)
            dfs.append(df)
            print(f"    -> OK: {len(df)} lançamento(s)")

        except Exception as e:
            erros.append((nome, str(e)))
            print(f"    -> ERRO: {e}")

    if not dfs:
        msg = "Nenhum lançamento foi extraído.\n"
        if vazios:
            msg += f"- PDFs lidos porém vazios: {len(vazios)}\n"
        if erros:
            msg += f"- PDFs com erro: {len(erros)}\n"
        raise ValueError(msg)

    df_all = pd.concat(dfs, ignore_index=True)

    if saida_xlsx is None:
        saida_xlsx = os.path.join(pasta_pdf, "LANCAMENTOS_CONSOLIDADO_MULTI_LAYOUT.xlsx")

    print(f"Gerando Excel: {saida_xlsx}")

    try:
        writer = pd.ExcelWriter(saida_xlsx, engine="xlsxwriter")
        engine = "xlsxwriter"
    except ModuleNotFoundError:
        writer = pd.ExcelWriter(saida_xlsx, engine="openpyxl")
        engine = "openpyxl"

    with writer:
        df_all.to_excel(writer, index=False, sheet_name="Lancamentos")

        if vazios:
            pd.DataFrame(vazios, columns=["Arquivo"]).to_excel(writer, index=False, sheet_name="Vazios")

        if erros:
            pd.DataFrame(erros, columns=["Arquivo", "Erro"]).to_excel(writer, index=False, sheet_name="Erros")

        if engine == "xlsxwriter":
            wb = writer.book
            ws = writer.sheets["Lancamentos"]
            money_fmt = wb.add_format({"num_format": "R$ #,##0.00; -R$ #,##0.00"})
            date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

            ws.set_column("A:A", 40)            # Arquivo
            ws.set_column("B:B", 12, date_fmt)  # Data
            ws.set_column("C:C", 6)             # Ano
            ws.set_column("D:D", 5)             # Mês
            ws.set_column("E:E", 100)           # Descrição
            ws.set_column("F:F", 16, money_fmt) # Valor
            ws.set_column("G:G", 5)             # Tipo

    print("Concluído.")
    print(f"Engine usada: {engine}")
    if vazios:
        print(f"Atenção: {len(vazios)} PDF(s) sem lançamentos (ver aba 'Vazios').")
    if erros:
        print(f"Atenção: {len(erros)} PDF(s) com erro (ver aba 'Erros').")

    return df_all, saida_xlsx


if __name__ == "__main__":
    pasta = selecionar_pasta_pdfs()
    df_final, arquivo_saida = processar_pasta_multi(pasta)
    print(arquivo_saida)
