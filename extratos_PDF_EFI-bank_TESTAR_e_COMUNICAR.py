# -*- coding: utf-8 -*-

import re
import os
from datetime import datetime
from tkinter import Tk, filedialog

import fitz  # PyMuPDF
import pandas as pd


DATE_RE = re.compile(r"^(\d{2}/\d{2}/\d{4})")
MONEY_TOKEN = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")
ONLY_DIGITS = re.compile(r"^\d{6,}$")
HEADER_HINTS = ("Lançamentos", "Data Descrição", "Protocolo", "Valor (R$)", "Saldo (R$)")


def rgb_from_int(color_int: int):
    r = (color_int >> 16) & 255
    g = (color_int >> 8) & 255
    b = color_int & 255
    return r, g, b


def classify_color(color_int: int):
    """Verde forte -> C; Vermelho forte -> D; neutro -> None."""
    r, g, b = rgb_from_int(color_int)
    if g - r > 40:
        return "C"
    if r - g > 40:
        return "D"
    return None


def norm_space(s: str) -> str:
    return re.sub(r"\s{2,}", " ", s.replace("\xa0", " ").strip())


def parse_pdf(path: str) -> pd.DataFrame:
    doc = fitz.open(path)

    rows = []
    current_date = None
    desc_buf = []

    for p in range(len(doc)):
        d = doc.load_page(p).get_text("dict")

        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                line_text = norm_space(" ".join(s["text"] for s in spans))

                # 1) Nova data?
                mdate = DATE_RE.match(line_text)
                if mdate:
                    dt = datetime.strptime(mdate.group(1), "%d/%m/%Y")
                    current_date = dt

                    if "Saldo do dia" in line_text:
                        desc_buf = []
                        continue

                    # Parte após a data: preparar descrição inicial
                    after = line_text[len(mdate.group(0)):].strip(" -")

                    # remove tokens de protocolo e números monetários
                    after = " ".join(tok for tok in after.split() if not ONLY_DIGITS.match(tok))
                    after = re.sub(r"\d{1,3}(?:\.\d{3})*,\d{2}", "", after).strip(" -")
                    desc_buf = [after] if after else []

                    # Verifica se há valor já nesta linha (olhando cor)
                    for s in spans:
                        t = s["text"].strip()
                        if MONEY_TOKEN.match(t):
                            tipo = classify_color(s["color"])
                            if not tipo:
                                continue  # saldo/neutro
                            val = float(t.replace(".", "").replace(",", "."))
                            if tipo == "D":
                                val = -val
                            desc = norm_space(" ".join(desc_buf))
                            if desc:
                                rows.append([
                                    dt.strftime("%d/%m/%Y"),
                                    dt.year,
                                    dt.month,
                                    desc,
                                    val,
                                    "C" if val > 0 else "D"
                                ])
                                desc_buf = []
                            break
                    continue

                # 2) Sem data: pode ser descrição adicional ou conter o valor
                # Ignora cabeçalhos / meta
                if any(h in line_text for h in HEADER_HINTS):
                    continue

                # Tenta localizar valor nesta linha (com cor de crédito/débito)
                found_value = False
                for s in spans:
                    t = s["text"].strip()
                    if MONEY_TOKEN.match(t):
                        tipo = classify_color(s["color"])
                        if not tipo:
                            continue  # saldo/neutro
                        val = float(t.replace(".", "").replace(",", "."))
                        if tipo == "D":
                            val = -val

                        # Descrição = buffer + textos anteriores ao valor nesta linha (tirando protocolos e valores)
                        before_parts = []
                        for s2 in spans:
                            if s2 is s:
                                break
                            t2 = s2["text"].strip()
                            if not t2 or ONLY_DIGITS.match(t2) or MONEY_TOKEN.match(t2):
                                continue
                            before_parts.append(t2)
                        if before_parts:
                            desc_buf.append(" ".join(before_parts))

                        desc = norm_space(" ".join(desc_buf))
                        if current_date and desc:
                            rows.append([
                                current_date.strftime("%d/%m/%Y"),
                                current_date.year,
                                current_date.month,
                                desc,
                                val,
                                "C" if val > 0 else "D"
                            ])
                        desc_buf = []
                        found_value = True
                        break

                if found_value:
                    continue

                # 3) Acumula descrição quando houver letras e não for meta
                parts = []
                for s in spans:
                    t = s["text"].strip()
                    if not t or ONLY_DIGITS.match(t) or MONEY_TOKEN.match(t):
                        continue
                    parts.append(t)
                part_txt = norm_space(" ".join(parts))
                if part_txt and current_date:
                    desc_buf.append(part_txt)

    df = pd.DataFrame(rows, columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])
    if not df.empty:
        df["Descrição da operação"] = df["Descrição da operação"].str.replace(r"\s{2,}", " ", regex=True).str.strip(" -")
        df["__dt"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
        df["__seq"] = range(len(df))
        df.sort_values(["__dt", "__seq"], inplace=True)
        df.drop(columns=["__dt", "__seq"], inplace=True)
    return df


def main():
    # Selecionar pasta via janela
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

    if not folder:
        raise SystemExit("Nenhuma pasta selecionada.")

    pdf_paths = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    ]

    if not pdf_paths:
        raise SystemExit("Nenhum PDF encontrado na pasta.")

    frames = []
    for path in sorted(pdf_paths):
        df = parse_pdf(path)
        if not df.empty:
            frames.append(df)
            print(f"OK: {os.path.basename(path)} -> {len(df)} linhas")
        else:
            print(f"Aviso: sem lançamentos em {os.path.basename(path)}")

    result = (
        pd.concat(frames, ignore_index=True)
        if frames
        else pd.DataFrame(columns=["Data", "Ano", "Mês", "Descrição da operação", "Valor", "Tipo"])
    )

    # Nome genérico + timestamp para não sobrescrever
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = os.path.join(folder, f"consolidado_lancamentos_{timestamp}.xlsx")

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result.to_excel(writer, index=False, sheet_name="Lancamentos")
        wb = writer.book
        ws = writer.sheets["Lancamentos"]

        money_fmt = wb.add_format({"num_format": "R$ #,##0.00; -R$ #,##0.00"})
        date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

        ws.set_column("A:A", 12, date_fmt)
        ws.set_column("B:B", 6)
        ws.set_column("C:C", 5)
        ws.set_column("D:D", 120)
        ws.set_column("E:E", 16, money_fmt)
        ws.set_column("F:F", 5)

    print("Concluído:", output)


if __name__ == "__main__":
    main()