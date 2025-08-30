# app.py
"""
Script per convertire report .txt (formato Volvo-like) in PDF con:
- colonne: NUMERO GARANZIA, SUFFISSO, JOB, TOTALE JOB
- rimozione duplicati su (NUMERO GARANZIA, SUFFISSO, JOB)
- ordinamento
- totali (Totale, IVA 22%, Totale IVA incl.)
Usa: trascina il .txt sull'exe (Windows) - l'exe legge i file passati come argomenti.
"""

import sys
import os
import re
from decimal import Decimal, ROUND_HALF_UP

try:
    import pandas as pd
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except Exception as e:
    print("ERRORE: moduli mancanti:", e)
    raise

ROW_RE = re.compile(r"^\s*(\d{7})\s+(\d+)\s+(\d+)\s+(-?\d+)\b", re.MULTILINE)

HEADER = ["NUMERO GARANZIA", "SUFFISSO", "JOB", "TOTALE JOB"]
DISCLAIMER = (
    "Disclaimer: I totali riportati nel presente documento sono stati calcolati automaticamente.\n"
    "A causa di possibili arrotondamenti e differenze di calcolo, potrebbero verificarsi scostamenti minimi di qualche euro rispetto ai valori ufficiali di fatturazione."
)

def parse_txt(path: str) -> pd.DataFrame:
    with open(path, "r", encoding="latin-1", errors="ignore") as f:
        text = f.read()
    matches = ROW_RE.findall(text)
    if not matches:
        raise ValueError("Nessuna riga valida trovata nel TXT (controlla il formato).")
    df = pd.DataFrame(matches, columns=HEADER)
    df["NUMERO GARANZIA"] = df["NUMERO GARANZIA"].astype(str)
    df["SUFFISSO"] = df["SUFFISSO"].astype(str)
    df["JOB"] = pd.to_numeric(df["JOB"], errors="coerce").fillna(0).astype(int)
    df["TOTALE JOB"] = pd.to_numeric(df["TOTALE JOB"], errors="coerce").fillna(0).astype(int)
    df = df.drop_duplicates(subset=["NUMERO GARANZIA", "SUFFISSO", "JOB"], keep="first")
    df = df.sort_values(by=["NUMERO GARANZIA", "SUFFISSO", "JOB"],
                        key=lambda c: pd.to_numeric(c, errors="ignore")).reset_index(drop=True)
    return df

def eur_fmt(val: Decimal) -> str:
    q = val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{q:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"

def build_pdf(df: pd.DataFrame, out_pdf: str):
    totale = Decimal(int(df["TOTALE JOB"].astype(int).sum()))
    iva = (totale * Decimal("0.22")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    ivato = (totale + iva).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    doc = SimpleDocTemplate(out_pdf, pagesize=A4, leftMargin=28, rightMargin=28, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    title = Paragraph("<b>Riepilogo Garanzie</b>", styles["Heading3"])

    data = [HEADER] + df.values.tolist()
    width = doc.width
    col_widths = [0.28 * width, 0.32 * width, 0.12 * width, 0.28 * width]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTSIZE", (0,0), (-1,0), 9),
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("ALIGN", (2,1), (2,-1), "CENTER"),
        ("ALIGN", (3,1), (3,-1), "RIGHT"),
        ("FONTSIZE", (0,1), (-1,-1), 8),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))

    totals = [
        ["Totale", f"{int(totale)}"],
        ["IVA 22%", eur_fmt(iva)],
        ["Totale IVA inclusa", eur_fmt(ivato)],
    ]
    t_table = Table(totals, colWidths=[0.55 * width, 0.45 * width])
    t_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("ALIGN", (0,0), (0,-1), "RIGHT"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
    ]))

    disclaimer = Paragraph("<font size=7><i>" + DISCLAIMER.replace("\n", "<br/>") + "</i></font>", styles["Normal"])
    story = [title, Spacer(1, 8), table, Spacer(1, 10), t_table, Spacer(1, 8), disclaimer]
    doc.build(story)

def output_path_for(txt_path: str) -> str:
    base, _ = os.path.splitext(os.path.basename(txt_path))
    directory = os.path.dirname(os.path.abspath(txt_path)) or os.getcwd()
    return os.path.join(directory, f"{base}_Riepilogo_Garanzie.pdf")

def process_file(txt_path: str):
    print(f"Elaborazione: {txt_path}")
    df = parse_txt(txt_path)
    print(f"Righe estratte: {len(df)}")
    out_pdf = output_path_for(txt_path)
    build_pdf(df, out_pdf)
    print(f"PDF generato: {out_pdf}")
    return out_pdf

def main():
    files = [p for p in sys.argv[1:] if os.path.isfile(p)]
    if not files:
        print("Nessun file passato. Trascina il file .txt sull'exe oppure esegui: RiepilogoGaranzie.exe percorso\\file.txt")
        return 0
    created = []
    for f in files:
        try:
            created.append(process_file(f))
        except Exception as e:
            print(f"ERRORE durante l'elaborazione di {f}: {e}")
            raise
    print("Elaborazione completata.")
    for c in created:
        print("CREATO:", c)
    return 0

if __name__ == "__main__":
    sys.exit(main())
