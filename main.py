import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    pdf.set_font(family="Helvetica", style="B", size=18)
    pdf.cell(w=5, h=10,txt=f"Invoice No:- {invoice_no}", ln=1)
    pdf.cell(w=3, h=10,txt=f"Date:- {date}")


    pdf.output(f"PDF/{filename}.pdf")