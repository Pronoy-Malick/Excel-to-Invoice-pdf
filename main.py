import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    pdf.set_font(family="Helvetica", style="B", size=18)
    pdf.cell(w=5, h=10,txt=f"Invoice No:- {invoice_no}", ln=1)
    pdf.cell(w=3, h=10,txt=f"Date:- {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("-", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100,0 ,100)
    pdf.cell(w=30, h=10, txt=columns[0], border=1)
    pdf.cell(w=50, h=10, txt=columns[1], border=1)
    pdf.cell(w=35, h=10, txt=columns[2], border=1)
    pdf.cell(w=30, h=10, txt=columns[3], border=1)
    pdf.cell(w=30, h=10, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(0,0,255)
        pdf.cell(w=30, h=10, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=10, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), border=1, ln=1)

    total_price = df["total_price"].sum()
    pdf.set_font(family="Times", size=10 )
    pdf.set_text_color(0, 0, 255)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=50, h=10, txt="", border=1)
    pdf.cell(w=35, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=str(total_price), border=1, ln=1)

    # Total price sentence
    pdf.set_font(family="Times", size=10, style="B" )
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=10, txt=f"the total price of the item {total_price}", ln=1)

    # Company Name
    pdf.set_font(family="Times", size=15, style="B" )
    pdf.set_text_color(100, 200, 200)
    pdf.cell(w=30, h=10, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDF/{filename}.pdf")