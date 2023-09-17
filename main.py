import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import time

filepaths = glob.glob("input/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    data = pd.read_excel(filepath, sheet_name="Sheet1")
    pdf.add_page()

    #Add invoice name
    filename = Path(filepath).stem[-1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt="Invoice nr."+filename, ln=1)

    #Add date and time
    pdf.set_font(family="Times", size=12)
    pdf.cell(w=50, h=8, txt="Invoice date: "+time.strftime("%d.%m.%Y %H:%M:%S"), ln=1)

    # Add a header
    columns = list(data.columns)
    columns = [x.replace("_", " ").title() for x in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    #Add rows to the table
    for index, row in data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(format(row["total_price"], ".2f")), border=1, ln=1)


    pdf.output(f"PDFs/invoice{filename}.pdf")
