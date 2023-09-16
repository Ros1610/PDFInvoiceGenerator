import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("input/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet1")
    print(df)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    filename = Path(filepath).stem[-1]
    pdf.cell(w=50, h=8, txt="Invoice nr."+filename)
    pdf.output(f"PDFs/invoice{filename}.pdf")
