import pandas as pd
import glob
import fpdf
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    pdf = fpdf.FPDF(orientation="P",unit="mm",format="A4")
    filename = Path(filepath).stem
    invoice_nr,date = filename.split("-")[0]
    pdf.add_page()
    pdf.set_font(family="Arial",size=16,style="B")
    pdf.cell(w=0,h=8,txt=f"Invoice nr.{invoice_nr}",ln=1)
    pdf.set_font(family="Arial",size=16,style="B")
    pdf.cell(w=0,h=8,txt=f"Date:{date}")
    pdf.output(f"{filename}.pdf")
    

