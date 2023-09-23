import pandas as pd
import glob
import fpdf
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    
    pdf = fpdf.FPDF(orientation="P",unit="mm",format="A4")
    filename = Path(filepath).stem
    invoice_nr,date = filename.split("-")
    pdf.add_page()
    pdf.set_font(family="Arial",size=16,style="B")
    pdf.cell(w=0,h=8,txt=f"Invoice nr.{invoice_nr}",ln=1)
    pdf.set_font(family="Arial",size=16,style="B")
    pdf.cell(w=0,h=8,txt=f"Date:{date}",ln=1)
   
    
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    
    columns = df.columns
    columns = [item.replace("_"," ").title() for item in columns]
    # headera
    pdf.set_font(family="Times",size=12,style="B")
    pdf.cell(w=30,h=10,txt=columns[0],ln=0,border=1)
    pdf.cell(w=40,h=10,txt=columns[1],ln=0,border=1)
    pdf.cell(w=40,h=10,txt=columns[2],ln=0,border=1)
    pdf.cell(w=30,h=10,txt=columns[3],ln=0,border=1)
    pdf.cell(w=30,h=10,txt=columns[4],ln=1,border=1)



    for index,row in df.iterrows():
        # adding the items list
        pdf.set_font(family="Times",size=10)
        pdf.cell(w=30,h=10,txt=str(row["product_id"]),ln=0,border=1)
        pdf.cell(w=40,h=10,txt=str(row["product_name"]),ln=0,border=1)
        pdf.cell(w=40,h=10,txt=str(row["amount_purchased"]),ln=0,border=1)
        pdf.cell(w=30,h=10,txt=str(row["price_per_unit"]),ln=0,border=1)
        pdf.cell(w=30,h=10,txt=str(row["total_price"]),ln=1,border=1)

    # adding the total sum
    columns_sum = df["total_price"].sum()
    pdf.set_font(family="Times",size=10)
    # pdf.cell(w=30,h=10,txt=str(row["product_id"]),ln=0,border=1)
    # pdf.cell(w=40,h=10,txt=str(row["product_name"]),ln=0,border=1)
    # pdf.cell(w=40,h=10,txt=str(row["amount_purchased"]),ln=0,border=1)
    # pdf.cell(w=30,h=10,txt=str(row["price_per_unit"]),ln=0,border=1)
    pdf.cell(w=140,h=10,txt=" ",ln=0,border=1)
    pdf.cell(w=30,h=10,txt=str(columns_sum),ln=1,border=1)

    # adding two lines of text to conclude
    pdf.set_font(family="Times",size=10)
    pdf.cell(w=30,h=10,txt=f"The total price is {columns_sum}",border=0,ln=0)
    


    pdf.output(f"{filename}.pdf")
