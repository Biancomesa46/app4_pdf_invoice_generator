import glob
from fpdf import FPDF
import pandas

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    pdf = FPDF(orientation="p", unit="mm", format="a4")
    pdf.set_auto_page_break(auto=False, margin=0)

    pdf.add_page()
    pdf.set_text_color(100, 100, 100)
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(txt=f"Invoice nr.{filepath[9:14]}", border=0, h=10, w=0, ln=1, align='l')
    pdf.cell(txt=f"Date {filepath[15:24]}", h=10, w=0, ln=1, border=0, align='l')

    df = pandas.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)
    replaced_columns = [item.replace("_", " ").title() for item in columns]

    pdf.set_font(family="Times", size=8, style="B")
    pdf.cell(txt=replaced_columns[0], h=8, w=30, border=1)
    pdf.cell(txt=replaced_columns[1], h=8, w=60, border=1)
    pdf.cell(txt=replaced_columns[2], h=8, w=30, border=1)
    pdf.cell(txt=replaced_columns[3], h=8, w=30, border=1)
    pdf.cell(txt=replaced_columns[4], h=8, w=30, border=1, ln=1)

    pdf.set_font(family="Times", size=8)
    for index, row in df.iterrows():
        pdf.cell(txt=str(row['product_id']), h=8, w=30, border=1)
        pdf.cell(txt=str(row['product_name']), h=8, w=60, border=1)
        pdf.cell(txt=str(row['amount_purchased']), h=8, w=30, border=1)
        pdf.cell(txt=str(row['price_per_unit']), h=8, w=30, border=1)
        pdf.cell(txt=str(row['total_price']), h=8, w=30, border=1, ln=1)
    pdf.cell(txt=" ", h=8, w=30, border=1)
    pdf.cell(txt=" ", h=8, w=60, border=1)
    pdf.cell(txt=" ", h=8, w=30, border=1)
    pdf.cell(txt=" ", h=8, w=30, border=1)
    pdf.cell(txt=str(df["total_price"].sum()), h=8, w=30, border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(border=0, ln=1, h=8, w=30, txt=f"The total due amount is {df['total_price'].sum()} Euros.")

    pdf.output("PDFs/" + filepath.strip('invoices\\.xlsx') + ".pdf")
