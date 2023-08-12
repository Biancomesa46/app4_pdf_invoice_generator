import glob
from fpdf import FPDF
import pandas



filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pandas.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="p", unit="mm", format="a4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    pdf.set_text_color(100, 100, 100)
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(txt=f"Invoice nr.{filepath[9:14]}", border=0, h=10, w=0, ln=1, align='l')
    pdf.cell(txt=f"Date {filepath[15:24]}", h=10, w=0, ln=1, border=0, align='l')
    pdf.output("PDFs/" + filepath.strip('invoices\\.xlsx') + ".pdf")