import glob

import pandas

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pandas.read_excel(filepath, sheet_name="Sheet 1")