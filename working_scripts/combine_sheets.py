import pandas as pd

workbook_file = pd.ExcelFile("c:\\users\\u6zn\\desktop\\Obsolete Inactive Excess 2021.xlsx")

sheets = workbook_file.sheet_names

# read_sheets = pd.read_excel(workbook_file, sheetname=None)
combined_sheets = pd.concat([pd.read_excel(workbook_file, sheet_name=s).assign(sheet_name=s) for s in sheets])

combined_sheets.to_excel("c:\\users\\u6zn\\desktop\\2021 Write-Off Combined Sheets.xlsx", index=False)
