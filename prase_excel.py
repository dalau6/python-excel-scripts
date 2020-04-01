import pandas as pd
import openpyxl

# read excel
files = ['https://github.com/datagy/mediumdata/raw/master/january.xlsx', 'https://github.com/datagy/mediumdata/raw/master/february.xlsx', 'https://github.com/datagy/mediumdata/raw/master/march.xlsx']
combined = pd.DataFrame()

for file in files:
  df = pd.read_excel(file, skiprows = 3)
  combined = combined.append(df, ignore_index = True)
  
combined.to_excel('combined.xlsx')

# gets specific cells from various references and append
files = [] #include paths to your files here
values = []

for file in files:
    wb = openpyxl.load_workbook(file)
    sheet = wb['Sheet1']
    value = sheet['F5'].value
    values.append(value)

# apply formulas for various references
for file in files:
    wb = openpyxl.load_workbook(file)
    sheet = wb['Sheet1']
    sheet['F9'] = '=SUM(F5:F8)'
    sheet['F9'].style = 'Currency'
    wb.save(file)