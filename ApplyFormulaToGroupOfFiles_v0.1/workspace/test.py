import openpyxl

wb = openpyxl.load_workbook('temp.xlsx', data_only=True)

ws = wb['Summary2']

print(ws['B2'].value)
