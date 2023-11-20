from sys import displayhook
import pandas as pd
import openpyxl
import datetime

workbook = openpyxl.load_workbook(r'C:\Users\HOME\OneDrive\Backup Notebook\Desktop\Clientes_Antigos_Otica.xlsx')
worksheet = workbook['Vixen_Clientes_Ultima_Compra'] # Especifica a planilha que deve ser lida

start_date = datetime.date(2020, 1, 1)
end_date = datetime.date(2020, 12, 31)

for row in worksheet.iter_rows(min_row=2, values_only=True):
    date_str = row[0] # get the date string from the row
    date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date() # convert the string to a datetime.date object
    if start_date <= date <= end_date:
        print(row)
