import openpyxl,pprint
import os
from openpyxl.cell import get_column_letter, column_index_from_string
os.chdir('/Users/Soudehm/Documents/July4thWTW/WBS')

wb=openpyxl.load_workbook('WTW_DB.xlsx')
sheet = wb.active
sheet.title = 'Spam Spam Spam'
sheet['A1']= 'Hello World!'
wb.save('WTW_DB_copy.xlsx')