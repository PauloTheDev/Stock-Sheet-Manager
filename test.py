import pyfiglet 
import openpyxl
from datetime import date


# ============== ACCESSING SHEETS IN EXCEL =============== #

wb = openpyxl.load_workbook('Estoque e registro.xlsx')
stockSheet = wb['Estoque']
registerSheet = wb['Registro']

print(registerSheet['A3'].value)
registerSheet['A3'].value = 'bungas'
print(registerSheet['A3'].value)
wb.save('Estoque e registro.xlsx')