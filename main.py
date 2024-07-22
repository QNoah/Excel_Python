from openpyxl import Workbook, load_workbook
from datetime import date

book = load_workbook('menu.xlsx')
# sheet = book.active
book.active = book['Blad1']

value = input('Typ een titel: \n')

row = 2 

while book.active[f'A{row}'].value is not None:
    row += 1
    
book.active[f'A{row}'].value = value
book.active[f'B{row}'].value = date.today()   

for cell in book.active['A']:
    print(cell.value)

book.save('menu.xlsx')

print(f'Titel "{value}" is toegevoegd in rij {row}.')


