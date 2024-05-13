import pyfiglet 
import openpyxl
from datetime import date


# ============== ACCESSING SHEETS IN EXCEL =============== #

wb = openpyxl.load_workbook('Estoque e registro.xlsx')
stockSheet = wb['Estoque']
registerSheet = wb['Registro']

# ============== FUNCTIONS ============  ================= #

def insertDateToBrazilianFormart(): ## date format
    today = date.today()
    
    strDay = str(today.day)
    strMonth = str(today.month)
    strYear = str(today.year)

    currentlyDate = strDay + '/' + strMonth + '/' + strYear

    return currentlyDate
    
def insertDataOnSheets(positionsValuesInOrder): # This function will check if the currently position is empty or not. If is empty, the empty space will be fiiled by the datas from inputs. Observartion: This functions just saves on Register Sheet.
    position = 3

    strPosition = str(position)

    cellColumnPosition = ['A','B','C','D','E']

    cellRegisterSelect = [registerSheet['A'+strPosition].value,registerSheet['B'+strPosition].value,registerSheet['C'+strPosition].value,registerSheet['D'+strPosition].value,registerSheet['E'+strPosition].value]

    for i in range(0,5): # this for loop, will check if all positions of sheet is empty
       # while cellRegisterSelect[i] != None:
        if cellColumnPosition[i] != None:
            print('bla')
        else:
            cellRegisterSelect[i] = str(positionsValuesInOrder[i])
        
       
    cellRegisterSelect[i] = positionsValuesInOrder[i]
    print(cellColumnPosition)
    print(cellRegisterSelect)
    

      

# ====== THE MAIN FUNCIONALITY ====== #

print('Bem-vindo ao sistema de gerenciamento de equipamentos da TI \n' + pyfiglet.figlet_format('Voxline', font='slant'))

device = str.upper(input('Qual equipamento vai ser movimentado: '))
qtd = int(input('Quantos {}s serão movimentados? '.format(device)))   
statusCondicional = 0
while statusCondicional == 0:
     ticketCondicional = str.upper(input('Tem número de chamado?  '))
    
     if ticketCondicional == 'SIM' or ticketCondicional == '':
        ticket = input('Insira o número do chamado: ')
        statusCondicional = statusCondicional + 1
     elif ticketCondicional == 'NÃO' or ticketCondicional == 'NAO' or ticketCondicional == 'N':
        print('O campo CHAMADO ficará NULO')
        ticket = 'NULO'
        statusCondicional = statusCondicional + 1
     else: 
        print('Resposta não reconhecida. Responda com SIM ou NÃO')
    
addOrChange = str.upper(input('O {} será uma troca ou adição à PA? '.format(device)))

currentDate = insertDateToBrazilianFormart()

positionsValuesInOrder = [ticket,device,qtd, addOrChange, currentDate ]



wb.save('Estoque e registro.xlsx') 