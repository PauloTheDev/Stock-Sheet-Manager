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
   
   pos = 3
   
   i = 0
   columnsPosition = ['A','B','C','D','E']
   
   while i < 5:
    posStr = str(pos)
    positions = str(columnsPosition[i]+posStr)
    
    if registerSheet[positions].value == None:
           registerSheet[positions].value = positionsValuesInOrder[i]
           i = i + 1
           
    else:
           
           pos = pos + 1
    #explaing the function: the while will verify if the i (index) is lower than 5. While i is lower than 5 the loop will verify if the currently position in sheet is free. 

def MakeChangesOnStock (device, qtd, inOrOut):
    tecladoValueCell = stockSheet['B3'].value
    
# ====== THE MAIN FUNCIONALITY ====== #

print('Bem-vindo ao sistema de gerenciamento de equipamentos da TI \n' + pyfiglet.figlet_format('Voxline', font='slant'))

statusCondicionalDevice = False
while statusCondicionalDevice == False:
    device = str.upper(input('Qual equipamento vai ser movimentado: '))

    if device == 'TECLADO' or device == 'MOUSE' or device == 'HEADSET KJ' or device == 'HEADSET PNP':
        statusCondicionalDevice = True

    elif device == 'PNP' or device == 'KJ':
        device = 'HEADSET'+' '+ device
        statusCondicionalDevice=True

    elif device == 'HEADSET': #this condicional will be used when the user write just "headset"
        chooseTypeOfHeadset = str.upper(input('Você digitou {}. Seria KJ ou PnP? '.format(device)))

        statusCondicionalHeadset = False
        while statusCondicionalHeadset == False:
            if chooseTypeOfHeadset == 'PNP' or chooseTypeOfHeadset == 'KJ':
                device = device + ' ' + chooseTypeOfHeadset.capitalize()
                statusCondicionalHeadset = True
                statusCondicionalDevice = True
   
    else:
        print('Você digitou {} e isso não foi reconhecido pelo gerenciador. Digite novamente.'.format(device))
    
    qtd = int(input('Qual a quantidade? '))
    



statusCondicionalTicket = False
while statusCondicionalTicket == False:
     ticketCondicional = str.upper(input('Tem número de chamado?  '))
    
     if ticketCondicional == 'SIM' or ticketCondicional == '':
        ticket = input('Insira o número do chamado: ')
        statusCondicionalTicket = True
     elif ticketCondicional == 'NÃO' or ticketCondicional == 'NAO' or ticketCondicional == 'N':
        print('O campo CHAMADO ficará NULO')
        ticket = 'NULO'
        statusCondicionalTicket = True
     else: 
        print('Resposta não reconhecida. Responda com SIM ou NÃO')
    
statusCondicionalAddOrChange = False
while statusCondicionalAddOrChange == False:
    addOrChange = str.upper(input('O {} será uma entrada ou saída de equipamento? '.format(device)))
    if addOrChange == 'ENTRADA' or addOrChange == 'SAÍDA' or addOrChange == 'SAIDA':
        statusCondicionalAddOrChange = True
    else:
        print('Por favor, escreva ENTRADA ou SAIDA')

currentDate = insertDateToBrazilianFormart()

positionsValuesInOrder = [ticket,device,qtd, addOrChange, currentDate ]

insertDataOnSheets(positionsValuesInOrder)
wb.save('Estoque e registro.xlsx')
print('Registro salvo com sucesso.')

