from dotenv import load_dotenv
import os
from trello import TrelloClient
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import logging
logging.basicConfig(filename = "application.log", level = logging.DEBUG)

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open_by_key(os.getenv("HASH_SHEET"))

worksheet = wks.get_worksheet(0)

client = TrelloClient(
    api_key= os.getenv("TRELLO_API_KEY"),
    api_secret= os.getenv("TRELLO_API_SECRET"),
)

all_boards = client.list_boards()
my_board = all_boards[3]
all_lists = my_board.list_lists()

#Separa as listas
doc_list = all_lists[0]
backlog_list = all_lists[1]
desenv_list = all_lists[3]
homolog_list = all_lists[5]
prod_list = all_lists[7]
cancel_list = all_lists[8]

#Recupera todos os cards das principais listas
cards_sonda_list = backlog_list.list_cards() + desenv_list.list_cards() + homolog_list.list_cards() + prod_list.list_cards() + cancel_list.list_cards()
logging.info(cards_sonda_list)

def dateFormater(dateForm):
    return f"{dateForm.day}/{dateForm.month}/{dateForm.year}"

#Atualiza celulas

cell_list = worksheet.range('A1:Q1')

cell_list[0].value = 'Projeto'
cell_list[1].value = 'Situação'
cell_list[2].value = 'Descrição'
cell_list[3].value = 'URL'
cell_list[4].value = 'Data entrega'
cell_list[5].value = 'Data criação'
cell_list[6].value = 'Labels'
cell_list[8].value = 'p1'
cell_list[9].value = 'p2'
cell_list[10].value = 'p3'
cell_list[11].value = 'p4'
cell_list[12].value = 'p5'
cell_list[13].value = 'p6'
cell_list[14].value = 'p7'
cell_list[16].value = 'Dev'

worksheet.update_cells(cell_list)

#Método para limpar os projetos alterados em cada card
quantidade = len(cards_sonda_list)
print("Número de cards para atualizar: " + str(quantidade))

# Determina o intervalo de células a serem limpas
range_start = 'I2'
range_end = 'O' + str(quantidade + 1)

# Limpa o intervalo de células
range_to_clear = worksheet.range(range_start + ':' + range_end)
for cell in range_to_clear:
    cell.value = ''

worksheet.update_cells(range_to_clear)

print('Projetos limpos das células ' + range_start + ' a ' + range_end)
logging.info('Projetos limpos das células ' + range_start + ' a ' + range_end)

#Contador de colunas e celulas
columns = 1
cell = 2
updates = []

#Método principal, onde atualiza todas as células de cada projeto
def upgradeCardSheet(card):
    global cell
    global updates
    logging.info(card)
    print(card)

    data = [
        ('A', card.name),
        ('B', card.trello_list.name),
        ('C', card.desc),
        ('D', card.short_url)
    ]

    if card.due_date != '' and card.trello_list.name == 'Entregue/Produção':
        date = dateFormater(card.due_date)
        data.append(('E', date))
    else:
        data.append(('E', ''))

    if card.card_created_date != '':
        date = dateFormater(card.card_created_date)
        data.append(('F', date))

    labels = ",".join([x.name for x in card.labels])
    data.append(('G', labels))

    countProjects = len(card.checklists)
    if countProjects >= 2:
        checkProjeto = card.checklists[1].items
        coluna = 9
        for i, x in enumerate(checkProjeto):
            nomeProjeto = x.get('name', None)
            data.append((chr(coluna + i + 64), nomeProjeto))

    data.append(('Q', card.idMembers[0]))

    for column, value in data:
        updates.append({
            'range': f'{column}{cell}',
            'values': [[value]]
        })

    cell += 1


#Preenche dados da planilha
for card in cards_sonda_list:
    upgradeCardSheet(card)


worksheet.batch_update(updates, value_input_option='USER_ENTERED')

#Informa que finalizou as cargas
logging.info("Cards generate on sheet")
print("Cards generate on sheet")
print("Pressione enter para fechar o programa.")
input()  # Aguarda o usuário pressionar Enter