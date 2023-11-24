from datetime import datetime

from Model.Cliente import Cliente
from Model.Dados import Dados
from Model.Extrato import Extrato

from uteis.autenticacao.Google import Create_Service

from googleapiclient.errors import HttpError

# Upload
from googleapiclient.http import MediaFileUpload

# Download
# import os
# import io

# Autenticação
CLIENT_SECRET_FILE = './uteis/autenticacao/credentials.json'

# Google Drive
# Parcial Oficial: 1Rr2sukEQdKsI1eNWqoJos-rCJ3PtX4Zg
# Parcial Teste: 14NQut5Ir15tKOTnDQRY4661yDmAr4VIY
# Mensal Oficial: 1SilcNjS-Z1qzi8KhqH0q1Tiq3k7vaMrC
# Mensal Teste: 1TdorkJC0obr-xDUei2is064drjL-81F8
FOLDER_PARCIAL_ID = '1Rr2sukEQdKsI1eNWqoJos-rCJ3PtX4Zg'
FOLDER_MENSAL_ID = '1SilcNjS-Z1qzi8KhqH0q1Tiq3k7vaMrC'
# Google Sheets
# Planilha Oficial:  1ogJGSwp6zLsAD3JYq30Ha9Vc9LHNY1Oi9uGiDBpesq8
# Planilha Teste: 1C2vPxShxMLokSewzb86HXRMMCRDgD00DypPt5wZOkHk
# Planilha OPERACIONAL Oficial:  tem q transformar em planilha eletronica pra funciona
# Planilha OPERACIONAL Teste: 1wrDddlkeG63usM_xcJgEzkh-d69rAtUcq5bsXKGmCLU
SHEET_ID = '1ogJGSwp6zLsAD3JYq30Ha9Vc9LHNY1Oi9uGiDBpesq8'
SHEET_ID_OPERACIONAL = '1wrDddlkeG63usM_xcJgEzkh-d69rAtUcq5bsXKGmCLU'

API_NAME_SHEET = 'sheets'
API_VERSION_SHEET = 'v4'
SCOPES_SHEET = ['https://www.googleapis.com/auth/spreadsheets']

API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

# TODO DEPOIS ORGANIZAR ESSES RANGES
RANGE_NAME = 'Dados Activ!A1:S'
RANGE_FORMULAS_DADOS_CLIENTES = 'Clientes!R3:S3'
RANGE_OPERACIONAL = 'Lots Sets!G4:G10'

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

service_sheet = Create_Service(CLIENT_SECRET_FILE, API_NAME_SHEET, API_VERSION_SHEET, SCOPES_SHEET)

testen_planilha = "Novo Resultado Robos ONLINE.xlsx"


def download():
    # UPLOAD
    folder_id = '1Qd55AZduTApSwel8SW_HhkkdFD110nVu'
    file_names = ['lion.xlsx']
    mime_types = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']

    file_metadata = {
        'name': file_names[0],
        'parents': [folder_id]
    }
    name = './{0}'.format(file_names[0])
    media = MediaFileUpload(name, mimetype=mime_types[0])

    service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()


def upload(filename):
    # UPLOAD
    print('Entro na função upload')
    # file_names = ['lion.xlsx']

    print(filename[6:8])
    print(filename)
    if filename[
       6:8] == '01':  # TODO Trocar isso pela logica dos feriado e finais de semana ou eu mando anderson escvrever mensl e parcial nos extratos que é mais facil :3
        FOLDER_ID = FOLDER_MENSAL_ID
    else:
        FOLDER_ID = FOLDER_PARCIAL_ID

    file_names = [filename]
    mime_types = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']

    file_metadata = {
        'name': file_names[0],
        'parents': [FOLDER_ID]
    }
    name = './folder/{0}'.format(file_names[0])
    media = MediaFileUpload(name, mimetype=mime_types[0])

    service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()


def cadastrar_activ(filename):
    try:
        # filename = f"23-10-21 Report Parcial Riquinho.xlsx"
        # Call the Sheets API
        # print(result['range'][18:])

        print("Acessando a planilha online")
        # Pegando a planilha Resultado robos, atraves da api do google sheets
        sheet = service_sheet.spreadsheets()

        print("Buscando as formulas da planilha!")
        # Buscando as fórmulas que serão usadas para cadastrar os novos dados
        formulas = sheet.values().get(spreadsheetId=SHEET_ID, range='Dados Activ!T3:V3',
                                      valueRenderOption='FORMULA').execute()
        formulas = formulas['values']

        print("Lendo o arquivo enviado. Isso pode demorar um pouco!")
        # Criando a Classe Extrato
        extrato = Extrato(filename)

        print("Conferindo se o cliente ja esta cadastrado!")
        # Criando a Classe Cliente
        cliente = Cliente(sheet, SHEET_ID, str(extrato.conta))  # TODO se o cliente não for encontrado

        print("Tudo Certo! Vamos pegar os dados ja cadastrados da activ")
        # Criando a Classe Dados
        dadosactiv = Dados(sheet, SHEET_ID, str(extrato.conta))

    except HttpError as erro:
        raise erro
    except Exception as erro:
        print(f"Houve um erro com a manipulação dos dados. Erro: {erro}")
        raise erro
    else:
        try:
            print("Vamos verificar qual foi o último dado inserido desse cliente.")

            if dadosactiv.ultlinha == "Cadastrar":
                print("ainda não tem dados desse cliente, todo o extrato será cadastrado!")

            extrato.filtrar_tab_por(dadosactiv.ultlinha)  # TODO se o extrato ja tiver sido cadastrado
        except ValueError as erro:
            print(f"Houve um erro no filtro de extrato. Erro: {erro}")
            raise erro
        else:
            if extrato.tab is not None:

                print("Vamos formatar os dados do extrato para colocar na planilha.")

                linha_max_dadosactiv = dadosactiv.max_row + 1
                novosdadosactiv = []

                for linha in extrato.tab:

                    novalinha = []

                    for letra in "ABCDE":
                        novalinha.append(f"=Clientes!{letra + cliente.row}")

                    for celula in linha:
                        if type(celula) is datetime:
                            novalinha.append(celula.strftime('%d/%m/%y %H:%M:%S'))
                        else:
                            novalinha.append(celula.replace('.', ','))

                    for celula in formulas[0]:
                        novalinha.append(celula.replace("3", str(linha_max_dadosactiv)))

                    linha_max_dadosactiv = linha_max_dadosactiv + 1
                    novosdadosactiv.append(novalinha)

                print("Inserindo novos dados na planilha!")
                sheet.values().update(spreadsheetId=SHEET_ID, range=f'Dados Activ!A{dadosactiv.max_row + 1}',
                                      valueInputOption="USER_ENTERED", body={'values': novosdadosactiv}).execute()

                print("Dados inseridos!")
                try:
                    print("Vamos subir o arquivo para o drive.")
                    upload(filename)
                except Exception as erro:
                    raise erro

                print("Processo finalizado!")


#################################################################  POST  ######################################################


def cadastrar_cliente(request):
    valor = {"Codinome": "Riquinho", "Corretora": "Activ", "Estrategia": "Boa", "Bolsa": "Activ", "Conta": "199559"}
    print("Cadastrando cliente na planilha")
    valor = [
        list(valor.values())
    ]
    # Pegando a planilha Resultado robos, atraves da api do google sheets
    sheet = service_sheet.spreadsheets()  # TODO VERIFICAR SE ESSA PUXADA DE SHEET PODE SER GLOBAL PARA CARREGAR SO 1 vez

    dados = Cliente.get_clientes(sheet,
                                 SHEET_ID)  # TODO depois deixar os link/codico tudo dentro dos model ou criar um arquivo com todos, que é so importar
    print(dados['insert_row'])
    sheet.values().update(spreadsheetId=SHEET_ID, range=f'Clientes!A{dados['insert_row']}',
                          valueInputOption="USER_ENTERED", body={'values': valor}).execute()

    print("Dados inseridos!")


def cadastrar_dolar(request):
    print("Cadastrando dolar na planilha")

    sheet = service_sheet.spreadsheets()

    dados = Dados.get_dolar(sheet,
                            SHEET_ID)  # TODO depois deixar os link/codico tudo dentro dos model ou criar um arquivo com todos, que é so importar

    valor = {"Mês": "teste", "Ano": "2022233", "Dolar": "911,48"}
    valor = list(valor.values())
    for i, campo in enumerate(dados['tab']):
        campo.append(valor[i])

    valor = dados['tab']
    # Pegando a planilha Resultado robos, atraves da api do google sheets
    print(dados['insert_row'])
    sheet.values().update(spreadsheetId=SHEET_ID, range=f'B3 Histórico!I22', valueInputOption="USER_ENTERED",
                          body={'values': valor}).execute()  # TODO VER UMA LOGICA DE INSERIR POR COLUNA

    print("Dados inseridos!")


def cadastrar_b3(request):
    print("Cadastrando cotacao b3 na planilha")

    sheet = service_sheet.spreadsheets()

    dados = Dados.get_b3(sheet,
                         SHEET_ID)  # TODO depois deixar os link/codico tudo dentro dos model ou criar um arquivo com todos, que é so importar

    valor = {"Mês": "10", "Ano": "2023", "b3": "113143,67"}
    valor = list(valor.values())

    # Para acessar o dado do ultimo mês para calcular a taxa
    tratamento = dados['tab'][2]
    tratamento = tratamento[len(tratamento) - 1]
    tratamento = float(tratamento.replace(',', '.'))

    part = float(valor[2].replace(',', '.')) / tratamento
    part = part - 1

    valor.append(part)
    for i, campo in enumerate(dados['tab']):
        campo.append(valor[i])

    valor = dados['tab']
    # Pegando a planilha Resultado robos, atraves da api do google sheets
    sheet.values().update(spreadsheetId=SHEET_ID, range=f'B3 Histórico!I17', valueInputOption="USER_ENTERED",
                          body={'values': valor}).execute()  # TODO VER UMA pra deixar o range automatico

    print("Dados inseridos!")


def cadastrar_dados_cliente(filename):
    info_cliente = {
        "Cliente": {
            "codinome": "XP MB",
            "Coordenada": ['A2', 'B2', 'C2', 'D2', 'E2']
        },
        "IR": "11234,98",
        "Custo básico": "720",
        "Taxa performance": "0",
        "Taxa IR": "0",
        "Data": "31/12/2023"
    }
    # Pegando a planilha Resultado robos, atraves da api do google sheets
    sheet = service_sheet.spreadsheets()
    formulas = sheet.values().get(spreadsheetId=SHEET_ID, range=RANGE_FORMULAS_DADOS_CLIENTES,
                                  valueRenderOption='FORMULA').execute()

    formulas = formulas['values']
    dados_clientes = Cliente.get_dados_clientes(sheet, SHEET_ID)
    novalinha = [[]]

    for campos in info_cliente.values():
        if type(campos) is dict:
            for coordenada in campos['Coordenada']:
                novalinha[0].append(f"={coordenada}")
        else:
            novalinha[0].append(campos)

    for formula in formulas[0]:
        novalinha[0].append(formula.replace("3", dados_clientes['insert_row']))

    print("Inserindo novos dados na planilha!")
    sheet.values().update(spreadsheetId=SHEET_ID, range=f'Clientes!H{dados_clientes['insert_row']}',
                          valueInputOption="USER_ENTERED", body={'values': novalinha}).execute()


def cadastrar_operacional(filename):
    valor = {"Lion": "4896", "Riquinho": "-850", "Day Trade B3": "", "Munra": "720", "Madalena B3": "",
             "Guerreira": "21355"}
    print("Cadastrando Floating na planilha")
    colunas = []

    for campo in valor.values():
        colunas.append([campo])

    # Pegando a planilha Resultado robos, atraves da api do google sheets
    sheet = service_sheet.spreadsheets()

    sheet.values().update(spreadsheetId=SHEET_ID_OPERACIONAL, range=f'Lots Sets!G5',
                          valueInputOption="USER_ENTERED", body={'values': colunas}).execute()

    print("Dados inseridos!")


#################################################################  GET  ######################################################

def get_cliente(request):
    sheet = service_sheet.spreadsheets()
    return Cliente.get_clientes(sheet, SHEET_ID)


if __name__ == '__main__':
    print(1 + 1)
    # cadastrar_activ('')
    # upload()
    # download()
