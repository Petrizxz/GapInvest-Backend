from datetime import datetime

from Model.Cliente import Cliente
from Model.Dados import Dados
from Model.Extrato import Extrato

from uteis.autenticacao.Google import Create_Service

from uteis.constantes.const import C

from googleapiclient.errors import HttpError

# Upload
from googleapiclient.http import MediaFileUpload

import logging

logging.basicConfig(level=logging.DEBUG)

service = Create_Service(C.CLIENT_SECRET_FILE(), C.API_NAME(), C.API_VERSION(), C.SCOPES())

service_sheet = Create_Service(C.CLIENT_SECRET_FILE(), C.API_NAME_SHEET(), C.API_VERSION_SHEET(), C.SCOPES_SHEET())


# teste

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
    folder_id = ''
    if filename[
       6:8] == '01':  # TODO Trocar isso pela logica dos feriado e finais de semana ou eu mando anderson escvrever mensl e parcial nos extratos que é mais facil :3
        folder_id = C.FOLDER_MENSAL_ID()
    else:
        folder_id = C.FOLDER_PARCIAL_ID()

    file_names = [filename]
    mime_types = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']

    file_metadata = {
        'name': file_names[0],
        'parents': [folder_id]
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
        print("Acessando a planilha online")
        # Pegando a planilha Resultado robôs, através da api do google sheets
        sheet = service_sheet.spreadsheets()

        print("Buscando as formulas da planilha!")
        # Buscando as fórmulas que serão usadas para cadastrar os novos dados
        formulas = sheet.values().get(spreadsheetId=C.SHEET_ID(), range='Dados Activ!T3:V3',
                                      valueRenderOption='FORMULA').execute()
        formulas = formulas['values']

        print("Lendo o arquivo enviado. Isso pode demorar um pouco!")
        # Criando a Classe Extrato
        extrato = Extrato(filename)

        print("Conferindo se o cliente ja esta cadastrado!")
        # Criando a Classe Cliente
        cliente = Cliente(sheet, C.SHEET_ID(), str(extrato.conta))  # TODO se o cliente não for encontrado

        print("Tudo Certo! Vamos pegar os dados ja cadastrados da activ")
        # Criando a Classe Dados
        dados_activ = Dados(sheet, C.SHEET_ID(), str(extrato.conta))

    except HttpError as erro:
        raise erro
    except Exception as erro:
        print(f"Houve um erro com a manipulação dos dados. Erro: {erro}")
        raise erro
    else:
        try:
            print("Vamos verificar qual foi o último dado inserido desse cliente.")

            if dados_activ.ultlinha == "Cadastrar":
                print("ainda não tem dados desse cliente, todo o extrato será cadastrado!")

            extrato.filtrar_tab_por(dados_activ.ultlinha)  # TODO se o extrato ja tiver sido cadastrado
        except ValueError as erro:
            print(f"Houve um erro no filtro de extrato. Erro: {erro}")
            raise erro
        else:
            if extrato.tab is not None:

                print("Vamos formatar os dados do extrato para colocar na planilha.")

                linha_max_dados_activ = dados_activ.max_row + 1
                novos_dados_activ = []

                for linha in extrato.tab:

                    nova_linha = []

                    for letra in "ABCDE":
                        nova_linha.append(f"=Clientes!{letra + cliente.row}")

                    for celula in linha:
                        if type(celula) is datetime:
                            nova_linha.append(celula.strftime('%d/%m/%y %H:%M:%S'))
                        else:
                            nova_linha.append(celula.replace('.', ','))

                    for celula in formulas[0]:
                        nova_linha.append(celula.replace("3", str(linha_max_dados_activ)))

                    linha_max_dados_activ = linha_max_dados_activ + 1
                    novos_dados_activ.append(nova_linha)

                print("Inserindo novos dados na planilha!")
                sheet.values().update(spreadsheetId=C.SHEET_ID(), range=f'Dados Activ!A{dados_activ.max_row + 1}',
                                      valueInputOption="USER_ENTERED", body={'values': novos_dados_activ}).execute()

                print("Dados inseridos!")
                try:
                    print("Vamos subir o arquivo para o drive.")
                    upload(filename)
                except Exception as erro:
                    raise erro

                print("Processo finalizado!")


#################################################################  POST  ######################################################


def cadastrar_cliente(request):
    print("Cadastrando cliente na planilha")
    # valor = {"Codinome": "Riquinho", "Corretora": "Activ", "Estrategia": "Boa", "Bolsa": "Activ", "Conta": "199559"}
    valor = [
        # list(valor.values())
        list(request.values())
    ]

    # Pegando a planilha Resultado robôs, através da api do google sheets
    sheet = service_sheet.spreadsheets()  # TODO VERIFICAR SE ESSA PUXADA DE SHEET PODE SER GLOBAL PARA CARREGAR SO 1 vez

    dados = Cliente.get_clientes(sheet, C.SHEET_ID())
    print(dados['insert_row'])
    sheet.values().update(spreadsheetId=C.SHEET_ID(), range=f'Clientes!A{dados["insert_row"]}',
                          valueInputOption="USER_ENTERED", body={'values': valor}).execute()
    print("Dados inseridos!")


def cadastrar_dolar(request):
    print("Cadastrando dolar na planilha")

    # valor = {"Mês": "teste", "Ano": "2022233", "Dolar": "911,48"}
    # valor = list(valor.values())
    valor = list(request.values())

    sheet = service_sheet.spreadsheets()

    dados = Dados.get_dolar(sheet, C.SHEET_ID())

    for i, campo in enumerate(dados['tab']):
        campo.append(valor[i])

    valor = dados['tab']
    # Pegando a planilha Resultado robôs, através da api do google sheets
    print(dados['insert_row'])
    sheet.values().update(spreadsheetId=C.SHEET_ID(), range=f'B3 Histórico!I22', valueInputOption="USER_ENTERED",
                          body={'values': valor}).execute()  # TODO VER UMA LOGICA DE INSERIR POR COLUNA

    print("Dados inseridos!")


def cadastrar_b3(request):
    print("Cadastrando cotação b3 na planilha")

    valor = {"Mês": "10", "Ano": "2023", "b3": "113143,67"}
    valor = list(valor.values())
    # valor = list(request.values())

    sheet = service_sheet.spreadsheets()

    dados = Dados.get_b3(sheet, C.SHEET_ID())

    # Para acessar o dado do ultimo mês para calcular a taxa
    print("fazendo o tratamento de dados")
    try:

        tratamento = dados['tab'][2]
        tratamento = tratamento[len(tratamento) - 1]
        tratamento = float(tratamento.replace(',', '.'))
        print("Tratamento feito")

        part = float(valor[2].replace(',', '.')) / tratamento
        part = part - 1
        print("Conta realizada")

    except ValueError as erro:
        raise ValueError(f"Não foi possível calcular a taxa da b3. Erro: {erro}")
    else:
        valor.append(part)

        for i, campo in enumerate(dados['tab']):
            campo.append(valor[i])

        valor = dados['tab']
        # Pegando a planilha Resultado robôs, através da api do google sheets
        print("Vai enviar para a planilha")
        sheet.values().update(spreadsheetId=C.SHEET_ID(), range=f'B3 Histórico!I17', valueInputOption="USER_ENTERED",
                              body={'values': valor}).execute()  # TODO VER UMA pra deixar o range automatico

        print("Dados inseridos!")


def cadastrar_dados_cliente(request):
    print("Atualizando dados na planilha")

    # info_cliente = {
    #    "Cliente": {
    #        "codinome": "XP MB",
    #        "Coordenada": ['A2', 'B2', 'C2', 'D2', 'E2']
    #    },
    #    "IR": "11234,98",
    #    "Custo básico": "720",
    #    "Taxa performance": "0",
    #    "Taxa IR": "0",
    #    "Data": "31/12/2023"
    # }

    # Pegando a planilha Resultado robôs, através da api do google sheets
    sheet = service_sheet.spreadsheets()
    formulas = sheet.values().get(spreadsheetId=C.SHEET_ID(), range=C.RANGE_FORMULAS_DADOS_CLIENTES(),
                                  valueRenderOption='FORMULA').execute()

    formulas = formulas['values']
    dados_clientes = Cliente.get_dados_clientes(sheet, C.SHEET_ID())
    nova_linha = [[]]

    # for campos in info_cliente.values():
    for campos in request.values():
        if type(campos) is dict:
            for coordenada in campos['Coordenadas']:
                nova_linha[0].append(f"={coordenada}")
        else:
            nova_linha[0].append(campos)

    for formula in formulas[0]:
        nova_linha[0].append(formula.replace("3", dados_clientes['insert_row']))

    sheet.values().update(spreadsheetId=C.SHEET_ID(), range=f'Clientes!H{dados_clientes['insert_row']}',
                          valueInputOption="USER_ENTERED", body={'values': nova_linha}).execute()
    print("Dados inseridos!")


def cadastrar_operacional(request):
    print("Cadastrando Floating na planilha")

    # valor = {"Lion": "4896", "Riquinho": "-850", "Day Trade B3": "", "Munra": "720", "Madalena B3": "",
    #         "Guerreira": "21355"}

    colunas = []

    # for campo in valor.values():
    for campo in request.values():
        colunas.append([campo])

    # Pegando a planilha Resultado robôs, através da api do google sheets
    sheet = service_sheet.spreadsheets()

    sheet.values().update(spreadsheetId=C.SHEET_ID_OPERACIONAL(), range=f'Lots Sets!G5',
                          valueInputOption="USER_ENTERED", body={'values': colunas}).execute()

    print("Dados inseridos!")


#################################################################  GET  ######################################################

def get_cliente():
    sheet = service_sheet.spreadsheets()
    return Cliente.get_clientes(sheet, C.SHEET_ID())


def get_verifica_cliente(conta, codinome):

    try:
        sheet = service_sheet.spreadsheets()
        dados_clientes = sheet.values().get(spreadsheetId=C.SHEET_ID(), range=C.RANGE_CLIENTE()).execute()
        dados_clientes = dados_clientes["values"]
    except ValueError as erro:
        raise erro
    else:
        try:
            conta_verificada = Cliente.verificar_cliente_activ(conta, dados_clientes)
        except ValueError as erro:
            raise erro
        else:
            if conta_verificada["Codinome"].lower() == codinome.lower():
                return conta_verificada
            else:
                raise ValueError(f"O Codinome {codinome} não foi encontrado.")





if __name__ == '__main__':
    print(1 + 1)
    # cadastrar_activ('')
    # upload()
    # download()
