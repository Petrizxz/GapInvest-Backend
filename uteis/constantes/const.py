######################################### LOCALIZAÇÃO DE ARQUIVOS  ############################################################################
CLIENT_SECRET_FILE = './uteis/autenticacao/credentials.json'
UPLOAD_FOLDER = 'folder'

####################################################CONFIGURAÇÃO DAS OPENPYXL ########################################################################################

# Caso os extratos tiver outro formato de Excel é so adicionar aqui
ALLOWED_EXTENSIONS_SHEET = set(['xlsx'])
ALLOWED_EXTENSIONS_IMG = set(['png', 'jpg', 'jpeg'])
######################################### CONFIGURAÇÃO DAS API ############################################################################
API_NAME_SHEET = 'sheets'
API_VERSION_SHEET = 'v4'
SCOPES_SHEET = ['https://www.googleapis.com/auth/spreadsheets']

API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

######################################### CONFIGURAÇÃO DAS PLANILHAS ############################################################################

SHEET_ID = '1ogJGSwp6zLsAD3JYq30Ha9Vc9LHNY1Oi9uGiDBpesq8'
SHEET_ID_OPERACIONAL = '1wrDddlkeG63usM_xcJgEzkh-d69rAtUcq5bsXKGmCLU'

""" 
Google Sheets - API - INFORMAÇÕES COMPLEMENTARES 
Planilha Oficial:  1ogJGSwp6zLsAD3JYq30Ha9Vc9LHNY1Oi9uGiDBpesq8
Planilha Teste: 1C2vPxShxMLokSewzb86HXRMMCRDgD00DypPt5wZOkHk
Planilha OPERACIONAL Oficial:  tem q transformar em planilha eletronica pra funciona
Planilha OPERACIONAL Teste: 1wrDddlkeG63usM_xcJgEzkh-d69rAtUcq5bsXKGmCLU
"""

######################################### CONFIGURAÇÃO DAS PASTAS PARA UPLOAD ############################################################################

FOLDER_PARCIAL_ID = '1Rr2sukEQdKsI1eNWqoJos-rCJ3PtX4Zg'
FOLDER_MENSAL_ID = '1SilcNjS-Z1qzi8KhqH0q1Tiq3k7vaMrC'

""" 
# Google Drive - API - INFORMAÇÕES COMPLEMENTARES
# Parcial Oficial: 1Rr2sukEQdKsI1eNWqoJos-rCJ3PtX4Zg
# Parcial Teste: 14NQut5Ir15tKOTnDQRY4661yDmAr4VIY
# Mensal Oficial: 1SilcNjS-Z1qzi8KhqH0q1Tiq3k7vaMrC
# Mensal Teste: 1TdorkJC0obr-xDUei2is064drjL-81F8
"""

######################################### CONFIGURAÇÃO DE RANGE DE PLANILHAS ############################################################################

RANGE_NAME = 'Dados Activ!A2:S'

RANGE_OPERACIONAL = 'Lots Sets!G4:G10'

RANGE_CLIENTE = 'Clientes!A2:E'
RANGE_FORMULAS_DADOS_CLIENTES = 'Clientes!R3:S3'
RANGE_DADOS_CLIENTES = 'Clientes!H2:S'

RANGE_DOLAR = 'B3 Histórico!I22:24'

RANGE_B3 = 'B3 Histórico!I17:20'


class C:

    @staticmethod
    def CLIENT_SECRET_FILE():
        return CLIENT_SECRET_FILE

    @staticmethod
    def UPLOAD_FOLDER():
        return UPLOAD_FOLDER

    @staticmethod
    def ALLOWED_EXTENSIONS_SHEET():
        return ALLOWED_EXTENSIONS_SHEET

    @staticmethod
    def ALLOWED_EXTENSIONS_IMG():
        return ALLOWED_EXTENSIONS_IMG

    @staticmethod
    def API_NAME_SHEET():
        return API_NAME_SHEET

    @staticmethod
    def API_VERSION_SHEET():
        return API_VERSION_SHEET

    @staticmethod
    def SCOPES_SHEET():
        return SCOPES_SHEET

    @staticmethod
    def API_NAME():
        return API_NAME

    @staticmethod
    def API_VERSION():
        return API_VERSION

    @staticmethod
    def SCOPES():
        return SCOPES

    @staticmethod
    def SHEET_ID():
        return SHEET_ID

    @staticmethod
    def SHEET_ID_OPERACIONAL():
        return SHEET_ID_OPERACIONAL

    @staticmethod
    def FOLDER_PARCIAL_ID():
        return FOLDER_PARCIAL_ID

    @staticmethod
    def FOLDER_MENSAL_ID():
        return FOLDER_MENSAL_ID

    @staticmethod
    def RANGE_NAME():
        return RANGE_NAME

    @staticmethod
    def RANGE_OPERACIONAL():
        return RANGE_OPERACIONAL

    @staticmethod
    def RANGE_CLIENTE():
        return RANGE_CLIENTE

    @staticmethod
    def RANGE_FORMULAS_DADOS_CLIENTES():
        return RANGE_FORMULAS_DADOS_CLIENTES

    @staticmethod
    def RANGE_DADOS_CLIENTES():
        return RANGE_DADOS_CLIENTES

    @staticmethod
    def RANGE_DOLAR():
        return RANGE_DOLAR

    @staticmethod
    def RANGE_B3():
        return RANGE_B3

