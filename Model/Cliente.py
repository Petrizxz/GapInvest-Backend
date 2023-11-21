from openpyxl.reader.excel import load_workbook

RANGE_NAME = 'Clientes!A2:F'


def _verificar_cliente_activ(conta, clientes):
    dic = None
    for i in range(len(clientes)):
        if clientes[i][4] == conta:
            dic = {"Nome": clientes[i][0], "Corretora": clientes[i][1], "Estrategia": clientes[i][2],
                   "Bolsa": clientes[i][3], "Conta": clientes[i][4], "Codinome": clientes[i][5],
                   "Coordenada": 'E' + str(i + 2), "Row": str(i + 2)} # TODO  isso pode dar errado
            break

    return dic


class Cliente:
    def __init__(self, sheet, sheet_id, conta):
        try:
            dadosclientes = sheet.values().get(spreadsheetId=sheet_id, range=RANGE_NAME).execute()

            dadosclientes = dadosclientes['values']
            dic = _verificar_cliente_activ(conta, dadosclientes)
            if dic is None:
                raise ValueError(f"A conta {conta} não foi encontrada na tabela de clientes. Cadastre o cliente na "
                                 f"aba Clientes da planilha")
        except ValueError as erro:
            raise erro
        else:
            self.tab = dadosclientes
            self.nome = dic["Nome"]
            self.conta = dic["Conta"]
            self.codinome = dic["Codinome"]
            self.row = dic["Row"]
            self.coordenada = dic["Coordenada"]
