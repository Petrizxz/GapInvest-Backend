from uteis.constantes.const import C


class Cliente:
    def __init__(self, sheet, sheet_id, conta):
        try:
            dados_clientes = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_CLIENTE()).execute()

            dados_clientes = dados_clientes["values"]
            for i in dados_clientes:
                print(i)
                for j in i:
                    print(j)
            dic = self.verificar_cliente_activ(conta, dados_clientes)

        except ValueError as erro:
            raise erro
        else:
            self.tab = dados_clientes
            self.conta = dic["Conta"]
            self.codinome = dic["Codinome"]
            self.row = dic["Row"]
            self.coordenada = dic["Coordenada"]

    @staticmethod
    def get_clientes(sheet, sheet_id):
        alfabeto = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                    'U', 'V', 'W', 'X', 'Y', 'Z']
        try:
            dados_clientes = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_CLIENTE()).execute()
            dados_clientes = dados_clientes['values']
            # TODO DEPOIS FAZER UMA LOGICA PARA deixar isso em todas os outros q tem q buscar dados, tem um site no whatsapp que mostra como gerar list de A ate ZZZ
            coordenadas = []
            for i in range(len(dados_clientes)):
                coordenadas.append([])
                for j in range(len(dados_clientes[i])):
                    coordenadas[i].append(alfabeto[j] + str(i + 2))

            dados_clientes = {
                "tab": dados_clientes,
                "coordenadas": coordenadas,
                "insert_row": str(len(dados_clientes) + 2)
            }
        except ValueError as erro:
            raise ValueError(f"Não foi possível buscar os clientes cadastrados na planilha. Erro: {erro}")
        else:
            return dados_clientes

    @staticmethod
    def get_dados_clientes(sheet, sheet_id):
        alfabeto = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                    'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        try:
            dados_clientes = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_DADOS_CLIENTES()).execute()
            dados_clientes = dados_clientes['values']
            # TODO DEPOIS FAZER UMA LOGICA PARA deixar isso em todas os outros q tem q buscar dados, tem um site no whatsapp que mostra como gerar list de A ate ZZZ
            coordenadas = []
            for i in range(len(dados_clientes)):
                coordenadas.append([])
                for j in range(len(dados_clientes[i])):
                    coordenadas[i].append(alfabeto[j] + str(i + 2))

            dados_clientes = {
                "tab": dados_clientes,
                "coordenadas": coordenadas,
                "insert_row": str(len(dados_clientes) + 2)
            }
        except ValueError as erro:
            raise ValueError(f"Não foi possível buscar os dados dos clientes da planilha. Erro: {erro}")
        else:
            return dados_clientes

    @staticmethod
    def verificar_cliente_activ(conta, clientes):
        dic = None
        for i in range(len(clientes)):
            if clientes[i][4] == conta:
                dic = {"Codinome": clientes[i][0], "Corretora": clientes[i][1], "Estrategia": clientes[i][2],
                       "Bolsa": clientes[i][3], "Conta": clientes[i][4],
                       "Coordenada": 'E' + str(i + 2), "Row": str(i + 2)}
                break
        if dic is None:
            raise ValueError(f"A conta {conta} não foi encontrada.")
        else:
            return dic
