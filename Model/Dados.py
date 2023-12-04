from uteis.constantes.const import C

class Dados:
    def __init__(self, sheet, sheet_id, conta):
        dados = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_NAME()).execute()
        self.tab = dados['values']
        # _formatar_tab(dados)
        self.ultlinha = self._procurar_ultima_linha_dadosactiv(conta)
        self.conta = conta
        self.max_row = len(dados['values']) + 1

    def _procurar_ultima_linha_dadosactiv(self, conta):
        numero = 0
        for i in range(len(self.tab)):
            try:
                if self.tab[i][4] == conta:
                    numero = i
            except Exception as err:
                print(err)  # TODO Aqui é onde acha o erro caso tenha alguma linha em branco

        if numero == 0:
            return "Cadastrar"
        else:
            print(numero)
            print(self.tab[numero][6])
            return self.tab[numero]

    @staticmethod
    def get_dolar(sheet, sheet_id):
        try:
            dados = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_DOLAR()).execute()
            dados = dados['values']
            dados = {
                "tab": dados,
                "insert_row": ''
            }
        except ValueError as erro:
            raise ValueError(f"Não foi possível buscar os dados da planilha clientes. Erro: {erro}")
        else:
            return dados

    @staticmethod
    def get_b3(sheet, sheet_id):
        try:
            print("VAmos Buscar os dados antigos")
            dados = sheet.values().get(spreadsheetId=sheet_id, range=C.RANGE_B3()).execute()
            dados = dados['values']
            dados = {
                "tab": dados,
                "insert_row": ''
            }
        except ValueError as erro:
            raise ValueError(f"Não foi possível buscar os dados da planilha clientes. Erro: {erro}")
        else:
            return dados
