RANGE_NAME = 'Dados Activ!A2:S'


class Dados:
    def __init__(self, sheet, sheet_id, conta):
        dados = sheet.values().get(spreadsheetId=sheet_id, range=RANGE_NAME).execute()
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
                print(err) #TODO Aqui é onde acha o erro caso tenha alguma linha em branco

        if numero == 0:
            return "Cadastrar"
        else:
            print(numero)
            print(self.tab[numero][6])
            return self.tab[numero]
