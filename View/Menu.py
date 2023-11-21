from scrapy.spidermiddlewares.httperror import HttpError
from Controller import Planilha


def menu(filename):
    try:
        # cnt = 1
        # while cnt != 0:
        # planilha = input('Qual planilha você deseja editar?
        # Aperte "r" para subir extrato para Resultados Robos ou "o" para editar operacional robos')
        escolha = 'a'
        match escolha:
            case 'o':
                print("operadcional")
                # aba = input('Qual aba você deseja editar? Aperte "l" para Lots Sets ou "b" para BTs')
                # if aba == 'l':
                # operacional.lotsSets(operacioal_robos, lots_sets, sheet)
            case 'r':
                print("XP")
                #
                # try:
                #     # Vou pegar o extrato na xp
                #     extrato_ativa = extrato.active
                #
                #     # Vou pegar a aba dadosXp
                #     planilha.active = planilha['Dados XP']
                #     dadosxp = planilha.active
                #
                #     # Vou pegar a aba clientes
                #     planilha.active = planilha['Clientes']
                #     dadosclientes = planilha.active
                # except Exception as erro:
                #     print(f"Não foi possivel encontrar as planilhas ou as abas \n Erro: {erro}")
                # else:
                #     # Vou verificar se o cliente existe para poder cadastrar os dados
                #     cliente = verificar_cliente(extrato_ativa, dadosclientes)
                #     if cliente is not None:
                #         # Verifico no Resultado Robos qual foi a ultima_data_inserida
                #         # Se os ultimalinha_dados = 0, não tem dados registrados ainda
                #         # dadoscliente = get_tabelacliente(dadosxp, cliente)
                #         ultimalinha_dados = ultima_linha_dadosxp(dadosxp, cliente)
                #
                #         # Vou no extrato e procuro a ultima_data_inserida B15 porque o padrao vem do ultimo pro primeiro
                #         ultima_linha_extrato = extrato_ativa["B15:G15"]
                #         # Dadosextratos = get_tabelaextrato(extrato_ativa, cliente)
                #         # Provavelmente tem que comparar todos os campos, pois podem ter um que finaliza com todos
                #         if ultimalinha_dados[0][1].value > ultima_linha_extrato[0][1].value:
                #             print("Esse extrato ja foi cadastrado!")
                #         else:
                #             # retorna o número
                #             inserir_apartir = procurar_linha_extrato_xp(ultimalinha_dados, extrato_ativa)
                #             # Senão ela for maior pode quer dizer que tem datas faltando:
                #             # Mostrar um alert com o intervalor entre as datas UltResultado e UltExtrato
                #             # Se a data do último registro no resultado robos for maior que a da primeira linha
                #             # do extrato o extrato é antigo
                #             # Se a data do último registro no resultado rbo
                #             linha_max_dadosxp = dadosxp.max_row + 1
                #
                #             for linha in extrato_ativa[f"B{15}:G{inserir_apartir}"]:
                #                 coluna = 6
                #                 if inserir_apartir == 302:
                #                     print("Quase la")
                #                 if linha[0].value is None:
                #                     break
                #                 # dadosxp.cell(row=linha_max_dadosxp, column=1,
                #                 #             value=f"=clientes!{cliente['Coordenada']}")
                #                 # dadosxp.cell(row=linha_max_dadosxp, column=2,
                #                 #             value=f"=clientes!{info_cliente[1].coordinate}")
                #                 # dadosxp.cell(row=linha_max_dadosxp, column=3,
                #                 #              value=f"=clientes!{info_cliente[2].coordinate}")
                #                 # dadosxp.cell(row=linha_max_dadosxp, column=4,
                #                 #              value=f"=clientes!{info_cliente[3].coordinate}")
                #                 # dadosxp.cell(row=linha_max_dadosxp, column=5,
                #                 #              value=f"=clientes!{info_cliente[4].coordinate}")
                #
                #                 for celula in linha:
                #                     dadosxp.cell(row=linha_max_dadosxp, column=coluna, value=celula.value)
                #                     coluna = coluna + 1
                #
                #                 linha_max_dadosxp = linha_max_dadosxp + 1
                #
                #             planilha.save("ProdutosOpenPy.xlsx")
            case 'a':
                print("Activ")
                Planilha.cadastrar_activ(filename)
            case _:
                print("Escolha invalida! Verifique Letras maiusculas e minusculas.")
        #   if input('Você deseja sair? Aperte "s" para sim ou "n" para não') == "s":
        #      cnt = 0

    except ValueError as err:
        print("Vai voltar para main")
        raise err
