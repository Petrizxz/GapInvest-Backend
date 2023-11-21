from __future__ import print_function

import os

from flask import Flask, request, jsonify, make_response
from flask_restful import Resource, Api
from werkzeug.utils import secure_filename
from flask_cors import CORS, cross_origin

from Controller import Planilha

UPLOAD_FOLDER = 'folder'
ALLOWED_EXTENSIONS = set(['xlsx'])

app = Flask(__name__)
CORS(app)
api = Api(app)
app.config['CORS_HEADERS'] = 'Content-Type'
app.config['folder'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
# TODO: fazer isso caso precise Aqui estão as funções para cadastrar o extrato da Xp(Estão incompleto)
# No finalzinho da logica, tem que olhar o fato de comparar todas as datas iguais
# Exemplo: 15/09 apareceu 4 vezes no extrato porem na planilha so tem 2, tem q fazer uma logica pra conferir isso
# da maneira que esta agr msm se falta ele vai var se a data é > que a da planilha, se for começa a inserir apartir
# do dia 16/09 sem conferir se ficou algo para trás
# Nome do extrato
# Intervalo dos dados de result
# Nome das páginas
# Posição da conta
# planilha = load_workbook("Novo Resultado Robos ONLINE.xlsx", data_only=True)


# def verificar_cliente(extrato_xp, clientes):
#     nome = extrato_xp["B5"].value
#     conta = int(extrato_xp["G5"].value.replace("Conta XP: ", ""))
#
#     for celula in clientes["E"]:
#
#         if celula.value == conta:
#             linha = clientes[f"A{celula.row}:F{celula.row}"]
#             dic = {"Nome": linha[0][0].value, "Corretora": linha[0][1].value, "Estrategia": linha[0][2].value,
#                    "Bolsa": linha[0][3].value, "Conta": linha[0][4].value, "Codinome": linha[0][5].value,
#                    "Coordenada": celula.coordinate}
#             return dic
#
#     print(f"Cliente '{nome}' não encontrado. Conta Xp: {conta}")
#     return None
#
#
# def procurar_linha_extrato_xp(ultima_linha_dados_xp, extrato_ativa):
#     for linha in extrato_ativa[f"A1:O{extrato_ativa.max_row}"]:
#
#         if (linha[0].value == ultima_linha_dados_xp[0][0].value and
#                 linha[1].value == ultima_linha_dados_xp[0][1].value and
#                 linha[2].value == ultima_linha_dados_xp[0][2].value and
#                 linha[3].value == ultima_linha_dados_xp[0][3].value and
#                 linha[4].value == ultima_linha_dados_xp[0][4].value and
#                 linha[5].value == ultima_linha_dados_xp[0][5].value):
#             # i = linha[0].row - 2
#             #  tabela = []
#             # while i > 15:
#             #   inserir = []
#             #  row = extrato_ativa[f"B{i}:G{i}"]
#             # for j in range(0, 6):
#             #    inserir.append(linha[0][j].value)
#             #   tabela.append(inserir)
#             #
#             #  del inserir
#             # i = i - 1
#             # Menos 2 porque é - a linha em si e menos dnv para voltar a celula
#             # for k in tabela:
#             #   print(k)
#             # print(tabela)
#             print(linha[0].row - 2)
#             return linha[0].row - 2
#         return "Deu ruiom"
#
#
# def ultima_linha_dadosxp(dadosxp, cliente):
#     numero = 0
#     for celula in dadosxp["E"]:
#
#         if celula.value == cliente["Conta"]:
#             # data = dadosxp[f"G{celula.row}"]
#             numero = celula.row
#
#     if numero == 0:
#         return 0
#     else:
#         print(numero)
#         return dadosxp[f"F{numero}:K{numero}"]
# def get_tabelacliente(dados, cliente):
#     tabela = []
#     for celula in dados["E"]:
#         inserir = []
#         if celula.value == cliente["Conta"]:
#             linha = dados[f"E{celula.row}:K{celula.row}"]
#             for j in range(0, 6):
#                 inserir.append(linha[0][j].value)
#                 tabela.append(inserir)
#                 del inserir
#     return tabela
#
#
# def get_extrato(dados, conta):
#     tabela = []
#     for celula in dados["E"]:
#         inserir = []
#         if celula.value == conta:
#             linha = dados[f"E{celula.row}:K{celula.row}"]
#             for j in range(0, 6):
#                 inserir.append(linha[0][j].value)
#             tabela.append(inserir)
#             del inserir
#     return tabela
#
#
# def ultima_linha(extrato_ativa):
#     entro = False
#     for celula in extrato_ativa["B"]:
#         if type(celula.value) == datetime or entro:
#             entro = True
#             if celula.row == 158:
#                 print("chegou")
#             if celula.value is None:
#                 return extrato_ativa[f"B{celula.row - 1}:G{celula.row - 1}"]

def allowed_file(filename):
    # '.' in file verifica se tem ponto no arquivo.
    # Rsplit separa a extensão do nome deixa minusculo e compara se a estensão é valida no vetor ALLOWED_EXTENSIONS
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


class UploadFiles(Resource):
    def post(self):
        # Se não enviar um arquivo
        if 'file' not in request.files:
            return make_response(jsonify({"message": "Arquivo obrigatório!"}), 500)

        file = request.files['file']  # confere extensão
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['folder'], filename))

            print(f"Arquivo {filename} recebido com sucesso!")
            print(f"Arquivo salvo no caminho: .{os.path.join(app.config['folder'], filename)}")

            try:
                Planilha.cadastrar_activ(filename)
            except ValueError as erro:
                print('vai remover arquivo')
                os.remove(f'./folder/{filename}')
                print('arquivo removido com sucesso!')
                return make_response(jsonify({"message": f"{erro}"}), 400)
            else:
                os.remove(f'./folder/{filename}')
                return make_response(jsonify({"message": "Upload realizado com sucesso!"}), 200)


api.add_resource(UploadFiles, '/')
# api.add_resource(UserById, '/users/<id>')

if __name__ == '__main__':
    # Planilha.cadastrar_activ(f"lion.xlsx")
    app.run(host="0.0.0.0", port=5000, debug=True)


