from __future__ import print_function

import os

from flask import Flask, request, jsonify, make_response, send_from_directory
from werkzeug.utils import secure_filename
from flask_cors import CORS

from uteis.constantes.const import C

from Controller import Planilha

app = Flask(__name__)
CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
app.config['folder'] = C.UPLOAD_FOLDER()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


# TODO: fazer isso caso precise Aqui estão as funções para cadastrar o extrato da Xp(Estão incompleto)
# No finalzinho da logica, tem que olhar o fato de comparar todas as datas iguais
# Exemplo: 15/09 apareceu 4 vezes no extrato porem na planilha so tem 2, tem q fazer uma logica pra conferir isso
# da maneira que esta agr msm se falta ele vai var se a data é > que a da planilha, se for começa a inserir a partir
# do dia 16/09 sem conferir se ficou algo para trás
# Nome do extrato
# Intervalo dos dados de result
# Nome das páginas
# Posição da conta
# planilha = load_workbook("Novo Resultado Robôs ONLINE.xlsx", data_only=True)


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
#         return "Deu ruim"
#
#
# def ultima_linha_dados_xp(dados_xp, cliente):
#     numero = 0
#     for celula in dados_xp["E"]:
#
#         if celula.value == cliente["Conta"]:
#             # data = dados_xp[f"G{celula.row}"]
#             numero = celula.row
#
#     if numero == 0:
#         return 0
#     else:
#         print(numero)
#         return dados_xp[f"F{numero}:K{numero}"]
# def get_tabela_cliente(dados, cliente):
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

def allowed_file(filename, ALLOWED_EXTENSIONS):
    # '.' in file verifica se tem ponto no arquivo.
    # Rsplit separa a extensão do nome deixa minusculo e compara se a extensão é valida no vetor ALLOWED_EXTENSIONS
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS()


def buscar_nome_arquivo(rota):
    vet = []
    for caminho, subpasta, arquivos in os.walk(rota):
        for nome in arquivos:
            vet.append(nome)
    return vet


@app.route('/upload-sheet/', methods=['POST'])
def upload():
    # Se não enviar um arquivo
    if 'file' not in request.files:
        return make_response(jsonify({"message": "Arquivo obrigatório!"}), 500)

    file = request.files['file']  # confere extensão
    if file and allowed_file(file.filename, C.ALLOWED_EXTENSIONS_SHEET):
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
    else:
        return make_response(jsonify({"message": f"Tipo de arquivo incorreto, envie uma planilha Excel!"}), 400)

@app.route('/upload-img/', methods=['POST'])
def upload_img():

    if 'file' not in request.files:
        return make_response(jsonify({"message": "Arquivo obrigatório!"}), 500)

    file = request.files['file']
    if file and allowed_file(file.filename, C.ALLOWED_EXTENSIONS_IMG):

        try:
            clientes = Planilha.get_cliente()
            clientes = list(map(lambda cliente: cliente[0].lower(), clientes["tab"]))
        except ValueError as erro:
            return make_response(jsonify({"message": f"{erro}"}), 400)
        else:
            filename = secure_filename(file.filename)
            nome_arquivo = ""
            for extensions in C.ALLOWED_EXTENSIONS_IMG():
                nome_arquivo = filename.removesuffix("." + extensions).lower()
                if nome_arquivo != filename.lower():
                    break

            if nome_arquivo in clientes:
                print(f"app.config['folder']: {app.config['folder']}")
                file.save(os.path.join(app.config['folder'], filename))
                print(f"caminho: {os.path.join(app.config['folder'], filename)}")
                return make_response(jsonify({"message": "Upload realizado com sucesso!"}), 200)
            else:
                return make_response(jsonify({"message": "Nome do Cliente não foi encontrado na Planilha!"}), 400)
    else:
        return make_response(jsonify({"message": f"Tipo de arquivo incorreto, envie um arquivo de midia!"}), 400)

@app.route('/get-img/<nome>', methods=['GET'])
def get_img(nome):
    arquivos = buscar_nome_arquivo(C.UPLOAD_FOLDER())

    for arquivo in arquivos:
        for extensions in C.ALLOWED_EXTENSIONS_IMG():
            if nome.lower() == arquivo.removesuffix("." + extensions).lower():
                return send_from_directory(C.UPLOAD_FOLDER(), arquivo, as_attachment=False)

    return make_response(jsonify({"message": "Nome do Cliente não foi encontrado nos Upload!"}), 400)


#################################################################  POST  ######################################################


@app.route('/planilha', methods=['POST'])
def post_cliente():
    try:
        Planilha.cadastrar_cliente(request.json)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return make_response(jsonify({"message": "Cliente cadastrado com sucesso!"}), 200)


@app.route('/dolar', methods=['POST'])
def post_dolar():
    try:
        Planilha.cadastrar_dolar(request.json)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return make_response(jsonify({"message": "Cotação atualizada com sucesso!"}), 200)


@app.route('/b3', methods=['POST'])
def post_b3():
    try:
        Planilha.cadastrar_b3(request.json)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return make_response(jsonify({"message": "Cotação atualizada com sucesso!"}), 200)


@app.route('/dados-cliente', methods=['POST'])
def post_dados_cliente():
    try:
        Planilha.cadastrar_dados_cliente(request.json)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return make_response(jsonify({"message": "Dados atualizados com sucesso!"}), 200)


@app.route('/operacional', methods=['POST'])
def post_operacional():
    try:
        Planilha.cadastrar_operacional(request.json)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return make_response(jsonify({"message": "Planilha atualizada com sucesso!"}), 200)


#################################################################  GET  ######################################################


@app.route('/planilha', methods=['GET'])
def get_cliente():
    try:
        print("foi feito um get para cliente")
        dados = Planilha.get_cliente()
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        # TODO Talvez esse 0 possa mudar caso mude a planilha clientes TALVEZ ISSO N FUNCIONE TEM Q TESTAR COM O POSTMAN
        clientes = []
        for i, linha in enumerate(dados['tab']):
            vet = {"Codinome": linha[0], "Coordenadas": dados['coordenadas'][i]}
            clientes.append(vet)

        # TODO TALVEZ NÂO SEJA ASSIM Q RETORNA JSON
        return jsonify(clientes)


@app.route('/planilha/<conta>/<codinome>', methods=['GET'])
def get_verifica_cliente(conta, codinome):
    try:
        print("foi feito um get para verificar cliente")
        cliente = Planilha.get_verifica_cliente(conta, codinome)
    except ValueError as erro:
        return make_response(jsonify({"message": f"{erro}"}), 400)
    else:
        return jsonify(cliente["Codinome"])


if __name__ == '__main__':
    # Planilha.cadastrar_activ(f"lion.xlsx")
    app.run(host="0.0.0.0", port=8080, debug=False)
    # Planilha.cadastrar_cliente(f"lion.xlsx")
    # Planilha.cadastrar_dolar(f"lion.xlsx")
    # Planilha.cadastrar_b3(f"lion.xlsx")
    # Planilha.get_cliente(f"lion.xlsx")
    # Planilha.cadastrar_dados_cliente(f"lion.xlsx")
    # Planilha.cadastrar_operacional(f"lion.xlsx")
    # get_cliente()
    # Planilha.get_verifica_cliente()
    # upload_img()
