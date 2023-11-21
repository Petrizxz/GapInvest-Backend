def lotsSets(idPlanilha, LocTabela, sheet):

    valores_adicionar = [5]
    valores_adicionar[0] = input("insira o valor de Advanced B3")
    valores_adicionar[1] = '-'
    valores_adicionar[2] = '-'
    valores_adicionar[3] = input("insira o valor de Índices + Moedas Activ")
    valores_adicionar[4] = input("insira o valor de Madalena B3")

    result = sheet.values().update(spreadsheetId=idPlanilha,
                                   range="I5", valueInputOption="USER_ENTERED",
                                   body={'values': valores_adicionar}).execute()

    result = sheet.values().get(spreadsheetId=idPlanilha,
                                range=LocTabela).execute()

    values = result.get('values', [])

    for row in values:
        if row[0] == '-':
            row[0] = row[0].replace('-', '0')
        if row[1] == '-':
            row[1] = row[1].replace('-', '0')

        print(int(row[0].strip('%')), int(row[1].strip('%')))

    message = client.messages.create(
        to="+553398755165",
        from_="+14067098800",
        body="FOI ENVIADO UM SMS PARA O CELULAR")

    print(message.sid)
