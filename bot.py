"""
Preciso automatizar minhas mensagens para clientes gostaria de saber valores, e gostaria de que entrassem em contato comigo para explicar melhor, quero poder mandar mensagens de cobrança em determinado dia com clientes com vencimento difente.
"""

import openpyxl
# Ler planilha e guardar as informações
planilha_teste = openpyxl.load_workbook('clientes.xlsx')
pag_execucao = planilha_teste['Planilha1']

#iterar por todas as linhas da planilha, iniciando na linha 2
for linha in pag_execucao.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value
    print(nome)
    print(telefone)
    print(data)
