"""
Preciso automatizar minhas mensagens para clientes gostaria de saber valores, e gostaria de que entrassem em contato comigo para explicar melhor, quero poder mandar mensagens de cobrança em determinado dia com clientes com vencimento difente.
"""
#lib para abrir planilhas
import openpyxl
#lib para formatar links para links de api
from urllib.parse import quote
#lib para criar tempo de espera
from time import sleep
#lib para abrir navegador
import webbrowser

webbrowser.open('https://web.whatsapp.com/')
sleep(20)

# Ler planilha e guardar as informações
planilha_teste = openpyxl.load_workbook('clientes.xlsx')
pag_execucao = planilha_teste['Planilha1']

#iterar por todas as linhas da planilha, iniciando na linha 2
for linha in pag_execucao.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value

    #Personalizar mensagem de envio\ Formatar data
    mensagem = f'Oi {nome}, seu boleto irá vencer em {data.strftime('%d/%m/%Y')}'

#Link exemplo = https://web.whatsapp.com/send?phone=&text=
#Criar links personalizados para enviar mensagens com base nos dados de cada cliente.
link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

webbrowser.open(link_mensagem)
