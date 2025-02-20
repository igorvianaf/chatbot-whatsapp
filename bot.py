#lib para abrir planilhas
import openpyxl
#lib para formatar links para links de api
from urllib.parse import quote
#lib para criar tempo de espera
from time import sleep
#lib para abrir navegador
import webbrowser
#lib automatizar click
import pyautogui

# Ler planilha e guardar as informações
planilha_teste = openpyxl.load_workbook('agenda.xlsx')
pag_execucao = planilha_teste['Planilha1']

#iterar por todas as linhas da planilha, iniciando na linha 2
for linha in pag_execucao.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value
    medico = linha[3].value

    #Personalizar mensagem de envio\ Formatar data
    mensagem = f'Oi {nome}, você tem uma consulta agendada para a data: {data.strftime('%d/%m/%Y')}. Com o profissional {medico}'

#Link exemplo = https://web.whatsapp.com/send?phone=&text=

    try:
        """
        Criar links personalizados para enviar mensagens com base nos dados de cada cliente e usar o quote para formatar texto de mensagem.
        """
        link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        sleep(15)
        #abrir navegador
        webbrowser.open(link_mensagem)
        sleep(7)
        #enviar mensagem com o botão enter
        pyautogui.press('enter')
        sleep(9)
        pyautogui.hotkey('ctrl', 'w')
    except:
        #tratamento de erro
        print(f'Não foi possível enviar mensagem para {nome}')
        #criar arquivo csv com os contatos que não foi possivel enviar mensagem
        with open('erros.csv', 'a', newline='', encoding='utf8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')