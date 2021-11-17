from os import close
import pandas as pd
import openpyxl
import PySimpleGUI as pag

class TelaPy:
    def __init__(self):
        #Layout
        layout = [
            [pag.Text('Dias'), pag.Input(key='dia')],
            [pag.Text('Mês'), pag.Input(key='mes')],
            [pag.Text('Ano'), pag.Input(key='ano')],
            [pag.Button('Gerar Planilha')]
        ]
        #Janela
        janela = pag.Window("Dias do Custo").layout(layout)

        #Extrair os dados da tela
        self.button, self.values = janela.Read()

    def  Iniciar(self):
        dia = self.values['dia']
        mes = self.values['mes']
        ano = self.values['ano']
        #CRIANDO LINHAS NECESSÁRIAS
        book = openpyxl.Workbook()
        data_page = book['Sheet']
        book.save('Planilha.xlsx')
        #tentando escrever na planilha
        for rows in data_page.iter_rows(min_row=1, max_row=31):
            for cell in rows:
                cell.value = mes
tela = TelaPy()
tela.Iniciar()
