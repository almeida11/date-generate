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
        #DATAS ENTRADA
        dia = self.values['dia']
        mes = self.values['mes']
        ano = self.values['ano']
        #CRIANDO PLANILHA
        book = openpyxl.Workbook()
        data_page = book['Sheet']
        #GERADOR DE LINHAS COM DATAS
        for c in range(1, (int(dia)) + 1):
            data_page.append(['FOLHA DE SALÁRIOS',
            '',
            '',
            '',
            '',
            '22.089,00',
            f'{c}/{mes}/{ano}'
            ])
            data_page.append(['TRANSM. DE ENERGIA',
            '',
            '',
            '',
            '',
            '10.323,00',
            f'{c}/{mes}/{ano}'
            ])
            book.save('Planilha.xlsx')
tela = TelaPy()
tela.Iniciar()
