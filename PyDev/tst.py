import openpyxl
book = openpyxl.Workbook()
data_page = book['Sheet']

data_page.append(['FOLHA DE SALÁRIOS', "", "", "", "", '22.089,00',])
data_page.append(['TRANSM. DE ENERGIA', "", "", "", "", '10.323,00',])

book.save('Planilha.xlsx')