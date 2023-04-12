from openpyxl import Workbook
from openpyxl.drawing.image import Image 
from openpyxl.styles.alignment import Alignment

#criando arquivo 
wb = Workbook()

#selecionando a planilha atual
ws = wb.active

#mudando o titulo da folha atual
ws.title = "Portaria {}".format(2023)

#carregar imagem
img = Image("logo.png")
img.height = 75
img.width = 180

#adicionar imagem a celula A1 da planilha
ws.add_image(img, "A1")

#mesclar celulas para imagem e titulo
ws.merge_cells("A1:C4")
ws.merge_cells("D1:K4")
ws["D1"].alignment = Alignment(horizontal="center")

#meclar celulas cabeçalho e ajustar alinhamento
ws.merge_cells("A5:B5")
ws["A5"].alignment = Alignment(horizontal="center")
ws.merge_cells("C5:D5")
ws["C5"].alignment = Alignment(horizontal="center")
ws.merge_cells("E5:F5")
ws["E5"].alignment = Alignment(horizontal="center")
ws.merge_cells("G5:K5")
ws["G5"].alignment = Alignment(horizontal="center")

#cabeçalho e titulo
ws["D1"] = 'ESCALA PORTARIA {}'.format(2023)
ws["A5"] = 'DATA'
ws["C5"] = 'TURNO'
ws["E5"] = 'FUNCIONARIO'
ws["G5"] = 'OBSERVAÇÃO'

data = '01-01-{}'.format(2023)

for i in range(5,374):    
    ws["A{}".format(i)] = data
    data += data



#salvando arquivo
wb.save('Portaria_Escala_{}.xlsx'.format(2023))