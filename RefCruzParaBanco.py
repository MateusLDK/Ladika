import pandas as pd 
from tkinter import filedialog as fd

folderPath = fd.askopenfilename()

dataFrame = pd.read_excel(folderPath)
colunasEncontradas = []
listaLojas=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,201,202,203,204,205,206,207,208,209,210,602,301,302,303,304,305,306,307,308,309,310]
colunasPlanilha = list(dataFrame.columns)

for coluna in colunasPlanilha:

    if coluna in listaLojas:
        colunasEncontradas.append(coluna)

for loja in colunasEncontradas:
    colunasPlanilha.remove(loja)

dataFrameArrumado = dataFrame.melt(
    id_vars=colunasPlanilha, value_vars=colunasEncontradas, var_name="Loja", value_name="Minimo")

dataFrameArrumado.to_excel(excel_writer="planilha arrumada.xlsx", index = False)
print("feito!")
