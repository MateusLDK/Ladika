import pandas as pd 
from tkinter import filedialog as fd

folderPath = fd.askopenfilename()

dataFrame = pd.read_excel(folderPath)
listaLojas=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,201,202,203,204,205,206,207,208,209,210,602,301,302,303,304,305,306,307,308,309,310]
colunasPlanilha = list(dataFrame.columns)

for coluna in colunasPlanilha:

    if coluna in str(listaLojas):
        colunasPlanilha.remove(coluna)
    
dataFrameArrumado = dataFrame.melt(
    id_vars=[colunasPlanilha], var_name="Lojas", value_name="Quantidade")

print(dataFrameArrumado)
dataFrameArrumado.to_excel(excel_writer="planilha arrumada.xlsx")
