from contextlib import suppress
import pandas as pd
from tkinter import filedialog as fd

excelPath = fd.askopenfilename()

colunas = ["destNome", "destRazao", "destLogradouro", "destNumero", "destBairro", "destCidade", "destUF", "destCEP", "NFnum", "NFvalor", 
                "pedido", "CTE", "Mpeso", "Fpeso", "valor_COBRADO", "valor_CTE", "valorCobradoComprador"]


tempDF = pd.read_excel(excelPath, usecols=colunas)

finalDF = tempDF.loc[:,["destNome", "destRazao", "destLogradouro", "destNumero", "destBairro", "destCidade", "destUF", "destCEP", "NFnum", "NFvalor", 
                "pedido", "CTE", "Mpeso", "Fpeso", "valor_COBRADO", "valor_CTE", "valorCobradoComprador"]]

finalDF.insert(12,"Tipo", " ")
finalDF.reset_index(drop=True, inplace=True)
finalDF.to_excel('PlanilhaBitHome.xlsx',index=False, header=True)
print("concluído!")