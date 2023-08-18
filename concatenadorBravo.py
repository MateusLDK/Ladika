import pandas as pd
import os
import glob
from tkinter import filedialog as fd
import time

finalDF = pd.DataFrame()
tempDF2 = pd.DataFrame()
colunasValidas = ["NUMERO CTRC", "NUMERO CT-E", "PLACA DE COLETA", "CLIENTE REMETENTE",	"CLIENTE DESTINATARIO",	"ENDERECO RECEBEDOR", 
                    "CIDADE ENTREGA", "UF ENTREGA", "CEP ENTREGA", "NOTA FISCAL", "FRETE PESO", "OUTROS", "VAL RECEBER", 
                    "NUMERO DA FATURA", "DATA ENTREGA", "IMPOSTOS REPAS", "ICMS TRANSP"]

filePath = fd.askdirectory()
timeStart = time.time()
os.chdir(filePath)
localArquivos = os.getcwd()
arquivosCSV = glob.glob(os.path.join(localArquivos, "*.csv"))

for f in arquivosCSV:
      
    tempDF = pd.read_csv(f, encoding = "ISO-8859-1", skiprows=[0], skipfooter=1, sep=';', usecols=colunasValidas, engine='python')
    tempDF2 = pd.concat([tempDF2, tempDF], ignore_index=False)

finalDF = tempDF2.loc[:,["NUMERO DA FATURA", "NOTA FISCAL", "NUMERO CTRC", "NUMERO CT-E", "DATA ENTREGA",  
                            "CLIENTE REMETENTE", "CLIENTE DESTINATARIO", "PLACA DE COLETA", "ENDERECO RECEBEDOR", "CIDADE ENTREGA", "UF ENTREGA", 
                            "FRETE PESO", "OUTROS", "VAL RECEBER"]]
finalDF.to_csv('PlanilhaTotal.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
timeEnd = time.time()

print("concluído!")

print("tempo de execução: ",end='')
print(timeEnd-timeStart)
print("Células finalizadas: ",end='')
print(len(finalDF)*len(finalDF.index))