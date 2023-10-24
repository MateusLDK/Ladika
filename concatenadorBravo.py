import pandas as pd
import os
import re
import glob
import NFBravo
from tkinter import filedialog as fd

folderPath = fd.askdirectory() 
nfsPath = folderPath + '/NFs'
xlsxPath = folderPath + '/Excel'

dataFrameNotas = NFBravo.converterNotas(nfsPath)

finalDF = pd.DataFrame()
tempDF2 = pd.DataFrame()
colunasValidas = ["NUMERO CTRC", "NUMERO CT-E", "PLACA DE COLETA", "CLIENTE REMETENTE",	"CLIENTE DESTINATARIO",	"ENDERECO RECEBEDOR", 
                  "CIDADE ENTREGA", "UF ENTREGA", "CEP ENTREGA", "NOTA FISCAL", "FRETE PESO", "OUTROS", "VAL RECEBER", 
                  "NUMERO DA FATURA", "DATA ENTREGA", "IMPOSTOS REPAS", "ICMS TRANSP"]


os.chdir(xlsxPath)
localArquivos = os.getcwd()
arquivosCSV = glob.glob(os.path.join(localArquivos, "*.csv"))

for f in arquivosCSV:
      
    tempDF = pd.read_csv(f, encoding = "ISO-8859-1", skiprows=[0], skipfooter=1, sep=';', usecols=colunasValidas, engine='python')
    tempDF2 = pd.concat([tempDF2, tempDF], ignore_index=False)

finalDF = tempDF2.loc[:,["NUMERO DA FATURA", "NOTA FISCAL", "NUMERO CTRC", "NUMERO CT-E", "DATA ENTREGA",  
                            "CLIENTE REMETENTE", "CLIENTE DESTINATARIO", "PLACA DE COLETA", "ENDERECO RECEBEDOR", "CIDADE ENTREGA", "UF ENTREGA", 
                            "FRETE PESO", "OUTROS", "VAL RECEBER"]]

finalDF.reset_index(inplace=True)

for i in range(len(finalDF)):

    if int(finalDF.loc[i]['NUMERO CT-E']) == 0:
        
        numeroCTRCbruto = (finalDF.loc[i]['NUMERO CTRC'])
        numeroCTRC = re.findall(r'CWN00(\d{4})',numeroCTRCbruto)
        try:

            finalDF.at[i,'NUMERO CT-E'] = dataFrameNotas.loc[dataFrameNotas[dataFrameNotas['CTRC'] == int(numeroCTRC[0])].index.values, 'NF'].values[0]
        except IndexError:

            print(numeroCTRC[0])

finalDF.to_csv('PlanilhaTotal.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)

print("concluído!")