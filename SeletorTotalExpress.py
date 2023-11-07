from contextlib import suppress
import pandas as pd
from tkinter import filedialog as fd
import os.path
import glob

finalDF = pd.DataFrame()
tempDF2 = pd.DataFrame()

folderPath = fd.askdirectory()
os.chdir(folderPath)
localArquivos = os.getcwd()
arquivosCSV = glob.glob(os.path.join(localArquivos, "*.csv"))


colunasCTE = ["Data Faturamento", "Fatura", "Nota Fiscal", "CTe", "Destinatario", "Cidade", 
                  "CEP", "UF", "Peso", "Valor NF", "Seguro", "Gris", "Frete", "ICMS", "Total Servico"]
colunasNFe = ["Data Faturamento", "NF", "Nota Fiscal", "Destinatario", "Cidade", 
                  "CEP", "UF", "Peso", "Valor NF", "Seguro", "Gris", "Frete", "Total Servico"]

for file in arquivosCSV:
    try:

        tempDF = pd.read_csv(file, encoding = "ISO-8859-1", skipfooter=1, sep=';', usecols=colunasCTE, engine='python')

    except ValueError:

        tempDF = pd.read_csv(file, encoding = "ISO-8859-1", skipfooter=1, sep=';', usecols=colunasNFe, engine='python')
        tempDF.insert(1,"Fatura", " ")
        tempDF.insert(13,"ICMS", 0)

    with suppress(KeyError,ValueError):
        tempDF['Fatura'] = tempDF['Fatura'].astype(int)
        tempDF['Nota Fiscal'] = tempDF['Nota Fiscal'].astype(int)
        tempDF['CTe'] = tempDF['CTe'].astype(int)
        tempDF['Data Faturamento'] = pd.to_datetime(tempDF['Data Faturamento'],format='%d%b%Y')

    tempDF2 = pd.concat([tempDF2, tempDF], ignore_index=False)


finalDF = tempDF2.loc[:,["Data Faturamento", "Fatura", "Nota Fiscal", "CTe", "Destinatario", "Cidade",
                        "CEP", "UF", "Peso", "Valor NF", "Seguro", "Gris", "Frete", "ICMS", "Total Servico"]]
finalDF.insert(1,"Vencimento", " ")
finalDF.insert(5,"Tipo", " ")
finalDF.reset_index(inplace=True)
finalDF = finalDF.drop(6)
finalDF.to_excel('PlanilhaTotal.xlsx',index=False, header=True)
print("concluído!")