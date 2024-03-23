import pandas as pd
import os
import os.path
import glob
from tkinter import filedialog as fd

tempDF2 = pd.DataFrame()
finalDF = pd.DataFrame()
colunasValidas = ["Nr.CT-e", "Unnamed: 2", "EMISSÃO", "CLIENTE REMETENTE", "CLIENTE DESTINATÁRIO", "CIDADE DESTINO", "FRETE TOTAL"]

folderPath = fd.askdirectory()
os.chdir(folderPath)
localArquivos = os.getcwd()
arquivosXLS = glob.glob(os.path.join(localArquivos, "*.xls"))

for file in arquivosXLS:

    transvaptDF = pd.read_excel(file)
    validade = transvaptDF.loc[1,'Unnamed: 16']

    transvaptDF = pd.read_excel(file, skiprows=4 ,skipfooter=9, usecols=colunasValidas)
    transvaptDF['Unnamed: 2'] = transvaptDF['Unnamed: 2'].shift(-1)
    transvaptDF[['lixo','NF']] = transvaptDF["Unnamed: 2"].str.split(" ", n=1, expand=True)
    transvaptDF[['lixo','CTE']] = transvaptDF["Nr.CT-e"].str.split(" ", n=1, expand=True)
    transvaptDF[['CIDADE','UF']] = transvaptDF["CIDADE DESTINO"].str.split("/", n=1, expand=True)
    transvaptDF.dropna(how="all", inplace=True)

    transvaptDF.reset_index(inplace=True)
    tempDF2 = pd.concat([tempDF2, transvaptDF], ignore_index=True)


finalDF = tempDF2.loc[:,["EMISSÃO", "CTE", "CLIENTE DESTINATÁRIO", 'NF', "CIDADE","UF", "FRETE TOTAL"]]
finalDF.insert(7,"Vencimento", validade)
finalDF.to_excel('PlanilhaTransvapt.xlsx', index=False, header=True)
print("concluído!")
