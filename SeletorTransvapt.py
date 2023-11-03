import pandas as pd
from tkinter import filedialog as fd

finalDF = pd.DataFrame()

FilePath = fd.askopenfilename() 

colunasValidas = ["Nr.CT-e", "Unnamed: 2", "EMISSÃO", "CLIENTE REMETENTE", "CLIENTE DESTINATÁRIO", "CIDADE DESTINO", "FRETE TOTAL"]

tempDF = pd.read_excel(FilePath)
validade = tempDF.loc[1,'Unnamed: 16']
print(validade)

transvaptDF = pd.read_excel(FilePath, skiprows=4 ,skipfooter=9, usecols=colunasValidas)
transvaptDF['Unnamed: 2'] = transvaptDF['Unnamed: 2'].shift(-1)
transvaptDF[['lixo','NF']] = transvaptDF["Unnamed: 2"].str.split(" ", n=1, expand=True)
transvaptDF[['lixo','CTE']] = transvaptDF["Nr.CT-e"].str.split(" ", n=1, expand=True)
transvaptDF[['CIDADE','UF']] = transvaptDF["CIDADE DESTINO"].str.split("/", n=1, expand=True)
transvaptDF.dropna(how="all", inplace=True)

finalDF = transvaptDF.loc[:,["EMISSÃO", "CTE", "CLIENTE DESTINATÁRIO", 'NF', "CIDADE","UF", "FRETE TOTAL"]]
finalDF.insert(7,"Vencimento", validade)
finalDF.reset_index(inplace=True)

finalDF.to_csv('PlanilhaTotal.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
print("concluído!")