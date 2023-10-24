import pandas as pd
from tkinter import filedialog as fd

finalDF = pd.DataFrame()

#Abrir o seletor de arquivo do windows
FilePath = fd.askopenfilename() 

#colunas que serão carregadas
colunasValidas = ["Data Faturamento", "Fatura", "Nota Fiscal", "CTe", "Destinatario", "Cidade", 
                  "CEP", "UF", "Peso", "Valor NF", "Seguro", "Gris", "Frete", "ICMS", "Total Servico"]

#ler arquivo CSV: Ignora a ultima linha e abre apenas as colunas selecionadas acima
tempDF = pd.read_csv(FilePath, encoding = "ISO-8859-1", skipfooter=1, sep=';', usecols=colunasValidas, engine='python')

#definindo as colunas numericas como INT para remover o .0 do final
tempDF['Fatura']      = tempDF['Fatura'].astype(int)
tempDF['Nota Fiscal'] = tempDF['Nota Fiscal'].astype(int)
tempDF['CTe']         = tempDF['CTe'].astype(int)

#Reordenando as colunas
finalDF = tempDF.loc[:,["Data Faturamento", "Fatura", "Nota Fiscal", "CTe", "Destinatario", "Cidade",
                          "CEP", "UF", "Peso", "Valor NF", "Seguro", "Gris", "Frete", "ICMS", "Total Servico"]]

#adicionando duas colunas vazias na posição 1 e 5.
finalDF.insert(1,"Vencimento", " ")
finalDF.insert(5,"Tipo", " ")

#resetando o index após as alterações
finalDF.reset_index(inplace=True)

#adicionando o DataFrame em uma planilha CSV nova
finalDF.to_csv('PlanilhaTotal.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
print("concluído!")