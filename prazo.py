from tkinter import filedialog as fd
import pandas as pd

excelSaida = fd.askopenfilename()
excelEntrada = fd.askopenfilename()

for codigo in excelEntrada['Código']:

    dataEntrada = excelEntrada.loc[excelEntrada[excelEntrada['Código'] == codigo].index.values, 'Data'].values[0]
    dataSaida   = excelSaida.loc[excelSaida[excelSaida['Código'] == codigo].index.values, 'Data'].values[0]

    lojaEntrada = excelEntrada.loc[excelEntrada[excelEntrada['Código'] == codigo].index.values, 'Loja'].values[0]
    lojaSaida   = excelSaida.loc[excelSaida[excelSaida['Código'] == codigo].index.values, 'Loja'].values[0]

    quantidadeEntrada = excelEntrada.loc[excelEntrada[excelEntrada['Código'] == codigo].index.values, 'Quantidade'].values[0]
    quantidadeSaida   = excelSaida.loc[excelSaida[excelSaida['Código'] == codigo].index.values, 'Quantidade'].values[0]

    if dataEntrada >= dataSaida & quantidadeEntrada == quantidadeSaida & lojaEntrada == lojaSaida:     

        excelEntrada.at[i,'NUMERO CT-E']

codigoBarra = planDF.loc[planDF[planDF['Codigo'] == int(codigoFinal[2])].index.values, 'Barra'].values[0]