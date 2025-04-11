import pandas as pd
from tkinter import filedialog as fd

finalDF = pd.DataFrame()
listaErros = []
listaCodigos = []
listaLojas = [1,2,4,5,6,7,8,10,11,12,14,19]
arquivoNotas = fd.askopenfilename()
arquivoSaidas = fd.askopenfilename() 

dfEntradas = pd.read_csv(arquivoNotas, encoding = "ISO-8859-1", sep=';', engine='python')
dfSaidas = pd.read_excel(arquivoSaidas)
dfEntradas['data_entrada'] = pd.to_datetime(dfEntradas['data_entrada'], dayfirst=True)
dfSaidas['Data'] = pd.to_datetime(dfSaidas['Data'], dayfirst=True)

for idx, codigo in enumerate(dfEntradas['produto_comprado_key']):

    listaIndex = dfSaidas.loc[dfSaidas['Codigo'] == codigo].index.tolist()
    k=0
    for i in listaIndex:
        cabecalho = "data_" + str(k)

        if dfSaidas.loc[i,'Data'] >= dfEntradas.loc[idx, 'data_entrada']:

            print()
            print(dfSaidas.loc[i,'Data'])
            print(dfSaidas.loc[listaIndex.index(i).value-1,'Data'])

            if dfSaidas.loc[i,'Data'] != dfSaidas.loc[i-1,'Data']:
                if cabecalho in dfEntradas:
                    dfEntradas.loc[idx,cabecalho] = dfSaidas.loc[i,'Data']
                    k+=1

                else:
                    dfEntradas.insert(len(dfEntradas.columns), cabecalho, "")
                    dfEntradas.loc[idx,cabecalho] = dfSaidas.loc[i,'Data']
                    k+=1

            else: pass
        else: pass

dfEntradas.to_excel('teste1.xlsx',index=False, header=True)
