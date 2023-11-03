import pandas as pd
from tkinter import filedialog as fd

finalDF = pd.DataFrame()
listaErros = []
listaLojas = [1,2,4,5,6,7,8,10,11,12,14,19]
arquivoNotas = fd.askopenfilename()
arquivoSaidas = fd.askopenfilename() 

dfNotas = pd.read_csv(arquivoNotas, encoding = "ISO-8859-1", sep=';', engine='python')
dfSaidas = pd.read_excel(arquivoSaidas)

for i in dfNotas['produto_comprado_key']:

    try:
        envio = dfSaidas.loc[dfSaidas[dfSaidas['Codigo'] == int(i)].index.values, 'Data'].values[0]
        dfNotas.at[i,'Data_envio'] = envio

    except IndexError:
        listaErros.append(i)

dfListaErros = pd.DataFrame(listaErros)
dfNotas.to_csv('notas.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
dfListaErros.to_csv('notasErros.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
print("concluído!")