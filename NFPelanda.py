import PyPDF2, re, os
import pandas as pd
from tkinter import filedialog as fd

dictFatura = {}
dadosFatura = {}
dadosFatura[3] = []
listaCaminhoes = ['HEM3C47','ATW5B14','AQI6E66']

arquivo = fd.askopenfilename()
arquivoPDF = PyPDF2.PdfReader(open(arquivo,'rb'))
endereco = os.path.split(arquivo)
os.chdir(endereco[0])

textoPaginaPDF = arquivoPDF.pages[0].extract_text()

numeroDoc      = re.findall(r'Faturamento (\d{8})',                             textoPaginaPDF) #numeroDoc
vencimentoDoc  = re.findall(r'Vencimento\n(\d{2}\/\d{2}\/\d{4})',               textoPaginaPDF) #vencimentoDoc
dadosFatura[0] = re.findall(r'MAESTRO (\d{2}\/\d{2}\/\d{2})',                   textoPaginaPDF) #faturaEmissoes
dadosFatura[1] = re.findall(r'MAESTRO \d{2}\/\d{2}\/\d{2} (\d+)',               textoPaginaPDF) #faturaNumeroCupons
dadosFatura[2] = re.findall(r'\d+,\d{2}\n(\d{5}) \d+\n',                        textoPaginaPDF) #faturaNFs
dadosFatura[7] = re.findall(r'\d+,\d{2}\n(\d+) \d,\d{2}|\d+,\d{2}\n( )\d,\d{2}',textoPaginaPDF) #faturaKM
dadosFatura[4] = re.findall(r'([A-Z]{3}\d{1}.\d{2}) \d+,\d{2}',                 textoPaginaPDF) #faturaPlacas
dadosFatura[5] = re.findall(r'[A-Z]{3}\d{1}.\d{2} (\d+,\d{2})',                 textoPaginaPDF) #faturaLitros
dadosFatura[6] = re.findall(r'[A-Z]{3}\d{1}.\d{2} \d+,\d{2}\n (.+,\d{2})\n',    textoPaginaPDF) #faturaValores
sequencia      = re.findall(r'\d+,\d{2}\n\d{5} (\d+)\n',                        textoPaginaPDF) #faturaSequencia


for i in range(len(dadosFatura[2])):
    dadosFatura[3].append(dadosFatura[7][i][0])

del dadosFatura[7]

x=0
while(True):
    try:
        if len(dadosFatura[x]) < len(dadosFatura[0]):
            dadosFatura[x].insert(0, "")
        x+=1
    
    except(KeyError):
        break

if len(sequencia) < len(dadosFatura[0]):
    sequencia.insert(0,1)

dataFrameBase = pd.DataFrame(data=dadosFatura, index=None)

dataFrameBase.columns       = ['KM','Emissão','Cupom','NF','Placa','Litros','Valor']
dataFrameBase['Documento']  = numeroDoc[0]
dataFrameBase['Vencimento'] = vencimentoDoc[0]
dataFrameBase["Sequencia"]  = sequencia

dataFrameOrdenado = dataFrameBase.loc[:,['Documento','Emissão','Sequencia','Cupom','NF','KM','Placa','Litros','Valor','Vencimento']]
dataFrameOrdenado['Sequencia']    = pd.to_numeric(dataFrameOrdenado['Sequencia'])
dataFrameOrdenado['Cupom']        = pd.to_numeric(dataFrameOrdenado['Cupom'])
dataFrameOrdenado['NF']           = pd.to_numeric(dataFrameOrdenado['NF'])
dataFrameOrdenado['KM']           = pd.to_numeric(dataFrameOrdenado['KM'])
dataFrameOrdenado['Documento']    = pd.to_numeric(dataFrameOrdenado['Documento'])

dataFrameCaminhoes = dataFrameOrdenado[(dataFrameOrdenado['Placa']=='ATW5B14') | (dataFrameOrdenado['Placa']=='HEM3C47') | (dataFrameOrdenado['Placa']=='AQI6E66')]
dataFrameCaminhoes.loc[len(dataFrameCaminhoes)] = pd.Series(dtype='float64')

dataFrameCarros = dataFrameOrdenado[(dataFrameOrdenado['Placa']!='ATW5B14') & (dataFrameOrdenado['Placa']!='HEM3C47') & (dataFrameOrdenado['Placa']!='AQI6E66')]

dataFrameFinal = pd.concat([dataFrameCaminhoes, dataFrameCarros], ignore_index=True)
dataFrameFinal.to_excel('excel.xlsx')