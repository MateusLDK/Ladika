import PyPDF2, re, os
import pandas as pd
from tkinter import filedialog as fd

dictFatura = {}
dadosFatura = {}
valorTemp = {}
listaCaminhoes = ['HEM3C47','ATW5B14','AQI6E66']

arquivo = fd.askopenfilename()
arquivoPDF = PyPDF2.PdfReader(open(arquivo,'rb'))
endereco = os.path.split(arquivo)
os.chdir(endereco[0])

textoPaginaPDF = arquivoPDF.pages[0].extract_text()

numeroDoc      = re.findall(r'Faturamento (\d{8})',                     textoPaginaPDF) #numeroDoc
vencimentoDoc  = re.findall(r'Vencimento\n(\d{2}\/\d{2}\/\d{4})',       textoPaginaPDF) #vencimentoDoc
dadosFatura[0] = re.findall(r'MAESTRO (\d{2}\/\d{2}\/\d{2})',           textoPaginaPDF) #faturaEmissoes
dadosFatura[1] = re.findall(r'MAESTRO \d{2}\/\d{2}\/\d{2} (\d{7})',     textoPaginaPDF) #faturaNumeroCupons
dadosFatura[2] = re.findall(r'\d+,\d{2}\n(\d{5}) \d\n',                 textoPaginaPDF) #faturaNFs
dadosFatura[3] = re.findall(r'\d+,\d{2}\n(\d+) \d,\d{2}',               textoPaginaPDF) #faturaKM
dadosFatura[4] = re.findall(r'([A-Z]{3}\d{1}.\d{2}) \d+,\d{2}',         textoPaginaPDF) #faturaPlacas
dadosFatura[5] = re.findall(r'(\d+,\d{2})\n .+,\d{2}\n\d+ \d,\d{2}',    textoPaginaPDF) #faturaLitros
dadosFatura[6] = re.findall(r'\d+,\d{2}\n (.+,\d{2})\n\d+ \d,\d{2}',    textoPaginaPDF) #faturaValores
sequencia      = re.findall(r'\d+,\d{2}\n\d{5} (\d{1})\n',              textoPaginaPDF) #faturaSequencia

dataFrame = pd.DataFrame(data=dadosFatura, index=None)

dataFrame.columns       = ['Emissão','Cupom','NF','KM','Placa','Litros','Valor']
dataFrame['Documento']  = numeroDoc[0]
dataFrame['Vencimento'] = vencimentoDoc[0]
dataFrame["Sequencia"]  = sequencia

finalDF = dataFrame.loc[:,['Documento','Emissão','Sequencia','Cupom','NF','KM','Placa','Litros','Valor','Vencimento']]

finalDF['Sequencia']    = pd.to_numeric(finalDF['Sequencia'])
finalDF['Cupom']        = pd.to_numeric(finalDF['Cupom'])
finalDF['NF']           = pd.to_numeric(finalDF['NF'])
finalDF['KM']           = pd.to_numeric(finalDF['KM'])
finalDF['Documento']    = pd.to_numeric(finalDF['Documento'])

dfCaminhoes = finalDF[(finalDF['Placa']=='ATW5B14') | (finalDF['Placa']=='HEM3C47') | (finalDF['Placa']=='AQI6E66')]
dfCarros    = finalDF[(finalDF['Placa']!='ATW5B14') & (finalDF['Placa']!='HEM3C47') & (finalDF['Placa']!='AQI6E66')]

dfCaminhoes.to_excel('Caminhoes.xlsx')
dfCarros.to_excel('Outros.xlsx')