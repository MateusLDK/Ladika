import PyPDF2
import re
import pandas as pd

from tkinter import filedialog as fd

dictFatura = {}
dadosFatura = {}
valorTemp = {}

arquivoPDF = PyPDF2.PdfReader(open(fd.askopenfilename(),'rb'))
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
dataFrame.insert(0,"Documento",'')
dataFrame.insert(2,"Sequencia",sequencia)
dataFrame.insert(9,"Vencimento",'')

dataFrame.columns = ['Documento','Emissão','Sequencia','Cupom','NF','KM','Placa','Litros','Valor','Vencimento']

dataFrame['Sequencia'] = pd.to_numeric(dataFrame['Sequencia'])
dataFrame['Cupom'] = pd.to_numeric(dataFrame['Cupom'])
dataFrame['NF'] = pd.to_numeric(dataFrame['NF'])
dataFrame['KM'] = pd.to_numeric(dataFrame['KM'])

dataFrame.to_excel('teste.xlsx')
