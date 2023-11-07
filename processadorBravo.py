import pytesseract as pyt
import re
import os.path
import glob
import pandas as pd
import threading as thr
from tkinter import filedialog as fd
from PIL import Image
from pdf2image import convert_from_path

class Processador():

    def __init__(self, folderPath):
        self.listaNotas = {}
        self.arquivosPDF = {}
        self.arquivosJPG = {}
        self.poplerPath = "C:/Users/mladika/Documents/Python/Ladika/poppler/Library/Bin"
        self.nfsPath = folderPath + '/NFs'
        self.xlsxPath = folderPath + '/Excel'
        pyt.pytesseract.tesseract_cmd = r"C:\Users\mladika\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

    def conversor(self):

        os.chdir(self.nfsPath)
        localArquivos = os.getcwd()

        self.arquivosPDF = glob.glob(os.path.join(localArquivos, "*.pdf"))

        thread1 = thr.Thread(target = Processador.converterNotas(i = 0, numeroNota = 0 ))
        thread2 = thr.Thread(target = Processador.converterNotas(i = 1, numeroNota = 10))
        thread3 = thr.Thread(target = Processador.converterNotas(i = 2, numeroNota = 20))
        thread4 = thr.Thread(target = Processador.converterNotas(i = 3, numeroNota = 30))

        thread1.start()
        thread2.start()
        thread3.start()
        thread4.start()

        thread1.join()
        thread2.join()
        thread3.join()
        thread4.join()

        self.arquivosJPG = glob.glob(os.path.join(localArquivos, "*.jpg"))

        thread1 = thr.Thread(target = Processador.processarNotas(k = 0))
        thread2 = thr.Thread(target = Processador.processarNotas(k = 1))
        thread3 = thr.Thread(target = Processador.processarNotas(k = 2))
        thread4 = thr.Thread(target = Processador.processarNotas(k = 3))

        thread1.start()
        thread2.start()
        thread3.start()
        thread4.start()

        thread1.join()
        thread2.join()
        thread3.join()  
        thread4.join()

        dataFrame = pd.DataFrame(data = self.listaNotas, index=[0])
        dataFrame = (dataFrame.T)
        dataFrame.reset_index(inplace=True)
        dataFrame.columns = ['CTRC','NF']
        dataFrame['CTRC'] = pd.to_numeric(dataFrame['CTRC'])
        dataFrame['NF'] = pd.to_numeric(dataFrame['NF'])

        Processador.processarExcel(dataFrame)

    def converterNotas(self, i, numeroNota):

        while True:

            try:
                pages = convert_from_path(self.arquivosPDF[i], 350, poppler_path=self.popplerPath)
            except IndexError:
                break
            
            for page in pages:
                image_name = "NF" + str(numeroNota) + ".jpg"
                page.save(image_name, "JPEG")
                print("Nota " + str(numeroNota) + " convertida com sucesso")
                numeroNota+=1

            i+=4

    def processarNotas(self, k):

        while True:

            try:
                img = Image.open(self.arquivosJPG[k])
            except (FileNotFoundError, IndexError):
                break
            
            text = pyt.image_to_string(img, config='--psm 11')
            numeroCTRC = re.findall(r'n°. (\d{4})', text)
            numeroNF = re.findall(r'CURITIBA\n\n(\d{4})', text)

            if numeroCTRC and numeroNF:
                self.listaNotas.update({numeroCTRC[0]:numeroNF[0]})
                print("Nota " + str(k) + " processada com sucesso")
                k+=4

            else:
                print("algo deu errado aí, meu patrau")
                print(text)
                k+=4
                pass

    def processarExcel(self, dataFrameNotas):

        finalDF = pd.DataFrame()
        tempDF2 = pd.DataFrame()

        colunasValidas = ["NUMERO CTRC", "NUMERO CT-E", "PLACA DE COLETA", "CLIENTE REMETENTE",	"CLIENTE DESTINATARIO",	"ENDERECO RECEBEDOR", 
                        "CIDADE ENTREGA", "UF ENTREGA", "CEP ENTREGA", "NOTA FISCAL", "FRETE PESO", "OUTROS", "VAL RECEBER", 
                        "NUMERO DA FATURA", "DATA ENTREGA", "IMPOSTOS REPAS", "ICMS TRANSP"]

        os.chdir(self.xlsxPath)
        localArquivos = os.getcwd()
        arquivosCSV = glob.glob(os.path.join(localArquivos, "*.csv"))

        for f in arquivosCSV:

            tempDF = pd.read_csv(f, encoding = "ISO-8859-1", skiprows=[0], skipfooter=1, sep=';', usecols=colunasValidas, engine='python')
            tempDF2 = pd.concat([tempDF2, tempDF], ignore_index=False)

        finalDF = tempDF2.loc[:,["NUMERO DA FATURA", "NOTA FISCAL", "NUMERO CTRC", "NUMERO CT-E", "DATA ENTREGA",  
                                    "CLIENTE REMETENTE", "CLIENTE DESTINATARIO", "PLACA DE COLETA", "ENDERECO RECEBEDOR", "CIDADE ENTREGA", "UF ENTREGA", 
                                    "FRETE PESO", "OUTROS", "VAL RECEBER"]]

        finalDF.reset_index(inplace=True)

        for i in range(len(finalDF)):

            if int(finalDF.loc[i]['NUMERO CT-E']) == 0:
                
                numeroCTRCbruto = (finalDF.loc[i]['NUMERO CTRC'])
                numeroCTRC = re.findall(r'CWN00(\d{4})',numeroCTRCbruto)
                try:
                    finalDF.at[i,'NUMERO CT-E'] = dataFrameNotas.loc[
                        dataFrameNotas[dataFrameNotas['CTRC'] == int(numeroCTRC[0])].index.values, 'NF'].values[0]
                except IndexError:

                    print(numeroCTRC[0])

        finalDF.insert(5,"Filial", " ")
        finalDF.to_csv('PlanilhaBravo.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)
        print("concluído!")

if __name__ == "__main__":
    
    folderPath = fd.askdirectory()
    dataFrameNotas = Processador(folderPath)
    dataFrameNotas.conversor()
