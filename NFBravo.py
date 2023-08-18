import pytesseract
import re
import os.path
import glob
import pandas as pd
import threading

from time import sleep
from PIL import Image
from pdf2image import convert_from_path
from tkinter import filedialog as fd


def lerArquivos(i, numeroNota):

    while True:

        try:
            pages = convert_from_path(arquivosPDF[i], 350, poppler_path=popplerPath)
        except IndexError:
            break
        
        for page in pages:
            image_name = "NF" + str(numeroNota) + ".jpg"
            page.save(image_name, "JPEG")
            print("Nota " + str(numeroNota) + " convertida com sucesso")
            numeroNota+=1

        i+=4

def processarArquivos(k):

    while True:

        try:
            img = Image.open(arquivosJPG[k])
        except (FileNotFoundError, IndexError):
            break
        
        text = pytesseract.image_to_string(img, config='--psm 11')

        numeroCTRC = re.findall(r'n°. (\d{4})', text)
        numeroNF = re.findall(r'CURITIBA\n\n(\d{4})', text)

        if numeroCTRC and numeroNF:
            listaNotas.update({numeroCTRC[0]:numeroNF[0]})
            print("Nota " + str(k) + " processada com sucesso")
            k+=4

        else:
            print("algo deu errado aí, meu patrau")
            print(text)
            k+=4
            pass

if __name__ == "__main__":

    listaNotas = {}
    arquivosPDF = {}
    popplerPath = "C:/Users/mladika/Documents/Python/Ladika/poppler/Library/Bin"

    pytesseract.pytesseract.tesseract_cmd = r"C:\Users\mladika\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

    filePath = fd.askdirectory()
    os.chdir(filePath)
    localArquivos = os.getcwd()


    arquivosPDF = glob.glob(os.path.join(localArquivos, "*.pdf"))

    thread1 = threading.Thread(target = lerArquivos(i = 0, numeroNota = 0 ))
    thread2 = threading.Thread(target = lerArquivos(i = 1, numeroNota = 10))
    thread3 = threading.Thread(target = lerArquivos(i = 2, numeroNota = 20))
    thread4 = threading.Thread(target = lerArquivos(i = 3, numeroNota = 30))

    thread1.start()
    thread2.start()
    thread3.start()
    thread4.start()

    thread1.join()
    thread2.join()
    thread3.join()
    thread4.join()

    arquivosJPG = glob.glob(os.path.join(localArquivos, "*.jpg"))

    thread1 = threading.Thread(target = processarArquivos(k = 0))
    thread2 = threading.Thread(target = processarArquivos(k = 1))
    thread3 = threading.Thread(target = processarArquivos(k = 2))
    thread4 = threading.Thread(target = processarArquivos(k = 3))

    thread1.start()
    thread2.start()
    thread3.start()
    thread4.start()

    thread1.join()
    thread2.join()
    thread3.join()  
    thread4.join()

    dataFrame = pd.DataFrame(data = listaNotas, index=[0])
    dataFrame = (dataFrame.T)
    dataFrame.reset_index(inplace=True)
    dataFrame.columns = ['CTRC','NF']
    dataFrame['CTRC'] = pd.to_numeric(dataFrame['CTRC'])
    dataFrame['NF'] = pd.to_numeric(dataFrame['NF'])

    dataFrame.to_csv('PlanilhaNotas.csv', encoding = "ISO-8859-1", sep=';', index=False, header=True)