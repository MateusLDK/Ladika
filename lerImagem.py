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


def lerArquivos():

    i = 1
    for arquivo in arquivosPDF:

        pages = convert_from_path(arquivo, 350, poppler_path=popplerPath)

        for page in pages:
            image_name = "NF" + str(i) + ".jpg"
            page.save(image_name, "JPEG")
            print("Nota " + str(i) + " convertida com sucesso")
            i+=1

def processarArquivos():

    
    i=1
    print(len(arquivosPDF))
    while i <= len(arquivosPDF):
        sleep(1)
        img = Image.open(r"C:\Users\mladika\Documents\Python\Ladika\NFs\NF" + str(i) + ".jpg")
        pytesseract.pytesseract.tesseract_cmd = tesseractPath
        text = pytesseract.image_to_string(img)

        numeroCTRC = re.findall(r'n°. (\d{4})', text)
        numeroNF = re.findall(r'CURITIBA (\d{4})', text)

        if numeroCTRC and numeroNF:
            listaNotas[numeroCTRC[0]] = numeroNF[0]
            print("Nota " + str(i) + " processada com sucesso")
            i+=1
        else:
            print("algo deu errado aí, meu patrau")

if __name__ == "__main__":

    listaNotas = {}
    arquivosPDF = {}
    popplerPath = "C:/Users/mladika/Documents/Python/Ladika/poppler/Library/Bin"
    tesseractPath = r"C:\Users\mladika\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

    filePath = fd.askdirectory()
    os.chdir(filePath)
    localArquivos = os.getcwd()
    arquivosPDF = glob.glob(os.path.join(localArquivos, "*.pdf"))

    thread1 = threading.Thread(target = lerArquivos)
    thread2 = threading.Thread(target = processarArquivos)

    thread1.start()
    sleep(6)
    thread2.start()

    thread1.join()
    thread2.join()

    dataFrame = pd.DataFrame(data=listaNotas, index=[0])
    dataFrame = (dataFrame.T)
    print(dataFrame)
    dataFrame.to_excel('teste.xlsx')