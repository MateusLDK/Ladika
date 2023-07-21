from tkinter import filedialog as fd
from pdf2image import convert_from_path

#filePath = fd.askdirectory()
 
#for file in filePath:

images = convert_from_path(poppler_path = r"C:\Users\mladika\Documents\Python\poppler-0.68.0_x86\poppler-0.68.0\bin", pdf_path = r"X:\Relatorio de Despesas\Despesas PR\Bravo Log\2023\07-10\NFE - FAT - 2805")

for i in range(len(images)):

    images[i].save('page'+ str(i) +'.jpg', 'JPEG')