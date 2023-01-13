import pytesseract
from wand.image import Image
from tkinter import filedialog as fd
from pathlib import Path

filePath = fd.askopenfilename()
print(filePath)

path = Path(filePath)

all_text = []

text = pytesseract.image_to_string(Image.open(path.name))
all_text.append(text)

print(all_text)