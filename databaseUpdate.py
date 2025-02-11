import pandas as pd
import mysql.connector
from sqlalchemy import create_engine
from tkinter import filedialog as fd

engine      = create_engine('mysql+pymysql://Clarity:159753@192.168.0.13:3306/ladika')
filePath    = fd.askopenfilename()
dfBase      = pd.read_excel(io=filePath, dtype='object')

dfBase.to_sql(name='contratos', con=engine, index=False, if_exists='append')