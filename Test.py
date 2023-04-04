import pandas as pd
from docx import Document
from docx.shared import Inches

# Leer los datos de Excel con pandas
pastYear = pd.read_csv('Past.csv', sep=',')
nextYear=pd.read_csv('Programming.csv', sep=',')

print (nextYear.iloc[0,9])