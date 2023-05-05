import pandas as pd
from docx import Document
from docx.shared import Inches

# Leer los datos de Excel con pandas
pastYear = pd.read_csv('pastYearCamarena.csv', sep=',')
nextYear=pd.read_csv('nextYearCamarena.csv', sep=',')
print (nextYear.head())