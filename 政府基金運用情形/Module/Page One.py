import PyPDF2
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import re
import os

PDF_files  = []
input_path = r'input'
for x in os.listdir(input_path):
    if '.pdf' in x:
        PDF_files.append(x)

        