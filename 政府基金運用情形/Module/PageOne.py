import PyPDF2
#import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import re
import os

PDF_files  = []
input_path = r'input'
for x in os.listdir(input_path):
    if '.pdf' in x:
        PDF_files.append(x)

GFI  = pd.read_excel(r'input/202111 政府基金運用資訊 v2.0.xlsx',sheet_name = 'Government Fund Info (Raw Data)',skiprows = 2)

#把index處理一下
new_index = []
old_index = GFI.set_index('Unnamed: 0').index
for i in GFI.set_index('Unnamed: 0').index:
    #只抓前面的中文部分
    if re.match('[\u4e00-\u9fa5]+',i):
        new_index.append(re.match('[\u4e00-\u9fa5]+',i)[0])


        