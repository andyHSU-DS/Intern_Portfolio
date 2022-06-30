import os
import functools
import numpy as np 
import pandas as pd 
from code1 import main_fnc as go
from code1 import append_excel


class ken_AML():
    def __init__(self,excel_file_path):
        self.path=excel_file_path
    def start(self):
        df=go(self.path)
        append_excel(excel_path=self.path,df=df)
        print(df)
        os.rename(r"C:\Users\kenc\Music\Ken_Chiang\input_excel_file\TXT_M_#15R_20201231_1.xlsx",r"C:\Users\kenc\Music\Ken_Chiang\input_excel_file\modified_TXT_M_#15R_20201231_1.xlsx")
        print("----------------------------ok!-------------------------------------")
    


excel_file_path=r"C:\Users\kenc\Music\Ken_Chiang\input_excel_file\TXT_M_#15R_20201231_1.xlsx"
work1=ken_AML(excel_file_path)
work1.start()
