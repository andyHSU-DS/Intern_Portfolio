# basic
import re
import os 
import pandas as pd 
import numpy as np 

# bs4
from bs4 import element 
from bs4 import BeautifulSoup

# selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait



def get_web_data(date):
    url  = r'https://www.sitca.org.tw/ROC/Industry/IN4001.aspx?PGMID=IN0401'
    path = r"D:\My Documents\kenc\桌面\ken_python\chromedriver.exe"                                  # - chromedriver

    driver = webdriver.Chrome(path)
    driver.get(url)

    select = Select(driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_ddlYYYYMM'))
    select.select_by_visible_text(date)

    href = driver.page_source
    return href

def result_table(href):
    """
    get current web page result table df
    """
    
    soup = BeautifulSoup(href, 'html.parser')
    tables = soup.find_all('table', {'id': 'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_TableMEMA1'})
    
    texts = tables[0].find_all('td')
    col = ['公司代號','公司名稱','契約數量','全體有效契約金額','投資型保單有效契約-數量(新台幣)','投資型保單有效契約-金額(新台幣)','投資型保單有效契約-數量(外幣)','投資型保單有效契約-金額(外幣)']
    output_df = pd.DataFrame(columns=col)
    check_row = len(col)

    a=0
    row_data=[]
    for i in range(10,len(texts)):
        
        if a==0 and len(row_data) !=0 :
            row_list = pd.Series( row_data ,index = output_df.columns )
            output_df = output_df.append(row_list,ignore_index=True)
            row_data = []

        row_data.append(texts[i].text)
        a += 1 

        if a == check_row :
            a = 0 
    
    return output_df

date = str('2021 年 06 月')
href = get_web_data(date=date)
table = result_table(href)

print(table)