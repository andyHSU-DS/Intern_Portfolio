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
    url  = r'http://www.sharpinvest.com/Product/OverView'
    path = r"D:\My Documents\kenc\桌面\ken_python\chromedriver.exe"                                  # - chromedriver

    driver = webdriver.Chrome(path)
    driver.get(url)

    href = driver.page_source
    return href

def result_table(href,精彩網_date):
    """
    get current web page result table df
    """
    
    soup = BeautifulSoup(href, 'html.parser')
    tables = soup.find('table', {'class': 'table'})
    columns = tables.find_all('th')
    cols = [column.text for column in columns] 
    result = pd.DataFrame(columns=cols)

    a=0
    row = []
    row_datas = tables.find_all("td")
    for row_data in row_datas : 

        if a == len(cols) and a != 0:
            row_list = pd.Series( row ,index = result.columns )
            result = result.append(row_list,ignore_index=True)
            row=[]
            a=0

        row.append(row_data.text.replace("\n","").replace(' ',""))        
        a+=1

    result['發行日期'] = pd.to_datetime(result['發行日期'])
    result.index = result['發行日期']
    result = result.drop(['發行日期'],axis=1)
    result = result.to_period("M")

    精彩網_date =  精彩網_date.replace("年","-")[:-1]
    result = result.reset_index()
    result = result[result['發行日期']==精彩網_date].reset_index(drop=True)

    return result


精彩網_date = str('2021 年 03 月')
href  = get_web_data(精彩網_date)
table = result_table(href,精彩網_date)
print(table)