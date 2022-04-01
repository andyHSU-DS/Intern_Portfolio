import warnings
import numpy as np
import pandas as pd

import bs4

import time

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys


path   = r'/Users/andyhsu/Desktop/野村/MoneyDJ/chromedriver'
driver = webdriver.Chrome(path)
driver.get('https://www.moneydj.com/')
#fixed_fund不要改，這是一個陰招，因為需要在第一次輸入基金資訊避開下拉式選單，但又不能輸入有()這類的特殊字元，所以用這個
def 搜尋標的(company = '瑞銀投信', fund = '瑞士銀行(瑞士)資產組合美元收益型(未申報生效)', fixed_fund = '瑞聯UBAM美國優質中期公司債券基金美元AC',\
    比較市場 = ['海外基金','海外基金'],比較公司 = ['瑞銀投信','瑞銀投信'],比較標的 = ['瑞士銀行(瑞士)歐洲機會基金(未申報生效)','瑞士銀行(盧森堡)歐元國家機會基金(未申報生效)'],\
    比較期間 = '三個月績效', 期間是否自訂 = False, 起始時間 = None, 結束時間 = None):

    #----main_code----#
    基金按鈕 = driver.find_element_by_xpath('//li[@name="fund"]')
    print(基金按鈕.text)
    基金按鈕.click()
    輸入搜尋 = driver.find_element_by_xpath('//input[@name="txt1search"]')
    輸入搜尋.send_keys(fixed_fund)
    搜尋按鈕 = driver.find_element_by_xpath('//input[@value="搜尋" and @name="btn1"]')
    搜尋按鈕.click()
    time.sleep(3)
    company_select = Select(driver.find_element_by_xpath('//select[@name="selFund_corp"]'))
    company_select.select_by_visible_text(company)
    fund_select    = Select(driver.find_element_by_xpath('//select[@name="selFund3"]'))
    fund_select.select_by_visible_text(fund)
    time.sleep(2)
    各種指標 = driver.find_elements_by_xpath('//div[@class = "Tab-Much"]/ul/li')
    for 指標 in 各種指標:
        print(指標.text)
        if 指標.text == '績效':
            指標.click()
            break

    選擇比較標的及區間 = driver.find_element_by_xpath('//div[@id = "yp012001Mark"]')
    選擇比較標的及區間.click()
    
    比較組合 = list(zip(比較市場,比較公司,比較標的))

    for x in range(len(比較組合)):
        市場選擇選單 = Select(driver.find_element_by_xpath('//select[@name="oFund_area"]'))
        市場選擇選單.select_by_visible_text(比較組合[x][0])
        公司選擇選單 = Select(driver.find_element_by_xpath('//select[@name="oFund_corp"]'))
        公司選擇選單.select_by_visible_text(比較組合[x][1])
        基金選擇選單 = Select(driver.find_element_by_xpath('//select[@name="oFund3"]'))
        基金選擇選單.select_by_visible_text(比較組合[x][2])
        增加按鈕    = driver.find_element_by_xpath('//a[@href="javascript:addFID();"]')
        增加按鈕.click()
    
    if 期間是否自訂 != True:
        績效區間選擇清單 = Select(driver.find_element_by_xpath('//select[@name="selYEAR"]'))
        績效區間選擇清單.select_by_visible_text(比較期間)

    else:
        全部時間       = []
        自由選擇區間按鈕 = driver.find_element_by_xpath('//input[@name = "radioMonth" and @onclick = "radioSelYear_onclick(false)"]')
        自由選擇區間按鈕.click()
        全部時間.extend(起始時間)
        全部時間.extend(結束時間)
        #兄節點
        按鈕 = Select(driver.find_element_by_xpath('//select[@name="Y2"]'))
        按鈕.select_by_visible_text(全部時間[0])
        for index,時間 in enumerate(全部時間[1:],start=1):
            #離當前節點最近的第x節點
            按鈕 = Select(driver.find_element_by_xpath('//select[@name="Y2"]/following-sibling::select['+str(index)+']'))
            按鈕.select_by_visible_text(時間)
            
        


    開始比較按鈕 = driver.find_element_by_xpath('//img[@src="/funddj/images/Extend/BT_Start.gif"]')
    開始比較按鈕.click()



if __name__ == '__main__':
    搜尋標的(期間是否自訂 = True, 起始時間 = ['2022','1','1'], 結束時間 = ['2022','1','11'])