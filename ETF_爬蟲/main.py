from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import re
import pandas as pd


def get_one_page(url='https://etfdb.com/etfs/industry/biotechnology/'):
    ETFS_list = []
    driverpath='/Users/andyhsu/Desktop/chromedriver'
    options = Options()
    # --start是給window用的
    options.add_argument('--kiosk')
    browser=webdriver.Chrome(driverpath, chrome_options = options)
    #找到url
    browser.get(url)
    #找到tbody
    body = browser.find_element_by_xpath('//*[@id="etfs"]/tbody')
    #找到tbody內每一個tag_name是tr的（代表每個ETF）
    Each_ETFS = body.find_elements_by_tag_name('tr')
    for ETF in Each_ETFS:
        Each_ETF = []
        #每個ETF內的資訊，但我只要class是overview開頭的
        etf_informations = ETF.find_elements_by_tag_name('td')
        for etf_info in etf_informations:
            if re.match('overview',etf_info.get_attribute('class')):
                #print(info.get_attribute('class'))
                Each_ETF.append(etf_info.text)
        ETFS_list.append(Each_ETF)
    #關閉所有頁面（完全關閉）
    #close關閉當前頁面
    browser.close() 

    final_df = pd.DataFrame(ETFS_list).iloc[:,:8]
    final_df.columns = ['Symbol','ETF Name','Sector','Total Assets($MM)','YTD','Avg Volume','Previous Closing Price','1-Day Change']
    return final_df

#用來抓全部資料的
####有分成industry跟sector
def get_ETF_information(kind = 'industry',kind_2 = 'biotechnology',pages = 0):
    if kind == 'industry':
        url = 'https://etfdb.com/etfs/industry/'+kind_2+'/#etfs&sort_name=assets_under_management&sort_order=desc&page='
    elif kind == 'sector':
        url = 'https://etfdb.com/etfs/sector/'+kind_2+'/#etfs&sort_name=assets_under_management&sort_order=desc&page'
    if pages == 0:
        url = url+str(1)
        result = get_one_page(url)
    else:
        total_df = pd.DataFrame({})
        for p in range(1,pages+1):
            print(p)
            url = url+str(p)
            temp_table = get_one_page(url)
            total_df = total_df.append(temp_table)
        total_df.reset_index(inplace=True,drop=True)
        total_df.columns = ['Symbol','ETF Name','Sector','Total Assets($MM)','YTD','Avg Volume','Previous Closing Price','1-Day Change']
        result = total_df
    result.to_csv('output/'+kind_2.capitalize()+'_ETF_information.csv',index=False,header=True,encoding='utf-8-sig')
    return result

if __name__ == '__main__':
    get_ETF_information()
    get_ETF_information('sector','healthcare',3)
    get_ETF_information('industry','pharmaceutical')