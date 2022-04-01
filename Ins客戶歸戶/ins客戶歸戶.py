#677 693 743是每次要改的
"""

使用者路徑在 --> if __name__ == '__main__': 這個底下改 

    Input_File_Case_Path = r"L:\Cross_Dept_Shared\AML&CFT\DB\Ken_Chiang\ins客戶歸戶\input"
    Selenium_Path        = r"D:\My Documents\kenc\桌面\ken_python\ins客戶歸戶\chromedriver.exe"


Note :
 
    (1.) 這個跑完
    (2.) 手動更新 2021_Month_ILP 裡面 的 Top_49 全委帳戶.xlsx 裡面 Account 的 Top 5 Holding 
    
    (3.) AIA_DB_Mapping.py

        (3-1.) 將 (1.) 跑完的 onshore and offshore output 放到 Mapping_Data 裡的 DB_Data
        (3-2.) 將 更新完的  Top_49 全委帳戶.xlsx 的路徑放到 [AIA_DB_Mapping.py] 的相對位置
        (3-3.) 檢查 有無少 Mapping or 錯誤 Mapping 
    
    (4.) Account Analysis.py

        (4-1.) 將 更新完的  Top_49 全委帳戶.xlsx 的路徑放到 [Account Analysis.py] 的相對位置
        (4-2.) onshore/offshore 分開跑一次得到 帳戶裡面 , 基金的表現情況
        (4-3.) 確認是否有名子相似的沒被 Groupby 起來 , 需要相加
    



1. ------------------------------- Data/Code Position------------------------------------------------------- 

ken chiang 資料夾裡的 -->  ins 客戶歸戶

input  資料夾 路徑 : L:\Cross_Dept_Shared\AML&CTF\DB\Ken_Chiang\ins 客戶統計\input
output 資料夾 路徑 : L:\Cross_Dept_Shared\AML&CTF\DB\Ken_Chiang\ins 客戶統計\output

2. ------------------------------- program logic ----------------------------------------------------------

-->  中間 classify 的 fnc 為歸戶邏輯 function 

--> (1.) 客戶歸戶邏輯 : 設定規則 
    
    --> 全委 , 委託 , 管理帳戶 , 受託保管 , 基金專戶 , 等關鍵字去歸檔客戶 

    [ 南山人壽委託復華投信 --> 客戶歸檔(復華投信)  , 元大投信受託保管xxx --> 客戶歸檔(元大投信) ] --> 邏輯可以問 Tony

    -->  最後方法 --> ( 是當設定規則都失效時, 才會跑到這邊的 conditional criteria ) --> 直接找投信公司去歸戶名子

--> (2.) 基金歸戶邏輯 : 設定規則

    -->  xx人壽保險股份有限公司         --> Lifer 
    -->  xx產物保險股份有限公司         --> Insurance 
    -->  人壽 , (委託/全權委託) , 投信  --> ILP
    -->  組合基金專戶 , 多重資產組合    --> FoF 


-->  輸出

"""


# basic
import os
import warnings
import numpy as np
import pandas as pd 

# excel
from module import Excel_Work

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

warnings.filterwarnings("ignore")


# -------------- 政府 web data  ----------------------------------------------------

def transport(value):
    #取代，後轉為float
    value = value.replace(",","")
    value = float(value)
    return value

def get_web_data(date,path):
    #爬蟲
    url  = r'https://www.sitca.org.tw/ROC/Industry/IN4001.aspx?PGMID=IN0401'
    path = path                                 # - chromedriver

    driver = webdriver.Chrome(path)
    driver.get(url)

    #下拉式選單至目前月份
    select = Select(driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_ddlYYYYMM'))
    select.select_by_visible_text(date)
    #回傳url
    href = driver.page_source
    return href

def result_table(href):
    #讀取網頁資料並回傳onshore資料及offshore資料
    """
    get current web page result table df
    """
    
    soup    = BeautifulSoup(href, 'html.parser')
    #投信(委任關係)-依公司別   的表格
    tables = soup.find_all('table', {'id': 'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_TableMEMA1'})
    #表格內的所有文字
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

    #把新台幣計價的數量及金額讀出來
    onshore_output_df = output_df[['公司名稱','契約數量','全體有效契約金額','投資型保單有效契約-數量(新台幣)','投資型保單有效契約-金額(新台幣)']]

    #使用transport函數轉換資料
    onshore_output_df['投資型保單有效契約-金額(新台幣)'] = onshore_output_df.apply(lambda x: transport(x['投資型保單有效契約-金額(新台幣)']),axis=1)
    onshore_output_df['全體有效契約金額'] = onshore_output_df.apply(lambda x: transport(x['全體有效契約金額']),axis=1)
    onshore_output_df['契約數量'] = onshore_output_df.apply(lambda x: transport(x['契約數量']),axis=1)
    #大到小
    onshore_output_df = onshore_output_df.sort_values(by="投資型保單有效契約-金額(新台幣)",ascending=False)
    onshore_output_df = onshore_output_df.reset_index(drop=True)

    #把外台幣計價的數量及金額讀出來
    offshore_output_df = output_df[['公司名稱','契約數量','全體有效契約金額','投資型保單有效契約-數量(外幣)','投資型保單有效契約-金額(外幣)']]

    #使用transport函數轉換資料
    offshore_output_df['投資型保單有效契約-金額(外幣)'] = offshore_output_df.apply(lambda x: transport(x['投資型保單有效契約-金額(外幣)']),axis=1)
    offshore_output_df['全體有效契約金額'] = offshore_output_df.apply(lambda x: transport(x['全體有效契約金額']),axis=1)
    offshore_output_df['契約數量'] = offshore_output_df.apply(lambda x: transport(x['契約數量']),axis=1)
    #大到小
    offshore_output_df = offshore_output_df.sort_values(by="投資型保單有效契約-金額(外幣)",ascending=False)
    offshore_output_df = offshore_output_df.reset_index(drop=True)

    return onshore_output_df , offshore_output_df

# --------------- 精彩網---------------------------------------------------------------

def 精彩網_get_web_data(date,path):
    #回傳精彩網的網址
    url  = r'http://www.sharpinvest.com/Product/OverView'
    path = path                                  # - chromedriver

    driver = webdriver.Chrome(path)
    driver.get(url)

    href = driver.page_source
    return href

def 精彩網_result_table(href,date):
    #輸入網址及日期
    """
    get current web page result table df
    """

    soup = BeautifulSoup(href, 'html.parser')
    #只有一個表格
    tables = soup.find('table', {'class': 'table'})
    #找column name
    columns = tables.find_all('th')
    cols = [column.text for column in columns] 
    result = pd.DataFrame(columns=cols)

    a=0
    row = []
    #每個row的資料
    row_datas = tables.find_all("td")
    for row_data in row_datas : 

        if a == len(cols) and a != 0:
            row_list = pd.Series( row ,index = result.columns )
            result = result.append(row_list,ignore_index=True)
            row=[]
            a=0
        #將\n及' '取代
        row.append(row_data.text.replace("\n","").replace(' ',""))        
        a+=1
    #轉換時間類型
    result['發行日期'] = pd.to_datetime(result['發行日期'])
    #將發行日期變成索引
    result.index = result['發行日期']
    result = result.drop(['發行日期'],axis=1)
    #將索引變成月為單位
    result = result.to_period("M")

    #將輸入的參數(date)變成類似2022-1的組成
    精彩網_date =  date.replace("年","-")[:-1]
    result = result.reset_index()
    #等於要找這個月新發行的基金
    result = result[result['發行日期']==精彩網_date].reset_index(drop=True)

    return result


# ---------------- 基金歸戶 , 客戶歸戶 function -----------------------------------------

def classify(month,df): 
#將處理過後轉換好的客戶名稱放在客戶歸戶內
    df["客戶歸戶"]=""
    
    for i in range(df.shape[0]):
        cn = df["客戶姓名"][i]
        print(i,cn)

        #------------------------------------------------- 設定規則 -----------------------------------------
        if ("全權委託" in cn) and ("－" not in cn) and ("管理帳戶" not in cn) :
            index  = cn.index("全權委託")
            _index = cn.index("投信")
            cn1 = cn[ index + 4 : _index + 2 ]
            df["客戶歸戶"][i]=cn1


        # ---------- 保德信好時債
        elif ("保德信好時債" in cn ):
            cn1=cn[:3]+str("投信")
            df["客戶歸戶"][i]=cn1
        
        elif ("富蘭克林全球債券組合基金" in cn) :
            cn1=cn[:4]+str("華美投信")
            df["客戶歸戶"][i]=cn1
        
        # --------- 2021/06/16 處理  "-中國人壽" 問題 ---------------

        elif ("-中國人壽" in cn and "富蘭克林華美" in cn) :
            cn1 = '富蘭克林華美投信'
            df["客戶歸戶"][i] = cn1
            # print(cn1)

        elif ('-中國人壽' in cn ) :
            cn1 = str(cn[:2] + "投信")
            df['客戶歸戶'][i] = cn1
            # print(cn1)
        #-----------------------------------------------------------

        elif "管理帳戶" in cn : 
            index = cn.index("管理帳戶")
            cn1 = cn [index+5:]
            df["客戶歸戶"][i]=cn1

        elif ("全權委託" in cn) and ("－"  in cn) and ("管理帳戶" not in cn) :
            index  = cn.index("全權委託")
            _index = cn.index("投信")
            cn1 = cn[ index + 4 : _index + 2 ]
            df["客戶歸戶"][i]=cn1

        elif ("委託" in cn ):
            index  = cn.index("委託")
            _index  = cn.index("投信")
            cn1 = cn[index+2:_index+2]
            df["客戶歸戶"][i]=cn1

        elif("受託保管" in cn):
            index= cn.index("受託保管")
            cn1 = cn[index+4:]
            cn1 = cn1[:2]+"投信"
            df["客戶歸戶"][i]=cn1

        elif("受託" in cn):
            index = cn.index("銀行")
            cn1 = cn[:index+2]
            df["客戶歸戶"][i]=cn1

        elif ("日盛目標" in cn ):
            index=cn.index("日盛")
            cn1=cn[0:index+2]+str("投信")
            df["客戶歸戶"][i]=cn1

        elif  ("基金專戶" in cn  and "-" not in cn) :
            cn1=cn[:2]+str("投信")
            df["客戶歸戶"][i]=cn1

        elif ("投資信託基金" in cn and "-" not in cn):
            cn1=cn[:2]+str("投信")
            df["客戶歸戶"][i]=cn1

        elif ("-" in cn) :
            index = cn.index("-")
            cn1=cn[index+1:]
            if ("富蘭克林" in cn1 ):
                cn1 = cn1[:4]+"華美投信"
                df["客戶歸戶"][i]=cn1
            else:
                cn1 = cn1[:2]+"投信"
                df["客戶歸戶"][i]=cn1
        
        #-------------------------------2021 / 06 / 16----------------------------------------------------------

        # elif ("保險股份有限公司" in cn and "產物" not in cn ):
        #     index = cn.index("保險股份有限公司")
        #     cn1 = cn[:index]
        #     df["客戶歸戶"][i]=cn1
        #     print(cn1)

        # elif ("保險事業股份有限公司" in cn):
        #     index = cn.index("保險事業股份有限公司")
        #     cn1 = cn[:index]
        #     df["客戶歸戶"][i]=cn1
        #     print(cn1)
        
        # elif ("股份有限公司" in cn):
        #     index = cn.index("股份有限公司")
        #     cn1 = cn[:index]
        #     df["客戶歸戶"][i]=cn1
        #     print(cn1)

        #-------------------------------------------------直接找名子最後方法-----------------------------------------
        
        elif ("群益" in cn ):
            index = cn.index("群益")
            cn1=cn[index:2]+str("投信")
            df["客戶歸戶"][i]=cn1
        elif ("復華" in cn ):
            index = cn.index("復華")
            cn1=cn[index:2]+str("投信")
            df["客戶歸戶"][i]=cn1
        elif ("合庫" in cn ):
            index = cn.index("合庫")
            cn1=cn[index:2]+str("投信")
            df["客戶歸戶"][i]=cn1

        elif ('法國巴黎人壽澳幣環球穩健投資帳戶' in cn) :
            index = cn.index("法國巴黎人壽")
            cn1=cn[:index+6]+str("股份有限公司")
            df["客戶歸戶"][i] = cn1

        elif ('富蘭克林新興趨勢傘型基金之積極回報債券組合基金' in cn):

            df["客戶歸戶"][i] = "富蘭克林華美投信"
        
        #-----------------------------------------------------------------------------------------------------------
        else : 
            df["客戶歸戶"][i]=cn
            # print(cn)
            
    # 其實在 PROGRAMMING 裡面 只有在讀 "客戶姓名" 去做條件式歸檔 ,  所以其他欄位的資料 是不會變的 ! 
    if "onshore" in month:
        df = df [['通路', '通路名稱', '客戶','客戶歸戶' , '客戶姓名', '股債別','交易-申購總額','交易-買回金額(匯出+轉申購)','交易-淨流入(申購總額-買回匯出-買回轉申)','月底AUM','月平均AUM','結存-月平均AUM(起迄月份合計後平均)']]  
    else:
        df = df [['通路', '通路名稱', '客戶','客戶歸戶' , '客戶姓名','基金公司', '股債別','交易-申購總額','交易-買回金額(匯出+轉申購)','交易-淨流入(申購總額-買回匯出-買回轉申)','月底AUM','月平均AUM','結存-月平均AUM(起迄月份合計後平均)']]  
    
    return df 

def classify_fund(month,df): 
#將處理過後轉換好的基金名稱放在基金歸戶內   
    df["基金歸戶"]=""
    for i in range(df.shape[0]):
        cn = df["客戶姓名"][i]

        if ("人壽保險股份有限公司" in cn or "人壽保險事業股份有限公司" in cn):
            
            df["基金歸戶"][i] = 'Lifer'
        
        elif ("產物保險股份有限公司" in cn) :

            df["基金歸戶"][i] = 'Insurance'
        
        elif ("人壽" in cn and "全權委託" in cn ) :

            df['基金歸戶'][i] ="ILP"
        
        elif ("組合証券" in cn  or "組合證券" in cn):

            df["基金歸戶"][i] = 'FoF'
        
        elif ("人壽" in cn and "委託" in cn and "投信" in cn):

            df["基金歸戶"][i] = 'ILP'

        elif ("組合基金" in cn):

            df["基金歸戶"][i] = 'FoF'

        elif ("多重資產基金專戶" in cn ) : 

            df['基金歸戶'][i] = "FoF"

        else:
            df["基金歸戶"][i] = 'Others'
    
    if "onshore" in month:
        df = df [['通路', '通路名稱', '客戶','客戶歸戶' , '客戶姓名',"基金歸戶", '股債別','交易-申購總額','交易-買回金額(匯出+轉申購)','交易-淨流入(申購總額-買回匯出-買回轉申)','月底AUM','月平均AUM','結存-月平均AUM(起迄月份合計後平均)']]  
    else:
        df = df [['通路', '通路名稱', '客戶','客戶歸戶' , '客戶姓名',"基金歸戶",'基金公司', '股債別','交易-申購總額','交易-買回金額(匯出+轉申購)','交易-淨流入(申購總額-買回匯出-買回轉申)','月底AUM','月平均AUM','結存-月平均AUM(起迄月份合計後平均)']]  
    
    return df

# ----------------- ILP powerpoint excel 整理 function --------------------------------

def offshore_ILP(df,offshore_output_df):
    #將基金歸戶後屬於ILP的挑出來
    ILP_df = df[ df['基金歸戶']=='ILP' ]
    #用客戶歸戶做groupby
    ILP_group = ILP_df.groupby('客戶歸戶')
    ILP_output_df = pd.DataFrame()
    site_list = []
    db_aum_list = []
    

    #ILP_df.to_excel(r"D:\My Documents\kenc\桌面\test.xlsx")
    # check for BD # of accounts 
    db_number_account_list = []
    for name,group_df in ILP_group : 
        #投信=客戶
        site = name
        #用0補足NA
        group_df = group_df.fillna(value=0)
        #因為是offshore所以乘28
        db_aum = group_df['月底AUM'].sum() *28

        site_list.append(site)
        db_aum_list.append(db_aum)

        account_list = []
        accounts_name = group_df['客戶姓名'].to_list()
        for accounts in accounts_name :
            #處理客戶姓名內的問題
            if  "月撥回" in str(accounts) :
                index = accounts.index("月撥回")
                accounts = accounts[:index-1]
                account_list.append(accounts)

            elif  "月撥現" in str(accounts) :
                index = accounts.index("月撥現")
                accounts = accounts[:index-1]
                account_list.append(accounts)
                

            elif  "雙月撥回" in str(accounts) :
                index = accounts.index("雙月撥回")
                accounts = accounts[:index-1]
                account_list.append(accounts)
                

            elif  "新台幣" in str(accounts) :
                index = accounts.index("新台幣")
                accounts = accounts[:index-1]
                account_list.append(accounts)
                

            elif  "新台幣" in str(accounts) :
                index = accounts.index("新台幣")
                accounts = accounts[:index-1]
                account_list.append(accounts)
            

            elif  "新台幣" in str(accounts) :
                index = accounts.index("新台幣")
                accounts = accounts[:index-1]
                account_list.append(accounts)
                
            
            elif  "新台幣" in str(accounts) :
                index = accounts.index("新台幣")
                accounts = accounts[:index-1]
                account_list.append(accounts)
                
            else:
                account_list.append(accounts)
        #算有幾個不重複的基金
        db_number_contract = len(list(set(account_list)))
        #將每個基金數加入db_number_account_list的list中
        db_number_account_list.append(db_number_contract)

    ILP_output_df['公司名稱']            = site_list
    ILP_output_df['DB AUM']             = db_aum_list
    ILP_output_df['DB_Contract(數量)']  = db_number_account_list
    powerpoint_2                        = pd.merge(offshore_output_df,ILP_output_df,how='outer').fillna(value=0)
    powerpoint_2['DB %']                = np.round(powerpoint_2['DB AUM']/powerpoint_2['投資型保單有效契約-金額(外幣)'] * 100,2)

    return powerpoint_2

def offhsore_ILP_fund(df):
    account = pd.DataFrame()
    #將基金歸戶後屬於ILP的挑出來
    df=df[df['基金歸戶']=='ILP']
    #將客戶歸戶及基金公司做groupby後將月底AUM加總
    account_fund_df = df.groupby(['客戶歸戶','基金公司'])['月底AUM'].agg('sum').reset_index(name='AUM')
    #account_fund_df.to_csv(r'account_fund_df.csv',encoding='utf-8-sig',index=False,header=True)
    #再用客戶歸戶groupby一次
    account_fund_group = account_fund_df.groupby(['客戶歸戶'])

    fund_company = ['IAM','NN','NAMU']
    customers_list = []
    for name,df in account_fund_group:
        #客戶加入customers_list
        customers_list.append(name)
        company_list = []
        AUM_list = []
        df_company_list = df['基金公司'].to_list()
        df_company_aum_list = df['AUM'].to_list()

        for 基金公司 in fund_company : 
            if 基金公司 in df_company_list :
                #如果IAM,NN,NAMU中的其中一個在df_company_list內，並將aum也放進去
                index = df_company_list.index(基金公司)
                aum = df_company_aum_list[index]
                company_list.append(基金公司)
                AUM_list.append(aum*28)
            else:
                #如果IAM,NN,NAMU中的其中一個在df_company_list內，並將0放進去
                company_list.append(基金公司)
                AUM_list.append(0)
        # print("---------------------------------------------------")
        # print(name)
        # print(AUM_list)
        #------------------------------------------------
        row_data = pd.Series(AUM_list,index=company_list)
        account = account.append(row_data,ignore_index=True)

    account['客戶歸戶'] = customers_list
    return account

def offshore_ILP_account_df(df):
    account_df = df.groupby('客戶姓名')['月底AUM'].agg('sum').reset_index(name='DB AUM').sort_values(by='DB AUM',ascending=False).reset_index(drop=True) 
            
    account = pd.DataFrame()
    #將客戶姓名及基金公司做groupby後將月底AUM加總
    account_fund_df = df.groupby(['客戶姓名','基金公司'])['月底AUM'].agg('sum').reset_index(name='AUM')
    account_fund_group = account_fund_df.groupby(['客戶姓名'])

    fund_company = ['IAM','NN','NAMU']
    customers_list = []
    for name,df in account_fund_group:
        customers_list.append(name)
        company_list = []
        AUM_list = []
        df_company_list = df['基金公司'].to_list()
        df_company_aum_list = df['AUM'].to_list()

        # 分基金公司 , 金額
        for 基金公司 in fund_company : 
            if 基金公司 in df_company_list :
                index = df_company_list.index(基金公司)
                aum   = df_company_aum_list[index]
                company_list.append(基金公司)
                AUM_list.append(aum)
            else:
                company_list.append(基金公司)
                AUM_list.append(0)
        #------------------------------------------------

        row_data = pd.Series(AUM_list,index=company_list)
        account = account.append(row_data,ignore_index=True)

    account['客戶姓名'] = customers_list
    account_output = pd.merge(account,account_df)
    #排去
    account_output = account_output[['客戶姓名','DB AUM','NAMU','NN','IAM']]
    return account_output

def onshore_ILP(df,web_table):
    #將基金歸戶後屬於ILP的挑出來
    ILP_df = df[ df['基金歸戶']=='ILP' ]
    #用客戶歸戶做groupby
    ILP_group = ILP_df.groupby('客戶歸戶')
    ILP_output_df = pd.DataFrame()
    site_list = []
    db_aum_list = []

    # check for BD # of accounts 
    db_number_account_list = []
    for name,group_df in ILP_group : 
        account_list = []
        accounts = group_df['客戶姓名'].to_list()
        for account in accounts :
            # print("---------------------")
            # print(account)
            if  "月撥回" in str(account) :
                index = account.index("月撥回")
                account = account[:index-1]
                account_list.append(account)

            elif  "月撥現" in str(account) :
                index = account.index("月撥現")
                account = account[:index-1]
                account_list.append(account)
                
            elif  "雙月撥回" in str(account) :
                index = account.index("雙月撥回")
                account = account[:index-1]
                account_list.append(account)
            #幣別
            elif  "新台幣" in str(account) :
                index = account.index("新台幣")
                account = account[:index-1]
                account_list.append(account)
            
            elif  "累積" in str(account) :
                index = account.index("累積")
                account = account[:index-1]
                account_list.append(account)
            
            elif  "（一）" in str(account) :
                index = account.index("（一）")
                account = account[:index]
                account_list.append(account)
            
            elif  "（ＩＩ）" in str(account) :
                index = account.index("（ＩＩ）")
                account = account[:index]
                account_list.append(account)
        
            # others
            else:
                account_list.append(account)
                pass
        #算有幾個不重複的基金
        db_number_contract = len(list(set(account_list)))
        

        site = name
        db_aum = group_df['月底AUM'].sum()
        site_list.append(site)
        db_aum_list.append(db_aum)
        db_number_account_list.append(db_number_contract)
            
    ILP_output_df['公司名稱']              = site_list
    ILP_output_df['DB AUM']               = db_aum_list
    ILP_output_df['DB_Contract(數量)']    = db_number_account_list
    powerpoint_1                         = pd.merge(web_table,ILP_output_df,how='outer').fillna(value=0)
    powerpoint_1['DB %']                 = np.round(powerpoint_1['DB AUM']/powerpoint_1['投資型保單有效契約-金額(新台幣)'] * 100,2)
    return powerpoint_1

# --------------- 客戶歸戶完 做　Net-IN FLow 整理 ---------------------------------------------------------

def Client_Netflow(df):

    client_groupby_inflow  = df.groupby('客戶歸戶')['交易-申購總額'].agg('sum').reset_index(name='Inflow')
    client_groupby_outflow = df.groupby('客戶歸戶')['交易-買回金額(匯出+轉申購)'].agg('sum').reset_index(name='outflow')
    client_groupby_netflow = df.groupby('客戶歸戶')['交易-淨流入(申購總額-買回匯出-買回轉申)'].agg('sum').reset_index(name='netflow')

    client_groupby_outflow = client_groupby_outflow.drop(['客戶歸戶'],axis=1)
    client_groupby_netflow = client_groupby_netflow.drop(['客戶歸戶'],axis=1)

    output_df = pd.concat([client_groupby_inflow,client_groupby_outflow,client_groupby_netflow],axis=1)
    return output_df




if __name__ == '__main__':


    Input_File_Case_Path = r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\input"
    Selenium_Path        = r"D:\My Documents\andyhs\桌面\chromedriver.exe"

    curr_path = os.getcwd()
    data_path = os.path.join( curr_path , Input_File_Case_Path )
    months    = os.listdir(data_path)
    date      = str('2021 年 12 月') 
    # Selenium ChromeDriver Path
    path      =  Selenium_Path

    # ------------------------------ 公會 ---------------------------------
    公會_href = get_web_data(date=date,path=path)
    web_table , offshore_output_df = result_table(公會_href)

    # ------------------------------  精彩網 ------------------------------ 
    精彩網_href  = 精彩網_get_web_data(date,path)
    精彩網_table = 精彩網_result_table(精彩網_href,date)

    # ---------------------------------------------------------------------

    for i in  range(len(months)) :
        #2022/03/04 Chris 將Client分頁改成只有股債別為股的資料之統計結果
        if   "onshore"  in months[i] and "MTD" in months[i]:
            
            excel_path=data_path+"\\"+str(months[i])
            #找到YTD資料
            onshore_YTD = excel_path.replace('MTD','YTD')

            # 基金歸戶大表
            df = pd.read_excel(excel_path,skiprows=3)
            df_YTD = pd.read_excel(onshore_YTD,skiprows=3)
            df = df[df['通路']==6]
            df_YTD = df_YTD[df_YTD['通路']==6]
            df = df.reset_index(drop=True)
            #我要將MTD YTD Merge
            #new_list就是將old_list 709row的那些加上_replace即可
            new_list = []
            for x in df_YTD.columns:
                if x in ['基金','基金簡稱','客戶','客戶姓名','通路','通路名稱','股債別','基金公司','基金公司名稱']:
                    df_YTD = df_YTD.rename(columns={ x: x+'_replace'})
                    new_list.append(x+'_replace')
            df = df.merge(df_YTD,left_on=['基金','基金簡稱','客戶','客戶姓名','通路','通路名稱','股債別','基金公司','基金公司名稱'],right_on = new_list ,how='left')



            print(len(df))
            df = classify(month=months[i],df=df)
            df = classify_fund(month=months[i],df=df)
            df_股 = df[df['股債別'] == '股']
            # Onshore ILP Mandate Wallet Share by SITE 
            powerpoint_1 = onshore_ILP(df,web_table).fillna(value=0)
            powerpoint_1   = powerpoint_1[['公司名稱','契約數量','全體有效契約金額','投資型保單有效契約-數量(新台幣)','DB_Contract(數量)','投資型保單有效契約-金額(新台幣)','DB AUM','DB %']]
            # Onshore ILP mandate Wallet share – On AIA List and off AA List (by contract) --> 可拆
            account_df = df.groupby('客戶姓名')['月底AUM'].agg('sum').reset_index(name='DB AUM').sort_values(by='DB AUM',ascending=False).reset_index(drop=True).reset_index()
            # client netflow & outflow
            Clinent_df = Client_Netflow(df_股)

            print("---------------------complete " +str(months[0])+ "---------------------")
            print(df)
            print(account_df)
            print(powerpoint_1)
            print(Clinent_df)

            # ---------------------------------------------------------- save file 
            excel = Excel_Work()
            wb=excel.write_excel(df=df,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
            wb.save(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=powerpoint_1,excel_sheet_name = "By Site")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=account_df,excel_sheet_name   = "By Account")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=Clinent_df,excel_sheet_name   = str(date)+"Client")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=精彩網_table,excel_sheet_name  = str(date)+"發行全委商品(精彩網)")
        
        elif "offshore" in months[i] and "MTD" in months[i]:
            

            excel_path=data_path+"\\"+str(months[i])
            #找到YTD資料                         
            offshore_YTD = excel_path.replace('MTD','YTD')

            # 基金歸戶大表
            df = pd.read_excel(excel_path,skiprows=3)
            df_YTD = pd.read_excel(offshore_YTD,skiprows=3)
            df = df[df['通路']==6]
            df_YTD = df_YTD[df_YTD['通路']==6]
            df = df.reset_index(drop=True)
            #我要將MTD YTD Merge
            #new_list就是將old_list 709row的那些加上_replace即可
            new_list = []
            for x in df_YTD.columns:
                if x in ['基金','基金簡稱','客戶','客戶姓名','通路','通路名稱','股債別','基金公司','基金公司名稱']:
                    df_YTD = df_YTD.rename(columns={ x: x+'_replace'})
                    new_list.append(x+'_replace')
            df = df.merge(df_YTD,left_on=['基金','基金簡稱','客戶','客戶姓名','通路','通路名稱','股債別','基金公司','基金公司名稱'],right_on = new_list ,how='left')



            print(len(df))
            df = classify(months[i],df)
            df = classify_fund(months[i],df)

            # Offshore ILP mandate Wallet share – On AIA List / off AIA List --> 可拆
            account = offshore_ILP_account_df(df).reset_index()

            # Offshore ILP Mandate Wallet Share by SITE 
            powerpoint_2   = offshore_ILP(df,offshore_output_df)
            powerpoint_2_1 = offhsore_ILP_fund(df)
            powerpoint_2   = pd.merge(powerpoint_2,powerpoint_2_1,left_on="公司名稱",right_on="客戶歸戶",how='outer')
            powerpoint_2   = powerpoint_2.drop(['客戶歸戶'],axis=1).head(38).fillna(value=0)
            powerpoint_2   = powerpoint_2[['公司名稱','契約數量','全體有效契約金額','投資型保單有效契約-數量(外幣)','DB_Contract(數量)','投資型保單有效契約-金額(外幣)','DB AUM','DB %']]
            
            # client netflow & outflow
            Clinent_df       = Client_Netflow(df) 
            Clinent_df.index = Clinent_df["客戶歸戶"]
            Clinent_df = Clinent_df.drop(['客戶歸戶'],axis=1)
            Clinent_df = Clinent_df * 28  #換匯率
            Clinent_df = Clinent_df.reset_index()


            print("---------------------complete " +str(months[i])+ "---------------------")
            print(df)
            print(account) 
            print(powerpoint_2)

            # ---------------------------------------------------------- save file 
            excel = Excel_Work()
            wb=excel.write_excel(df=df,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
            wb.save(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=powerpoint_2,excel_sheet_name  = "By Site(Currency=28)")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=account,excel_sheet_name       = "By Account")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=Clinent_df,excel_sheet_name    = str(date)+"Client")
            excel.append_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\output/"+(str(months[i]).replace('MTD ',''))+".xlsx",df=精彩網_table,excel_sheet_name   = str(date)+"發行全委商品(精彩網)")






