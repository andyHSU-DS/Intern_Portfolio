#%%

#目前還沒處理的問題:onshore開戶是0

# baisc
import os 
import re
import warnings

import numpy    as np 
import pandas   as pd
import datetime as dt
from   tqdm     import tqdm
warnings.filterwarnings("ignore")

# excel 
from Module import Excel_Work

#%%


# ----- mask & exchange rate ----------

start = '2022-06-01'
end   = '2022-06-02'
exchange_rate = 28


""" 

(Note) 樓上的 Time Range  , 記得改
---------------------------------------------------------------------------------------------
報表 locate 的位置 : L:\Cross_Dept_Shared\AML&CTF\DB\Ken_Chiang\業務例行報表
---------------------------------------------------------------------------------------------
MTD_Python.py 吃 MTD onshore/offshore excel 
YTD_Python.py 吃 YTD onshore/offshore excel 
---------------------------------------------------------------------------------------------

(1.) Loading 5 Files : Onshore , Offshore , 申贖人數 , 聯繫紀錄 , Focus and Promotion Fund (每一季可能不一樣 , 跟 Joan 拿 Excel File)

(2.) 計算個別 Sales 的 Indicator (Current AUM , AVG AUM ,....)  貨幣型的計算方法有點複雜 , 不懂計算邏輯問 Joan (Reverune 要計算貨幣 , 其他不用)

(3.) Top 5 Inflow/Outflow funds 

(4.) Focus and Promotion funds' Perforamnce 

(5.) Open Account and Current Total Account --> ( 申贖人數.xlsx )

(6.) Sales & Customer 聯繫紀錄               --> ( 聯繫紀錄.xlsx )

(7.) Output to Excel Sheet & Use VBA to Clean Data

---------------------------------------------------------------------------------------------

"""

# input_file_case_path = r''
# output_file_case_path = r''

# ------- Mapping Sales Excel Sheet ---------

業務名單_df         = pd.read_excel(r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\業務姓名.xlsx')
業務名單_English_df = 業務名單_df[['Name','Section']]


# ----- import os read file case --> locate file -------------

curr_path = os.getcwd()
data_path = os.path.join(curr_path , r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\test\input')
months    = os.listdir(data_path)

#　onshore / offshore / 開戶大表（申贖人數）／ 聯繫紀錄

for i in tqdm(range(len(months))) :

    excel_path=data_path+"\\"+str(months[i])
    
    if "MTD_Onshore" in excel_path:
        onshore_df  = pd.read_excel(excel_path,skiprows=3)
        onshore_df  = onshore_df[onshore_df['通路'] != 9]
        #如果客戶是遠雄且通路為四就要更改sales為台北一組
        target_index = onshore_df[(onshore_df['客戶姓名'] == '遠雄人壽保險事業股份有限公司') & (onshore_df['通路'] == 4)].index
        #print(target_index)
        if len(target_index) != 0:
            for i in target_index:
                onshore_df.loc[i,'Sales姓名'] = '台北一組'#使用iloc會是錯的

    elif "MTD_Offshore" in excel_path:
        offshore_df = pd.read_excel(excel_path,skiprows=3)
        offshore_df  = offshore_df[offshore_df['通路'] != 9]
        #如果客戶是遠雄且通路為四就要更改sales為台北一組
        target_index = offshore_df[(offshore_df['客戶姓名'] == '遠雄人壽保險事業股份有限公司') & (offshore_df['通路'] == 4)].index
        #print(target_index)
        if len(target_index) != 0:
            for i in target_index:
                offshore_df.loc[i,'Sales姓名'] = '台北一組'#使用iloc會是錯的

    elif "申贖人數" in excel_path : 
        Open_Account = pd.read_excel(excel_path)

    elif "聯繫紀錄" in excel_path : 
        Contact_Information = pd.read_excel(excel_path)

    elif "Focus and Promotion" in excel_path:
        Focus_Promotion_Fund    = pd.read_excel(excel_path)

    else :
        pass 


# %%


# Onshore
def Agent_onshore_groupby(onshore_df,業務名單_df):

    業務_list        = 業務名單_df['姓名'].to_list()
    output_df        = pd.DataFrame()

    agent_group      = onshore_df.groupby('Sales姓名')
    Sales            = []
    Aum              = []
    Sales_Department = []
    Avg_AUM          = []
    一般申購          = []
    匯出_轉申購       = []
    淨流入            = []
    手續費            = []
    新錢              = []
    管理費            = []
    current_year = np.sort(list(set(onshore_df['年月'].to_list())))[-1]

    for name,df in agent_group : 
        
        revenue_df   = df # revenue 要包含 貨幣型
        df           = df[df['股/債/貨幣/平衡']!='貨幣'] # 算一般數字 (Gross Sales , AUM ,...etc) 都不能包含貨幣型的商品
        
        df = df.reset_index(drop=True)
        Current_df = df[df['年月']==current_year]

        Sales_Name       = name
        Sales_部門       = df['部門名稱'].to_list()[0]
        Current_AUM      = Current_df['結存-月底AUM(只有迄月資料)'].sum()
        
        Current_AVG_AUM      = Current_df['結存-月平均AUM(起迄月份合計後平均)'].sum()
        Current_Gross_Sales  = df['交易-申購總額'].sum() 
        Current_OutFlow      = df['交易-買回金額(匯出+轉申購)'].sum()
        Current_NetFlow      = df['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
        Current_New_Money    = df['交易-新錢'].sum()

        Transaction_Fee  = revenue_df['交易-手續費收入'].sum() 
        Management_Fee   = revenue_df['結存-預估管理費'].sum() 

        Sales.append(Sales_Name)
        Sales_Department.append(Sales_部門)
        #----------------------------------------
        Aum.append(Current_AUM)
        Avg_AUM.append(Current_AVG_AUM)
        一般申購.append(Current_Gross_Sales)
        #---------------------------------------
        匯出_轉申購.append(Current_OutFlow)
        淨流入.append(Current_NetFlow)
        新錢.append(Current_New_Money)

        手續費.append(Transaction_Fee)
        管理費.append(Management_Fee)
    
    # -------------- 處理 DCB Total -------------------------　

    onshore_revenue_df = onshore_df
    onshore_df         = onshore_df[onshore_df['股/債/貨幣/平衡']!='貨幣']

    def Section_Total(name,channel=None):


        if channel == 7 : #　通路　9 包含在 DCB (後來說不要)

            DCB_Onshore_DF  = onshore_df[ (onshore_df['通路']==channel)  ]
            DCB_Revenue_DF  = onshore_revenue_df[ (onshore_revenue_df['通路']==channel) ]
            
            DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year ]
            DCB_Current_DF  =  DCB_Current_DF[ ( DCB_Current_DF['通路']==channel)  ]
   
        elif channel :

            DCB_Onshore_DF  = onshore_df[ onshore_df['通路']==channel]
            DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year]
            DCB_Current_DF  =  DCB_Current_DF[ DCB_Current_DF['通路']==channel ]
            DCB_Revenue_DF  = onshore_revenue_df[ onshore_revenue_df['通路']==channel ]

        else :

            DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year]
            DCB_Revenue_DF  = onshore_revenue_df
            DCB_Onshore_DF  = onshore_df

        # ----------  月底 AUM 抓最新的月份

        DCB_Total_Name           = name
        DCB_部門_Name            =  name
        DCB_Onshore_DF_AUM       =  DCB_Current_DF['結存-月底AUM(只有迄月資料)'].sum() 
        DCB_Current_AVG_AUM      =  DCB_Current_DF['結存-月平均AUM(起迄月份合計後平均)'].sum() 
        DCB_Current_Gross_Sales  =  DCB_Onshore_DF['交易-申購總額'].sum() 
        DCB_Current_OutFlow      =  DCB_Onshore_DF['交易-買回金額(匯出+轉申購)'].sum()
        DCB_Current_NetFlow      =  DCB_Onshore_DF['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
        DCB_Current_New_Money    = DCB_Onshore_DF['交易-新錢'].sum()

        DCB_Transaction_Fee  =  DCB_Revenue_DF['交易-手續費收入'].sum() 
        DCB_Management_Fee   =  DCB_Revenue_DF['結存-預估管理費'].sum() 

        Sales.append(DCB_Total_Name)
        Sales_Department.append(DCB_部門_Name)
        #----------------------------------------
        Aum.append(DCB_Onshore_DF_AUM)
        Avg_AUM.append(DCB_Current_AVG_AUM)
        一般申購.append(DCB_Current_Gross_Sales)
        新錢.append(DCB_Current_New_Money)
        #---------------------------------------
        匯出_轉申購.append(DCB_Current_OutFlow)
        淨流入.append(DCB_Current_NetFlow )

        手續費.append(DCB_Transaction_Fee)
        管理費.append(DCB_Management_Fee)
    
    Section_Total(name="DCB Total" ,channel=7)
    Section_Total(name="C&I Total" ,channel=4)
    Section_Total(name="Ins Total" ,channel=6)
    Section_Total(name='DB Total')
    
    #---------------------------------------------------------------------------------
    # Append to  output Dataframe
    output_df['Sales Names']       = Sales
    output_df['Sales Department']  = Sales_Department
    #------------------------------------------------
    output_df['月底 AUM']     = Aum
    output_df['AVG AUM']      = Avg_AUM
    #----------------------------------------
    output_df['一般申購']      = 一般申購
    output_df['匯出/轉申購']   = 匯出_轉申購
    output_df['淨流入']        = 淨流入
    output_df['新錢']          = 新錢
    #----------------------------------------
    output_df['手續費']        = 手續費
    output_df['管理費']        = 管理費

    output_df.index = output_df['Sales Names']
    output_df = output_df.drop(['Sales Names'],axis=1)

    cols = output_df.columns 
    cols = ["onshore "+str(col)  for col in cols]
    output_df.columns  = cols
    
    return output_df

def Agent_offshore_groupby(offshore_df,業務名單_df,exchange_rate=None):


    output_df   = pd.DataFrame()
    agent_group = offshore_df.groupby('Sales姓名')
    Sales            = []
    Sales_Department = []
    Aum              = []
    Avg_AUM          = []
    一般申購          = []
    匯出_轉申購       = []
    淨流入            = []
    新錢              = []
    手續費            = []
    管理費            = []

    current_year        = np.sort(list(set(offshore_df['年月'].to_list())))[-1]
    offshore_revenue_df = offshore_df
    offshore_df         = offshore_df[offshore_df['股/債/貨幣/平衡']!='貨幣']

    for name,df in agent_group : 
        
        revenue_df = df # Sales Revenue ( Transaction Fee , Management Fee ) 要包含 貨幣型
        df         = df[df['股/債/貨幣/平衡']!='貨幣']
        Current_df = df[df['年月']==current_year]

        Sales_Name            = name
        Sales_部門            = df['部門名稱'].to_list()[0]
        Current_AUM          = Current_df['結存-月底AUM(只有迄月資料)'].sum() 
        Current_AVG_AUM      = Current_df['結存-月平均AUM(起迄月份合計後平均)'].sum()
        Current_Gross_Sales  = df['交易-申購總額'].sum() 
        Current_OutFlow      = df['交易-買回金額(匯出+轉申購)'].sum()
        Current_NetFlow      = df['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
        Current_New_Money    = df['交易-新錢'].sum()

        Transaction_Fee      = revenue_df['交易-手續費收入'].sum() 
        Management_Fee       = revenue_df['結存-預估管理費'].sum()
        

        Sales.append(Sales_Name)
        Sales_Department.append(Sales_部門)
        #---------------------
        Aum.append(Current_AUM)
        Avg_AUM.append(Current_AVG_AUM)
        # --------------------
        一般申購.append(Current_Gross_Sales)
        匯出_轉申購.append(Current_OutFlow)
        淨流入.append(Current_NetFlow)
        新錢.append(Current_New_Money)
        # --------------------
        手續費.append(Transaction_Fee)
        管理費.append(Management_Fee)

    



    def Section_Total(name,channel=None):
        
        if channel == 7 : #　通路　9 包含在 DCB

            DCB_Onshore_DF  = offshore_df[ ( offshore_df['通路']==channel)  ]
            DCB_Current_DF  = offshore_df[  offshore_df['年月']==current_year]
            DCB_Current_DF  = DCB_Current_DF[ ( DCB_Current_DF['通路']==channel)  ]
            DCB_Revenue_DF  = offshore_revenue_df[ (offshore_revenue_df['通路']==channel) ]

        elif channel :

            DCB_Onshore_DF  =  offshore_df[  offshore_df['通路']==channel]
            DCB_Current_DF  =  offshore_df[  offshore_df['年月']==current_year]
            DCB_Current_DF  =  DCB_Current_DF[DCB_Current_DF['通路']==channel]
            DCB_Revenue_DF  =  offshore_revenue_df[ (offshore_revenue_df['通路']==channel) ]


        else :

            DCB_Current_DF  =  offshore_df[ offshore_df['年月']==current_year]
            DCB_Revenue_DF  =  offshore_revenue_df
            DCB_Onshore_DF  =  offshore_df



        DCB_Total_Name           = name
        DCB_部門_Name            = name
        DCB_Onshore_DF_AUM       =  DCB_Current_DF['結存-月底AUM(只有迄月資料)'].sum() 
        DCB_Current_AVG_AUM      =  DCB_Current_DF['結存-月平均AUM(起迄月份合計後平均)'].sum() 
        DCB_Current_Gross_Sales  =  DCB_Onshore_DF['交易-申購總額'].sum() 
        DCB_Current_OutFlow      =  DCB_Onshore_DF['交易-買回金額(匯出+轉申購)'].sum()
        DCB_Current_NetFlow      =  DCB_Onshore_DF['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
        DCB_Current_New_Money    = DCB_Onshore_DF['交易-新錢'].sum()

        DCB_Transaction_Fee  =  DCB_Revenue_DF['交易-手續費收入'].sum() 
        DCB_Management_Fee   =  DCB_Revenue_DF['結存-預估管理費'].sum() 

        Sales.append(DCB_Total_Name)
        Sales_Department.append(DCB_部門_Name)
        #----------------------------------------
        Aum.append(DCB_Onshore_DF_AUM)
        Avg_AUM.append(DCB_Current_AVG_AUM)
        一般申購.append(DCB_Current_Gross_Sales)
        新錢.append(DCB_Current_New_Money)
        #---------------------------------------
        匯出_轉申購.append(DCB_Current_OutFlow)
        淨流入.append(DCB_Current_NetFlow )

        手續費.append(DCB_Transaction_Fee)
        管理費.append(DCB_Management_Fee)
        #-------------------------------------
    
    Section_Total(name="DCB Total" ,channel=7)
    Section_Total(name="C&I Total" ,channel=4)
    Section_Total(name="Ins Total" ,channel=6)
    Section_Total(name='DB Total')


    output_df['Sales Names']       = Sales
    output_df['Sales Department']  = Sales_Department
    #--------------------------------------
    output_df['月底 AUM']     = Aum
    output_df['AVG AUM']      = Avg_AUM
    #----------------------------------------
    output_df['一般申購']      = 一般申購
    output_df['匯出/轉申購']   = 匯出_轉申購
    output_df['淨流入']        = 淨流入
    output_df['新錢']          = 新錢
    #----------------------------------------
    output_df['手續費']        = 手續費
    output_df['管理費']        = 管理費

    output_df.index = output_df['Sales Names']
    output_df = output_df.drop(['Sales Names'],axis=1)

    if exchange_rate : 
        output_df[output_df.columns[1:]] = output_df[output_df.columns[1:]] * exchange_rate

    cols = output_df.columns 
    cols = ["offshore "+str(col)  for col in cols]
    output_df.columns  = cols

    return output_df

# Mapping
def Agent_onshore_Mapping(Agent_onshore_df,業務名單_df,業務名單_English_df):

    Agent_onshore_df = pd.merge(業務名單_df,Agent_onshore_df,right_index=True,left_on='姓名')
    Agent_onshore_df = Agent_onshore_df.groupby(['Name','Section'])['onshore 月底 AUM','onshore AVG AUM','onshore 一般申購','onshore 匯出/轉申購','onshore 淨流入','onshore 新錢','onshore 手續費'	,'onshore 管理費'].agg('sum').reset_index()
    Agent_onshore_df = pd.merge(業務名單_English_df,Agent_onshore_df,how='outer')
    Agent_onshore_df = Agent_onshore_df.drop_duplicates()
    
    Agent_onshore_df = Agent_onshore_df.reset_index(drop=True)
    Agent_onshore_df = Agent_onshore_df.fillna(value=0)

    Agent_onshore_df = Agent_onshore_df[['Section','Name','onshore 月底 AUM', 'onshore AVG AUM', 'onshore 一般申購',
       'onshore 匯出/轉申購', 'onshore 淨流入','onshore 新錢', 'onshore 手續費', 'onshore 管理費']]

    Agent_onshore_df['onshore Revenue'] = Agent_onshore_df['onshore 管理費'] + Agent_onshore_df['onshore 手續費']

    return Agent_onshore_df

def Agent_Offshore_Mapping(Agent_onshore_df,業務名單_df,業務名單_English_df):

        Agent_onshore_df = pd.merge(業務名單_df,Agent_onshore_df,right_index=True,left_on='姓名')
        Agent_onshore_df = Agent_onshore_df.groupby(['Name','Section'])['offshore 月底 AUM','offshore AVG AUM','offshore 一般申購','offshore 匯出/轉申購','offshore 淨流入','offshore 新錢','offshore 手續費','offshore 管理費'].agg('sum').reset_index()
        Agent_onshore_df = pd.merge(業務名單_English_df,Agent_onshore_df,how='outer')
        Agent_onshore_df = Agent_onshore_df.drop_duplicates()
        
        Agent_onshore_df = Agent_onshore_df.reset_index(drop=True)
        Agent_onshore_df = Agent_onshore_df.fillna(value=0)

        Agent_onshore_df = Agent_onshore_df[['Section','Name','offshore 月底 AUM', 'offshore AVG AUM',
       'offshore 一般申購', 'offshore 匯出/轉申購', 'offshore 淨流入', 'offshore 新錢','offshore 手續費',
       'offshore 管理費']]

        Agent_onshore_df['offshore Revenue'] = Agent_onshore_df['offshore 管理費'] + Agent_onshore_df['offshore 手續費']

        return Agent_onshore_df

# Address Others
def Address_Others_Offshore(Agent_offshore_df,業務名單_df):

    Other_index   = []
    Current_index = []
    業務_list = 業務名單_df['姓名'].to_list()
    sales     = Agent_offshore_df.index.to_list()

    for i in range(Agent_offshore_df.shape[0]):
        sale = sales[i]
        if sale not in  業務_list : 
            Other_index.append(i)
        
        elif sale  in  業務_list :
            Current_index.append(i)

    Agent_Department_Group   = Agent_offshore_df.iloc[Current_index]
    Agent_Department_Group   = Agent_Department_Group.reset_index()

    Agent_Other_offshore_df  = Agent_offshore_df.iloc[Other_index]
    Agent_Other_offshore_df  = Agent_Other_offshore_df .reset_index()
    Other_Department_Group   = Agent_Other_offshore_df.groupby('offshore Sales Department')['offshore 月底 AUM','offshore AVG AUM','offshore 一般申購','offshore 匯出/轉申購','offshore 淨流入','offshore 新錢','offshore 手續費','offshore 管理費'].agg('sum').reset_index()
    Other_Department_Group ['Sales Names'] = ""

    for i in range(Other_Department_Group.shape[0]) :

        if "法人行銷台北一組" in Other_Department_Group['offshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台北一組"
        
        elif "法人行銷台北二組" in Other_Department_Group['offshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台北二組"
        
        elif "台中分公司" in Other_Department_Group['offshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台中業務公單"
        
        elif "高雄分公司" in Other_Department_Group['offshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "高雄業務公單"
        
        # elif "機構法人部" in Other_Department_Group['offshore Sales Department'][i]:
            
        #     Other_Department_Group['Sales Names'][i] = "機構法人部"

        else:
            Other_Department_Group['Sales Names'][i] = Other_Department_Group['offshore Sales Department'][i]


    Agent_offshore_df = pd.concat([Agent_Department_Group,Other_Department_Group])
    Agent_offshore_df = Agent_offshore_df.reset_index(drop=True)
    Agent_offshore_df = Agent_offshore_df.groupby('Sales Names')['offshore 月底 AUM','offshore AVG AUM','offshore 一般申購','offshore 匯出/轉申購','offshore 淨流入','offshore 新錢','offshore 手續費','offshore 管理費'].agg('sum')

    return Agent_offshore_df

def Address_Others_Onshore(Agent_onshore_df,業務名單_df):

    Other_index = []
    Current_index = []
    業務_list = 業務名單_df['姓名'].to_list()
    sales     = Agent_onshore_df .index.to_list()

    for i in range(Agent_onshore_df .shape[0]):
        sale = sales[i]
        if sale not in  業務_list : 
            Other_index.append(i)
        
        elif sale  in  業務_list :
            Current_index.append(i)


    Agent_Department_Group   = Agent_onshore_df.iloc[Current_index]
    Agent_Department_Group   = Agent_Department_Group.reset_index()

    Agent_Other_onshore_df  = Agent_onshore_df.iloc[Other_index]
    Agent_Other_onshore_df  = Agent_Other_onshore_df .reset_index()
    Agent_Other_onshore_df

    Other_Department_Group   = Agent_Other_onshore_df.groupby('onshore Sales Department')['onshore 月底 AUM','onshore AVG AUM','onshore 一般申購','onshore 匯出/轉申購','onshore 淨流入','onshore 手續費','onshore 管理費'].agg('sum').reset_index()
    Other_Department_Group ['Sales Names'] = ""


    for i in range(Other_Department_Group.shape[0]) :

        if "法人行銷台北一組" in Other_Department_Group['onshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台北一組"
        
        elif "法人行銷台北二組" in Other_Department_Group['onshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台北二組"
        
        elif "台中分公司" in Other_Department_Group['onshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "台中業務公單"
        
        elif "高雄分公司" in Other_Department_Group['onshore Sales Department'][i]:
            
            Other_Department_Group['Sales Names'][i] = "高雄業務公單"
        
        # elif "機構法人部" in Other_Department_Group['offshore Sales Department'][i]:
            
        #     Other_Department_Group['Sales Names'][i] = "機構法人部"

        else:

            Other_Department_Group['Sales Names'][i] = Other_Department_Group['onshore Sales Department'][i]


    Agent_onshore_df = pd.concat([Agent_Department_Group,Other_Department_Group])
    Agent_onshore_df = Agent_onshore_df.reset_index(drop=True)
    Agent_onshore_df = Agent_onshore_df.groupby('Sales Names')['onshore 月底 AUM','onshore AVG AUM','onshore 一般申購','onshore 匯出/轉申購','onshore 淨流入','onshore 新錢','onshore 手續費','onshore 管理費'].agg('sum')

    return Agent_onshore_df

# Combine onshore / offshore
def DB_Total_df(Agent_onshore_df,Agent_offshore_df):

    Agent_AUM_Columns  = [ str(col)[9:] for col in Agent_offshore_df.columns  ] 
    Agent_AUM_df       = pd.DataFrame(Agent_onshore_df.values + Agent_offshore_df.values,columns=Agent_AUM_Columns )
    
    Agent_AUM_df = Agent_AUM_df.drop([''],axis=1)
    Agent_AUM_df.insert( 0 ,column='Name', value=Agent_offshore_df['Name'].values)
    Agent_AUM_df.insert( 0 ,column='Department', value=Agent_offshore_df['Section'].values)

    return Agent_AUM_df

# 1.) onshore
Agent_onshore_df  = Agent_onshore_groupby(onshore_df,業務名單_df)
Agent_onshore_df  = Address_Others_Onshore(Agent_onshore_df,業務名單_df)
Agent_onshore_df  = Agent_onshore_Mapping(Agent_onshore_df,業務名單_df,業務名單_English_df)
Agent_onshore_df

# 2.) offshore
Agent_offshore_df = Agent_offshore_groupby(offshore_df,業務名單_df,exchange_rate=exchange_rate)
Agent_offshore_df = Address_Others_Offshore(Agent_offshore_df,業務名單_df)
Agent_offshore_df = Agent_Offshore_Mapping(Agent_offshore_df,業務名單_df,業務名單_English_df)
Agent_offshore_df

# 3.) Total
Agent_AUM_df = DB_Total_df(Agent_onshore_df,Agent_offshore_df)
Agent_AUM_df


# %%


"""
----------
通路 | 部門
-----------
 4   | C&I
-----------
 7   | DCB 
-----------
 6   | Ins
-----------
 9   | 投信總公司(未加入)
-----------

Department Flow --> Onshore/Offshore (DCB,Ins,C&I)
Offshore_Department_Product_Output --> Offshore 不同基金公司的 Inflow / Outflow
"""

def Onshore_Department_Flow(onshore_df):

    Channel_List = [4,7,6]
    Channel_Name = ['C&I','DCB','Ins']
    Channel_onshore_group = onshore_df.groupby('通路')

    section_name = []
    section_newmoney = []
    section_purchase = []
    section_outflow  = []
    section_netflow  = []
    section_AUM      = []
    section_AVG_AUM  = []


    for name,df in Channel_onshore_group : 
        
        if name in Channel_List : 

            index = Channel_List.index(name)
            name  = Channel_Name[index]
            
            Total_Purchase = df['交易-申購總額'].sum()
            Total_OutFlow  = df['交易-買回金額(匯出+轉申購)'].sum()
            Total_NetFlow  = df['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
            Total_NewMoney = df['交易-新錢'].sum()
            Total_AUM      = df['結存-月底AUM(只有迄月資料)'].sum()
            Total_AVG_AUM  = df['結存-月平均AUM(起迄月份合計後平均)'].sum()

            section_name.append(name)
            section_purchase.append(Total_Purchase)
            section_newmoney.append(Total_NewMoney)
            section_outflow.append(Total_OutFlow)
            section_netflow.append(Total_NetFlow)
            section_AUM.append(Total_AUM)
            section_AVG_AUM.append(Total_AVG_AUM)


    Department_Output = pd.DataFrame()
    Department_Output['Department'] = section_name
    Department_Output['New money']  = section_newmoney
    Department_Output['Gross Sales'] = section_purchase
    Department_Output['redemption and switch out'] = section_outflow
    Department_Output['net flow']        = section_netflow
    Department_Output['Onshore AUM']     = section_AUM
    Department_Output['Onshore AVG AUM'] = section_AVG_AUM


    Department_Output = Department_Output.transpose()
    Department_Output.columns = Department_Output.loc['Department']
    Department_Output = Department_Output.drop(['Department'],axis=0)
    
    cols = Department_Output.columns
    cols = ["onshore "+col for col in cols]
    Department_Output.columns = cols


    return Department_Output

def Offshore_Department_Flow(offshore_df,exchange_rate=None):


    Channel_List = [4,7,6]
    Channel_Name = ['C&I','DCB','Ins']
    Channel_offshore_group = offshore_df.groupby('通路')
    a = 0

    for name,df in Channel_offshore_group : 
        
        df         = df[df['股/債/貨幣/平衡']!='貨幣']
        if name in Channel_List : 

            index = Channel_List.index(name)
            name  = Channel_Name[index]

            output = df.groupby('基金公司名稱')['交易-新錢','交易-申購總額','交易-買回金額(匯出+轉申購)','交易-淨流入(申購總額-買回匯出-買回轉申)','結存-月底AUM(只有迄月資料)','結存-月平均AUM(起迄月份合計後平均)'].agg('sum')
            cols   = output.columns
            cols   = [name+" "+str(col) for col in cols]
            output.columns = cols 
            

            if a == 0 :
                Final_output = output
                a+=1
            else : 
                Final_output = pd.concat([Final_output,output],axis=1)
    
    # -------------- Return final output (by 基金 Product)

    section_name = []
    section_newmoney = []
    section_purchase = []
    section_outflow  = []
    section_netflow  = []
    section_AUM      = []
    section_AVG_AUM  = []


    for name,df in Channel_offshore_group : 
        
        if name in Channel_List : 

            index = Channel_List.index(name)
            name  = Channel_Name[index]
            
            Total_Purchase = df['交易-申購總額'].sum()
            Total_OutFlow  = df['交易-買回金額(匯出+轉申購)'].sum()
            Total_NetFlow  = df['交易-淨流入(申購總額-買回匯出-買回轉申)'].sum()
            Total_NewMoney = df['交易-新錢'].sum()
            Total_AUM      = df['結存-月底AUM(只有迄月資料)'].sum()
            Total_AVG_AUM  = df['結存-月平均AUM(起迄月份合計後平均)'].sum()

            section_name.append(name)
            section_purchase.append(Total_Purchase)
            section_newmoney.append(Total_NewMoney)
            section_outflow.append(Total_OutFlow)
            section_netflow.append(Total_NetFlow)
            section_AUM.append(Total_AUM)
            section_AVG_AUM.append(Total_AVG_AUM)
    

    Department_Output = pd.DataFrame()
    Department_Output['Department'] = section_name
    Department_Output['New money']  = section_newmoney
    Department_Output['Gross Sales'] = section_purchase
    Department_Output['redemption and switch out'] = section_outflow
    Department_Output['net flow']        = section_netflow
    Department_Output['Onshore AUM']     = section_AUM
    Department_Output['Onshore AVG AUM'] = section_AVG_AUM


    Department_Output = Department_Output.transpose()
    Department_Output.columns = Department_Output.loc['Department']
    Department_Output = Department_Output.drop(['Department'],axis=0)
    
    cols = Department_Output.columns
    cols = ["offshore "+col for col in cols]
    Department_Output.columns = cols

    if exchange_rate:
        Department_Output = Department_Output*28
        Final_output =Final_output*28

    Final_output = Final_output.reset_index()
    Final_output = Final_output.transpose()
    Final_output.columns = Final_output.loc['基金公司名稱']
    Final_output = Final_output.drop(['基金公司名稱'],axis=0)
    Final_output = Final_output.reset_index()

    return Final_output , Department_Output

def Deparment_Flow(Onshore_Department_Flow_df,Offshore_Department_Flow_df):

    Total = pd.DataFrame(Onshore_Department_Flow_df.values + Offshore_Department_Flow_df.values)
    Total.index = Onshore_Department_Flow_df.index
    Total.columns = ['Total C&I','Total DCB','Total Ins']

    output = pd.concat([Onshore_Department_Flow_df,Offshore_Department_Flow_df,Total],axis=1)
    output = output.reset_index()
    
    return output

def Offshore_Department_Product(Offshore_Department_Product_Output):
    Offshore_Total = []
    for i in range(Offshore_Department_Product_Output.shape[0]):
        sum = Offshore_Department_Product_Output.iloc[i][1:].sum()
        Offshore_Total.append(sum)

    Offshore_Department_Product_Output.insert(1, "Total", Offshore_Total)
    
    return Offshore_Department_Product_Output


Onshore_Department_Flow_df                                     =  Onshore_Department_Flow(onshore_df)
Offshore_Department_Product_Output,Offshore_Department_Flow_df =  Offshore_Department_Flow(offshore_df,exchange_rate=28)
Offshore_Department_Product_Output = Offshore_Department_Product(Offshore_Department_Product_Output)

Offshore_Department_Product_Output



# %%

"""
Onshore/Offshore
------------------
Top 5 Inflow funds 
Top 5 Outflow funds 
"""

def Address_Fund_Name(Fund_Name):

    if "基金-" in Fund_Name  and "晉達" not in Fund_Name  and "荷寶" not in Fund_Name :

        index =  Fund_Name.index("基金-")
        Fund_Name = Fund_Name[:index+2]

        return Fund_Name

    elif "債-" in Fund_Name :

        index =  Fund_Name.index("債-")
        Fund_Name = Fund_Name[:index+1]

        return Fund_Name

    elif "基金 -" in Fund_Name and "晉達" not in Fund_Name  and "荷寶" not in Fund_Name  :
        
        index =  Fund_Name.index("基金 -")
        Fund_Name = Fund_Name[:index+2] 
        # print(Fund_Name)

        return Fund_Name

    elif "野村基金(愛爾蘭系列)" in Fund_Name :

        index_1     =  Fund_Name.index("野村基金(愛爾蘭系列)")
        Fund_Name_0 = Fund_Name[:index_1+11] 

        Fund_Name_1 = Fund_Name[index_1+11:] 

        index_2     =  Fund_Name_1.index("基金")
        Fund_Name_2 = Fund_Name_1[:index_2+2] 

        Fund_Name = Fund_Name_0 + Fund_Name_2

        return Fund_Name

    elif "鋒裕匯理基金" in Fund_Name:

        if "債券" in Fund_Name :

            index     =  Fund_Name.index("債券")
            Fund_Name = Fund_Name[:index+2] 

        elif "股票" in Fund_Name :

            index     =  Fund_Name.index("股票")
            Fund_Name = Fund_Name[:index+2] 

        return Fund_Name

    elif "晉達" in Fund_Name  :
        
        index     =  Fund_Name.index("基金")
        Fund_Name = Fund_Name[index+3:] 
        Fund_Name = "晉達"+ Fund_Name 
        
        if "債券" in Fund_Name :

            index     =  Fund_Name.index("債券")
            Fund_Name = Fund_Name[:index+2] +"基金"

            return Fund_Name

        elif "股票" in Fund_Name :

            index     =  Fund_Name.index("股票")
            Fund_Name = Fund_Name[:index+2] +"基金"

            return Fund_Name

        elif "基金" in Fund_Name :

            index     =  Fund_Name.index("基金")
            Fund_Name = Fund_Name[:index+2] 

            return Fund_Name

        elif "股份" in Fund_Name :

            index     =  Fund_Name.index("股份")
            Fund_Name = Fund_Name[:index+2] +"基金"

            return Fund_Name

    elif "荷寶" in Fund_Name  :
        
        index     =  Fund_Name.index("基金")
        Fund_Name = Fund_Name[index+3:] 
        Fund_Name = "荷寶-"+ Fund_Name 
        
        if "債券" in Fund_Name :

            index     =  Fund_Name.index("債券")
            Fund_Name = Fund_Name[:index+2] +"基金"

        elif "股票" in Fund_Name :

            index     =  Fund_Name.index("股票")
            Fund_Name = Fund_Name[:index+2] +"基金"

        elif "基金" in Fund_Name :

            index     =  Fund_Name.index("基金")
            Fund_Name = Fund_Name[:index+2] 

        elif "股份" in Fund_Name :

            index     =  Fund_Name.index("股份")
            Fund_Name = Fund_Name[:index+2] +"基金"
        
        return Fund_Name

    else:

        try :
            index =  Fund_Name.index("基金")
            Fund_Name = Fund_Name[:index+2]

            return Fund_Name
        except:
            return Fund_Name

onshore_df  = onshore_df.reset_index(drop=True)
offshore_df = offshore_df.reset_index(drop=True)

onshore_df['基金簡稱(調整)']  = onshore_df.apply(lambda x : Address_Fund_Name(x['基金簡稱']),axis=1)
offshore_df['基金簡稱(調整)'] = offshore_df.apply(lambda x : Address_Fund_Name(x['基金簡稱']),axis=1)

def Onshore_Flow(onshore_df):

    # Onshore Inflow
    Onshore_Inflow = onshore_df.groupby('基金簡稱(調整)')['交易-申購總額'].agg('sum').reset_index(name='Inflow')
    Onshore_Inflow.columns = ['Top 5 Inflow 基金','Inflow']
    Onshore_Inflow = Onshore_Inflow.sort_values(by='Inflow',ascending=False)
    Onshore_Inflow = Onshore_Inflow.reset_index(drop=True)
    
    # Onshore Inflow
    Onshore_Outflow = onshore_df.groupby('基金簡稱(調整)')['交易-買回金額(匯出+轉申購)'].agg('sum').reset_index(name='Outflow')
    Onshore_Outflow.columns = ['Top 5 Outflow 基金','Outflow']
    Onshore_Outflow = Onshore_Outflow.sort_values(by='Outflow',ascending=False)
    Onshore_Outflow = Onshore_Outflow.reset_index(drop=True)

    # Onshore Net Flow
    NetFlow = pd.merge(Onshore_Outflow ,Onshore_Inflow,left_on='Top 5 Outflow 基金',right_on='Top 5 Inflow 基金')
    NetFlow['NetFlow'] = NetFlow['Inflow'] - NetFlow['Outflow']

    Net_Inflow  = NetFlow.sort_values(by="NetFlow",ascending=False)
    Net_Outflow = NetFlow.sort_values(by="NetFlow",ascending=True)

    Net_Inflow         = Net_Inflow[['Top 5 Outflow 基金','NetFlow']]
    Net_Inflow.columns = ['Top 5 Net Inflow 基金','Amount']
    Net_Inflow         = Net_Inflow.reset_index(drop=True)


    Net_Outflow         = Net_Outflow[['Top 5 Outflow 基金','NetFlow']]
    Net_Outflow.columns = ['Top 5 Net Outflow 基金','Amount']
    Net_Outflow         = Net_Outflow.reset_index(drop=True)

    Net_Inflow     = Net_Inflow[:5]
    Net_Outflow    = Net_Outflow[:5]

    Onshore_Inflow  = Onshore_Inflow[:5]
    Onshore_Outflow = Onshore_Outflow[:5]
    Onshore_Fund_Flow = pd.concat([Onshore_Inflow,Onshore_Outflow,Net_Inflow,Net_Outflow],axis=1)
    Onshore_Fund_Flow = Onshore_Fund_Flow.reset_index()

    return Onshore_Fund_Flow

def Offshore_Flow(offshore_df,exchange_rate=None):

    # Offshore Inflow &　Outflow Dataframe
    Offshore_Inflow = offshore_df.groupby('基金簡稱(調整)')['交易-申購總額'].agg('sum').reset_index(name='Inflow')
    Offshore_Inflow.columns = ['Top 5 Inflow 基金','Inflow']
    Offshore_Inflow['Inflow'] *=28 
    Offshore_Inflow = Offshore_Inflow.sort_values(by='Inflow',ascending=False)
    Offshore_Inflow = Offshore_Inflow.reset_index(drop=True)

    Offshore_Outflow = offshore_df.groupby('基金簡稱(調整)')['交易-買回金額(匯出+轉申購)'].agg('sum').reset_index(name='Outflow')
    Offshore_Outflow.columns = ['Top 5 Outflow 基金','Outflow']
    Offshore_Outflow['Outflow'] *=28 
    Offshore_Outflow = Offshore_Outflow.sort_values(by='Outflow',ascending=False)
    Offshore_Outflow = Offshore_Outflow.reset_index(drop=True)
    
    
    # Onshore Net Flow
    NetFlow = pd.merge(Offshore_Outflow ,Offshore_Inflow,left_on='Top 5 Outflow 基金',right_on='Top 5 Inflow 基金')
    NetFlow['NetFlow'] = NetFlow['Inflow'] - NetFlow['Outflow']

    Net_Inflow  = NetFlow.sort_values(by="NetFlow",ascending=False)
    Net_Outflow = NetFlow.sort_values(by="NetFlow",ascending=True)

    Net_Inflow         = Net_Inflow[['Top 5 Outflow 基金','NetFlow']]
    Net_Inflow.columns = ['Top 5 Net Inflow 基金','Amount']
    Net_Inflow         = Net_Inflow.reset_index(drop=True)

    Net_Outflow            = Net_Outflow[['Top 5 Outflow 基金','NetFlow']]
    Net_Outflow.columns    = ['Top 5 Net Outflow 基金','Amount']
    Net_Outflow            = Net_Outflow.reset_index(drop=True)

    # Concate Dataframe
    Net_Inflow     = Net_Inflow[:5]
    Net_Outflow    = Net_Outflow[:5]

    Offshore_Inflow = Offshore_Inflow[:5]
    Offshore_Outflow = Offshore_Outflow[:5]
    Offshore_Fund_Flow = pd.concat([Offshore_Inflow,Offshore_Outflow,Net_Inflow,Net_Outflow],axis=1)


    Offshore_Fund_Flow = Offshore_Fund_Flow.reset_index()
    return Offshore_Fund_Flow

def Focus_Promot_Fund(onshore_df):

    # 做 Promotion & Focus 的 Money Flow
    # Inflow
    Onshore_Inflow = onshore_df.groupby('基金簡稱')['交易-申購總額'].agg('sum').reset_index(name='Inflow')
    Onshore_Inflow.columns = ['基金','Inflow']
    Onshore_Inflow = Onshore_Inflow.sort_values(by='Inflow',ascending=False)
    Onshore_Inflow = Onshore_Inflow.reset_index(drop=True)
    Onshore_Inflow = Onshore_Inflow[:5]

    # Outflow 
    Onshore_Outflow = onshore_df.groupby('基金簡稱')['交易-買回金額(匯出+轉申購)'].agg('sum').reset_index(name='Outflow')
    Onshore_Outflow.columns = ['基金','Outflow']
    Onshore_Outflow = Onshore_Outflow.sort_values(by='Outflow',ascending=False)
    Onshore_Outflow = Onshore_Outflow.reset_index(drop=True)
    Onshore_Outflow = Onshore_Outflow[:5]
    
    # New Money
    Onshore_NewMoney = onshore_df.groupby('基金簡稱')['交易-新錢'].agg('sum').reset_index(name='New Money')
    Onshore_NewMoney.columns = ['基金','New Money']
    Onshore_NewMoney = Onshore_NewMoney.sort_values(by='New Money',ascending=False)
    Onshore_NewMoney = Onshore_NewMoney.reset_index(drop=True)
    Onshore_NewMoney = Onshore_NewMoney[:5]


    Onshore_Fund_Flow = pd.merge(Onshore_Inflow,Onshore_Outflow)
    Onshore_Fund_Flow = pd.merge(Onshore_Fund_Flow,Onshore_NewMoney)
    Onshore_Fund_Flow = Onshore_Fund_Flow.reset_index()

    return Onshore_Fund_Flow


Onshore_Fund_Flow  = Onshore_Flow(onshore_df[(onshore_df['通路'] == 4) & (onshore_df['股/債/貨幣/平衡'] != '貨幣')  ] )
Offshore_Fund_Flow = Offshore_Flow( offshore_df[ (offshore_df['通路'] == 4) & (offshore_df['股/債/貨幣/平衡'] != '貨幣') ] )



# 一張報表就好了
sapce_df = pd.DataFrame(np.zeros((2,Offshore_Fund_Flow.shape[1])),columns=Offshore_Fund_Flow.columns)
sapce_df = sapce_df.replace(0,"")
list_1        = pd.DataFrame([['onshore','-','-','-','-','-','-','-','-']],columns=Offshore_Fund_Flow.columns)
list_3        = pd.DataFrame([['offshore','-','-','-','-','-','-','-','-']],columns=Offshore_Fund_Flow.columns)
offshore_list = pd.DataFrame([['index','Top 5 Inflow 基金','Inflow','Top 5 Outflow 基金','Outflow','Top 5 Net Inflow 基金','Net Inflow','Top 5 Net Outflow 基金','Net Outflow']],columns=Offshore_Fund_Flow.columns)
Fund_Flow     = pd.concat([list_1,Onshore_Fund_Flow,sapce_df,list_3,offshore_list,Offshore_Fund_Flow])
Fund_Flow 


# ------------------- Focus / Promotion Fund ----------------------------------------

Onshore_Focus_Promotion_Fund  = Focus_Promotion_Fund[Focus_Promotion_Fund['基金公司'] == '投信']
Offshore_Focus_Promotion_Fund = Focus_Promotion_Fund[Focus_Promotion_Fund['基金公司'] != '投信']

Onshore_Focus_Promotion_Groupby  = Onshore_Focus_Promotion_Fund.groupby('Focus /Promotion Fund')
Offshore_Focus_Promotion_Groupby = Offshore_Focus_Promotion_Fund.groupby('Focus /Promotion Fund')
for name,df in Onshore_Focus_Promotion_Groupby:
    print(name)
def Focus_Promotion(Focus_Promotion_Groupby,Fund_df,exchange_rate,Offshore=None):

    Classification_list = []
    Fund_list           = []
    Fund_Name_list      = []

    for name,df in Focus_Promotion_Groupby : 

        Classification = name
        fund_name      = df['基金'].to_list()
        funds          = df['基金代號'].to_list()

        Classification_list.append(Classification)
        Fund_Name_list.append(fund_name)
        Fund_list.append(funds)
    #print('-'*50)
    #print('Classification_list:',len(Classification_list))
    #print('-'*50)
    #print('Fund_Name_list:',len(Fund_Name_list))
    #print('-'*50)
    #print('Fund_list:',len(Fund_list))


    # Classification_list
    Foucs_Promotion_Df = pd.DataFrame()
    #print('Fund df:',Fund_df)
    for i in range(len(Classification_list)):
        #print(Fund_list[i])
        #這邊很重要，我們Fund_list裡面的數字都是int，但在SA抓下來的資料都是str且前面會有0(如果小於10)
        Fund_list_str = [str(x) if len(str(x))>1 else str(0)+str(x) for x in Fund_list[i]]
        #print(Fund_list_str)
        row_index       = [ index for index,value in enumerate(Fund_df['基金']) if str(value) in Fund_list_str ]
        #print(row_index)
        onshore_groupby = Fund_df.iloc[row_index]

        Focus_Promotion_Fund_Flow           = Focus_Promot_Fund(onshore_groupby)
        Focus_Promotion_Fund_Flow['index']  = Classification_list[i]
        #print('Focus_Promotion_Fund_Flow:',Focus_Promotion_Fund_Flow)
        Foucs_Promotion_Df = Foucs_Promotion_Df.append(Focus_Promotion_Fund_Flow)
        #因為如果你是一年初開始可能會碰到有的沒有紀錄，這樣會報錯
        if len(Foucs_Promotion_Df)>0:
            Foucs_Promotion_Df.reset_index(drop=True,inplace=True)
            Foucs_Promotion_Df['Net Flow'] = Foucs_Promotion_Df['Inflow'] - Foucs_Promotion_Df['Outflow']
            #print(Foucs_Promotion_Df)
            #print('-'*50)

            if Offshore:
                Foucs_Promotion_Df['Net Flow']   *= exchange_rate 
                Foucs_Promotion_Df['Inflow']     *= exchange_rate
                Foucs_Promotion_Df['Outflow']    *= exchange_rate
                Foucs_Promotion_Df['New Money']  *= exchange_rate
        else:
            Foucs_Promotion_Df = pd.DataFrame(['-','-',0,0,0,0])
            Foucs_Promotion_Df.columns = ['index','基金','Inflow','Outflow','New Money','Net Flow']
            #print('No Data')

    return Foucs_Promotion_Df

Onshore_Focus_Fund_Flow  = Focus_Promotion(Onshore_Focus_Promotion_Groupby  , onshore_df  , exchange_rate                )
Onshore_Focus_Fund_Flow['index'] = Onshore_Focus_Fund_Flow['index'].apply(lambda x:re.sub(r'\ATW S&M Cap-Promotional','Nomura TW S&M Cap-Promotional',x))
Offshore_Focus_Fund_Flow = Focus_Promotion(Offshore_Focus_Promotion_Groupby , offshore_df , exchange_rate ,Offshore=True )


Focus_Fund_Flow   = pd.concat([Onshore_Focus_Fund_Flow,Offshore_Focus_Fund_Flow])
Promotion_Index   = [index for index,value in enumerate(Focus_Fund_Flow['index'].to_list() ) if "Promotional" in value]
Focus_Index       = [index for index,value in enumerate(Focus_Fund_Flow['index'].to_list() ) if "Focus"       in value]


Promotion_df      = Focus_Fund_Flow.iloc[Promotion_Index].reset_index(drop=True)
Focus_df          = Focus_Fund_Flow.iloc[Focus_Index].reset_index(drop=True)

Promotion_df['Type'] = "Promotion"
Focus_df['Type']     = "Focus"


Focus_Fund_Flow         = pd.concat([Promotion_df,Focus_df])
Focus_Fund_Flow_Groupby = Focus_Fund_Flow.groupby(['index','Type'])[['Inflow','Outflow','Net Flow','New Money']].agg('sum').sort_values(by='Type').reset_index()

list_5        = pd.DataFrame( [['Total','-','-','-','-','-','-']] , columns=Focus_Fund_Flow.columns )
sapce_df      = pd.DataFrame(np.zeros((2,Focus_Fund_Flow.shape[1])),columns=Focus_Fund_Flow.columns)
sapce_df      = sapce_df.replace(0,"")


Focus_Fund_Flow = pd.concat([Focus_Fund_Flow,sapce_df,list_5,Focus_Fund_Flow_Groupby]).fillna(value="-")
Focus_Fund_Flow 

# %%


# Sales New Open Account &　AUM Calculation
def Onshre_Offshore_df(Open_Account,start,end,exchange_rate,mask=None):
    #　SA Data 開戶日　裡面會有 0 --> 真的傻眼 (try,except or delet the data)
    Open_Account = Open_Account[Open_Account['開戶日']!=0]
    def Address_Datetime(date):
        date = pd.to_datetime(date,format='%Y%m%d')
        return date
    
    def Current_Months(df,start,end):
        # locate Time Range
        mask = ( (df['開戶日'] >= start ) & (df['開戶日'] <= end ) ) 
        df   = df.loc[mask]

        return df

    def Address_offshore_AUM(df,exchange_rate):
        # offshore AUM to NTD
        cols = df.columns
        cols = [col for col in cols if "AUM" in col]
        df[cols] = exchange_rate * df[cols]

        return df 


    Open_Account['開戶日']   = Open_Account.apply(lambda x : Address_Datetime(x['開戶日']),axis=1)
    Open_Onshore_Account    = Open_Account[ Open_Account['資料別']=='Onshore' ]
    Open_Offshore_Account   = Open_Account[(Open_Account['資料別']=='Offshore') | (Open_Account['資料別']=='Omnibus') ]
    
    if mask:
        Current_Open_Onshore_Account  = Current_Months(Open_Onshore_Account,start,end)
        Current_Open_Offshore_Account = Current_Months(Open_Offshore_Account,start,end)
    else : 
        Current_Open_Onshore_Account  = Open_Onshore_Account
        Current_Open_Offshore_Account = Open_Offshore_Account

    Current_Open_Offshore_Account = Address_offshore_AUM(Current_Open_Offshore_Account,exchange_rate)
    
    return Current_Open_Onshore_Account,Current_Open_Offshore_Account

def Calculate_Open_Account_Numbers(Current_Open_Onshore_Account,Current_Open_Offshore_Account):

    def Customer_Recognition(name):
        if len(str(name)) > 4 :

            return "Corporate" 

        elif len(str(name)) <= 4 :

            return "Individual"
    
    def Calculate(Current_Open_Onshore_Account):
        print('Current_Open_Onshore_Account len:',len(Current_Open_Onshore_Account))
        if len(Current_Open_Onshore_Account)>0:
            Sales_Open_Account = Current_Open_Onshore_Account.groupby(['姓名','Client Type'])['戶號'].count().reset_index(name='Numbers of Open Accounts')
            Sales_Open_Account['Individual'] = 0 
            Sales_Open_Account['Corporate']  = 0 

            # ----------------------------------------------------
            for i in range(Sales_Open_Account.shape[0]):

                Client_Types = Sales_Open_Account['Client Type'][i]
                Numbers      = Sales_Open_Account['Numbers of Open Accounts'][i]

                if Client_Types == 'Individual' : 

                    Sales_Open_Account['Individual'][i] = Numbers

                if Client_Types == 'Corporate' : 

                    Sales_Open_Account['Corporate'][i] = Numbers


            Sales_Open_Account = Sales_Open_Account.drop(['Client Type'],axis=1)
            Sales_Open_Account = Sales_Open_Account.groupby('姓名').sum().reset_index()
        
        else:
            print('Calculate:Onshore無人開戶')
            Sales_Open_Account = pd.DataFrame({'姓名':業務名單_df['姓名'].to_list(),'Numbers of Open Accounts':[0]*len(業務名單_df['姓名'].to_list())\
                                            ,'Individual':[0]*len(業務名單_df['姓名'].to_list()),'Corporate':[0]*len(業務名單_df['姓名'].to_list())})
        
        return Sales_Open_Account

    if len(Current_Open_Onshore_Account)>0:
        Current_Open_Onshore_Account['Client Type']  = Current_Open_Onshore_Account.apply(lambda x : Customer_Recognition(x['戶名']),axis=1)
        Current_Open_Onshore_Account['姓名']   = Current_Open_Onshore_Account['業務名稱']
    else:
        pass
    Current_Open_Onshore_Account          = Calculate(Current_Open_Onshore_Account)
    print('Current_Open_Onshore_Account:',Current_Open_Onshore_Account)
    

    # Offshore 比較會有 --> 完全沒有人開戶的問題
    if Current_Open_Offshore_Account.shape[0] > 0 :
        Current_Open_Offshore_Account['Client Type'] = Current_Open_Offshore_Account.apply(lambda x : Customer_Recognition(x['戶名']),axis=1)
        Current_Open_Offshore_Account['姓名']        = Current_Open_Offshore_Account['業務名稱']
        Current_Open_Offshore_Account = Calculate(Current_Open_Offshore_Account)
    
    if Current_Open_Offshore_Account.shape[0] == 0 :
        cols = Current_Open_Onshore_Account.columns
        Current_Open_Offshore_Account = pd.DataFrame(columns=cols)
        
    # Cols Naming 
    cols   = []
    cols_1 = [ str(col) for i,col in enumerate(Current_Open_Onshore_Account.columns)  if i < 1]
    cols_2 = ["Offshore " + str(col) for i,col in enumerate(Current_Open_Onshore_Account.columns)  if i >= 1]
    cols.extend(cols_1)
    cols.extend(cols_2)
    Current_Open_Offshore_Account.columns = cols

    # Cols Naming 
    cols   = []
    cols_1 = [ str(col) for i,col in enumerate(Current_Open_Onshore_Account.columns)  if i < 1]
    cols_2 = ["Onshore " + str(col) for i,col in enumerate(Current_Open_Onshore_Account.columns)  if i >= 1]
    cols.extend(cols_1)
    cols.extend(cols_2)
    Current_Open_Onshore_Account.columns = cols
    #print('Current_Open_Onshore_Account:',Current_Open_Onshore_Account.columns)
    #print('Current_Offen_Offshore_Account:',Current_Open_Offshore_Account.columns)
    return Current_Open_Onshore_Account , Current_Open_Offshore_Account 

def Final_OnShore_Offshore(Current_Open_Onshore_Account , Current_Open_Offshore_Account) : 


    onshore = pd.merge(業務名單_df,Current_Open_Onshore_Account,how='outer')
    print('onshore:',onshore)
    onshore = onshore.groupby(['Name','Section']).sum().reset_index()
    print('onshore_groupby:',onshore)
    onshore = pd.merge(業務名單_df[['Name','Section']],onshore,how='outer')
    print('onshore_merge:',onshore)

    offshore = pd.merge(業務名單_df,Current_Open_Offshore_Account,how='outer')
    print('offshore:',offshore)
    offshore = offshore.groupby(['Name','Section']).sum().reset_index()
    print('offshore_groupby:',offshore)
    offshore = pd.merge(業務名單_df[['Name','Section']],offshore,how='outer')
    print('offshore_merge:',offshore)

    onshore  = onshore.drop_duplicates()
    offshore = offshore.drop_duplicates()

    Open_Account = pd.merge( onshore , offshore ,on=['Name','Section'])
    print('Open Account:',Open_Account.columns)
    Open_Account = Open_Account[['Name','Section','Onshore Individual','Onshore Corporate','Offshore Individual','Offshore Corporate','Onshore Numbers of Open Accounts','Offshore Numbers of Open Accounts']]

    return onshore , offshore , Open_Account

# ---------------------------------------------------------------

# New Account AUM by Sales
def Sales_New_Account_AUM(Current_Open_Onshore_Account,Current_Open_Offshore_Account):

    def Caculate_AUM(Current_Open_Onshore_Account):

        cols = Current_Open_Onshore_Account.columns
        Current_AUM_columns = [col for col in cols if "AUM" in col][0]

        if len(Current_Open_Onshore_Account)<1:
            print('Salse_New_Account_AUM:Onshore無人開戶')
            Onshore_Account_AUM = pd.DataFrame({'姓名':業務名單_df['姓名'].to_list(),'AUM':[0]*len(業務名單_df['姓名'].to_list())\
                                            ,'Individual AUM':[0]*len(業務名單_df['姓名'].to_list()),'Corporate AUM':[0]*len(業務名單_df['姓名'].to_list())})
        else:
            Onshore_Account_AUM = Current_Open_Onshore_Account.groupby(['姓名','Client Type'])[Current_AUM_columns].agg('sum').reset_index(name='AUM')
            Onshore_Account_AUM['Individual AUM'] = 0
            Onshore_Account_AUM['Corporate AUM']  = 0


            for i in range(Onshore_Account_AUM.shape[0]):

                client_type = Onshore_Account_AUM['Client Type'][i]
                AUM_Numbers = Onshore_Account_AUM['AUM'][i]

                if client_type == 'Individual' : 

                    Onshore_Account_AUM['Individual AUM'][i] = AUM_Numbers
                
                if client_type == 'Corporate' : 

                    Onshore_Account_AUM['Corporate AUM'][i] = AUM_Numbers


            Onshore_Account_AUM =  Onshore_Account_AUM.drop(['Client Type'],axis=1)
            Onshore_Account_AUM =  Onshore_Account_AUM.groupby('姓名').sum().reset_index()
        
        return Onshore_Account_AUM

    Onshore_Account_AUM  = Caculate_AUM(Current_Open_Onshore_Account)

    if Current_Open_Offshore_Account.shape[0] > 0:
        Offshore_Account_AUM = Caculate_AUM(Current_Open_Offshore_Account)
    
    if Current_Open_Offshore_Account.shape[0] == 0 :
        cols = Onshore_Account_AUM.columns
        Offshore_Account_AUM = pd.DataFrame(columns=cols)

    # ---------- Cols --------------------
    # offshore
    cols  = []
    col_1 = [ str(col) for i,col in enumerate(Onshore_Account_AUM.columns) if i<1 ] 
    col_2 = [ "Offshore " + str(col) for i,col in enumerate(Onshore_Account_AUM.columns) if i>=1 ] 
    cols.extend(col_1)
    cols.extend(col_2)

    Offshore_Account_AUM.columns = cols 
    # onshore
    cols  = []
    col_1 = [ str(col) for i,col in enumerate(Onshore_Account_AUM.columns) if i<1 ] 
    col_2 = [ "Onshore " + str(col) for i,col in enumerate(Onshore_Account_AUM.columns) if i>=1 ] 
    cols.extend(col_1)
    cols.extend(col_2)
    Onshore_Account_AUM.columns = cols 

    return Onshore_Account_AUM , Offshore_Account_AUM

def Final_Sales_AUM(Onshore_Account_AUM,Offshore_Account_AUM):

    onshore = pd.merge(業務名單_df,Onshore_Account_AUM ,how='outer')
    onshore = onshore.groupby(['Name','Section']).sum().reset_index()
    onshore = pd.merge(業務名單_df[['Name','Section']],onshore,how='outer')

    offshore = pd.merge(業務名單_df,Offshore_Account_AUM,how='outer')
    offshore = offshore.groupby(['Name','Section']).sum().reset_index()
    offshore = pd.merge(業務名單_df[['Name','Section']],offshore,how='outer')

    onshore  = onshore.drop_duplicates()
    offshore = offshore.drop_duplicates()

    Open_Account_AUM = pd.merge( onshore , offshore )
    Open_Account_AUM = Open_Account_AUM[['Name','Section','Onshore Individual AUM','Onshore Corporate AUM','Offshore Individual AUM','Offshore Corporate AUM','Onshore AUM','Offshore AUM']]
    Open_Account_AUM  = Open_Account_AUM[(Open_Account_AUM['Section']!='Institution') & (Open_Account_AUM['Section']!='Institution SubTotal') & (Open_Account_AUM['Section']!='DCB Total')  & (Open_Account_AUM['Section']!='Ins Total') & (Open_Account_AUM['Section']!='DB Total')]
    
    return onshore ,offshore , Open_Account_AUM


# Numbers of Account
Current_Open_Onshore_Account,Current_Open_Offshore_Account  = Onshre_Offshore_df(Open_Account,start,end,exchange_rate,mask=True)
Onshore_Account , Offshore_Account                          = Calculate_Open_Account_Numbers(Current_Open_Onshore_Account,Current_Open_Offshore_Account)
Account_onshore , Account_offshore  , Open_Account_Number   = Final_OnShore_Offshore(Onshore_Account , Offshore_Account)
# AUM
Onshore_Account_AUM , Offshore_Account_AUM = Sales_New_Account_AUM(Current_Open_Onshore_Account,Current_Open_Offshore_Account)
Account_Onshore_AUM ,Account_Offshore_AUM , Open_Account_AUM = Final_Sales_AUM(Onshore_Account_AUM,Offshore_Account_AUM)

New_open_account = pd.merge(Open_Account_Number,Open_Account_AUM)

New_open_account['Total Individual'] = New_open_account['Onshore Individual'] + New_open_account['Offshore Individual'] 
New_open_account['Total Corporate']  = New_open_account['Onshore Corporate']  + New_open_account['Offshore Corporate'] 
New_open_account


def Sales_Customer(Total_Onshore_Account , Total_Offshore_Account):


    def Address_Datetime(date):
        #　SA Data 開戶日　裡面會有 0 --> 真的傻眼 (try,except or delet the data)
        date = str(date)
        date = pd.to_datetime(date,format='%Y%m%d')
        return date

    Total_Account = pd.concat([Total_Onshore_Account, Total_Offshore_Account])
    #2022/02/10進行更改，現在是Total AUM(不含MMF)，所以將貨幣的去除
    aum_cols_total     = [col for col in Total_Account.columns if "AUM" in col][0]
    #print(aum_cols_total)
    aum_cols_貨幣      = [col for col in Total_Account.columns if '貨幣' in col][0]
    #print(aum_cols_貨幣)
    Total_Account['AUM(NMMF)'] = Total_Account[aum_cols_total] - Total_Account[aum_cols_貨幣]
    Total_Account_AUM  = Total_Account.groupby(['戶號','業務名稱'])['AUM(NMMF)'].agg('sum').reset_index()
    Total_Account_Date = Total_Account.groupby(['戶號','業務名稱'])[['最後申購日','最後買回日']].agg('max').reset_index()
    Total_Account      = pd.merge(Total_Account_AUM,Total_Account_Date)
 

    Total_Account['最後申購日'] = Total_Account.apply(lambda x : Address_Datetime(x['最後申購日']),axis=1)
    Total_Account['最後買回日'] = Total_Account.apply(lambda x : Address_Datetime(x['最後買回日']),axis=1)
    year = str(dt.datetime.now())[:4]
    Sales_Client  = Total_Account.groupby('業務名稱')

    sales = []
    sales_1億   = []
    sales_3千萬 = []
    sales_2千萬 = []
    sales_1千萬 = []

    sales_5百萬 = []
    sales_3百萬 = []
    sales_1百萬以上 = []
    sales_1百萬以下 = []

    sales_no_aum_account    = []
    sales_total_account     = []
    sales_with_aum_account  = []
    sales_active_percentage = []

    sales_current_active = []
    sales_current_active_percentage = []


    for name,df in tqdm(Sales_Client) :

        sales_name = name
        sales_1億以上_account   = df[df['AUM(NMMF)']  >= 100000000].shape[0]
        sales_3千萬以上_account = df[(df['AUM(NMMF)'] >= 30000000)  & (df['AUM(NMMF)'] < 100000000)].shape[0]
        sales_2千萬以上_account = df[(df['AUM(NMMF)'] >= 20000000)  & (df['AUM(NMMF)'] < 30000000) ].shape[0]
        sales_1千萬以上_account = df[(df['AUM(NMMF)'] >= 10000000)  & (df['AUM(NMMF)'] < 20000000) ].shape[0]

        sales_5百萬以上_account = df[(df['AUM(NMMF)']  >= 5000000)  & (df['AUM(NMMF)'] < 10000000) ].shape[0]
        sales_3百萬以上_account = df[(df['AUM(NMMF)']  >= 3000000)  & (df['AUM(NMMF)'] < 5000000)  ].shape[0]
        sales_1百萬以上_account = df[(df['AUM(NMMF)']  >= 1000000)  & (df['AUM(NMMF)'] < 3000000)  ].shape[0]
        sales_1百萬以下_account = df[(df['AUM(NMMF)']  <  1000000)  & (df['AUM(NMMF)'] > 0)        ].shape[0]

        sales_total               = df.shape[0]
        sales_No_AUM_account      = df[df['AUM(NMMF)'] == 0 ].shape[0]
        Sales_Aum_account         = df[df['AUM(NMMF)'] > 0  ].shape[0]
        Avtive_Account_Percentage = np.round( Sales_Aum_account/sales_total ,decimals=2)

        Current_Active_account    = df[ ( df['最後申購日'] > year) | ( df['最後買回日'] > year)  ].shape[0]
        Current_Active_Percentage = np.round( Current_Active_account / sales_total ,decimals=2 )
        # --------------------------------
        sales .append(sales_name)
        sales_1億 .append(sales_1億以上_account)
        sales_3千萬 .append(sales_3千萬以上_account)
        sales_2千萬 .append(sales_2千萬以上_account)
        sales_1千萬 .append(sales_1千萬以上_account)

        sales_5百萬 .append(sales_5百萬以上_account)
        sales_3百萬 .append(sales_3百萬以上_account)
        sales_1百萬以上 .append(sales_1百萬以上_account)
        sales_1百萬以下 .append(sales_1百萬以下_account)

        sales_no_aum_account    .append(sales_No_AUM_account)
        sales_total_account     .append(sales_total)
        sales_with_aum_account  .append(Sales_Aum_account)
        sales_active_percentage .append(Avtive_Account_Percentage)

        sales_current_active .append(Current_Active_account)
        sales_current_active_percentage .append(Current_Active_Percentage)

    Sales_Customer_by_AUM = pd.DataFrame()
    Sales_Customer_by_AUM['姓名']       = sales
    Sales_Customer_by_AUM['≧100M']      = sales_1億
    Sales_Customer_by_AUM['≧30M,<100M'] = sales_3千萬
    Sales_Customer_by_AUM['≧20M,<30M'] = sales_2千萬
    Sales_Customer_by_AUM['≧10M,<20M'] = sales_1千萬
    Sales_Customer_by_AUM['≧5M,<10M']  = sales_5百萬
    Sales_Customer_by_AUM['≧3M,<5M']   = sales_3百萬
    Sales_Customer_by_AUM['≧1M,<3M']   = sales_1百萬以上
    Sales_Customer_by_AUM['<1M,>0']    = sales_1百萬以下

    Sales_Customer_by_AUM['AUM=0']      = sales_no_aum_account
    Sales_Customer_by_AUM['Total']      = sales_total_account 
    Sales_Customer_by_AUM['有庫存客戶']  = sales_with_aum_account
    Sales_Customer_by_AUM["(有庫存庫戶/Total)%"] = sales_active_percentage 

    Sales_Customer_by_AUM['Active( Transaction within 1y)'] = sales_current_active
    ###2022/02/09更改，(Active/Total)計算方式更改
    Sales_Customer_by_AUM['(有庫存客戶/Total)%']                = round(Sales_Customer_by_AUM['Active( Transaction within 1y)']/Sales_Customer_by_AUM['有庫存客戶'],2)

    return Sales_Customer_by_AUM

def Sales_by_AUM(Sales_Customer_by_AUM):

    Sales_AUM  = pd.merge(業務名單_df ,Sales_Customer_by_AUM)
    Sales_AUM  = Sales_AUM.groupby(['Name','Section']).sum().reset_index()
    Sales_AUM  = pd.merge(業務名單_df[['Name','Section']],Sales_AUM,how='outer')
    Sales_AUM  = Sales_AUM.drop_duplicates().reset_index(drop=True)
    Sales_AUM  = Sales_AUM.fillna(value=0)

    Sales_AUM  = Sales_AUM[(Sales_AUM['Section']!='Institution') & (Sales_AUM['Section']!='Institution SubTotal') & (Sales_AUM['Section']!='DCB Total')  & (Sales_AUM['Section']!='Ins Total') & (Sales_AUM['Section']!='DB Total')]

    return Sales_AUM

def Total_AUM(Sales_AUM,Total_Onshore_Account,Total_Offshore_Account):

    Total_Account      = pd.concat([Total_Onshore_Account, Total_Offshore_Account])
    #2022/02/10進行更改，現在是Total AUM(不含MMF)，所以將貨幣的去除
    aum_cols_total     = [col for col in Total_Account.columns if "AUM" in col][0]
    #print(aum_cols_total)
    aum_cols_貨幣      = [col for col in Total_Account.columns if '貨幣' in col][0]
    #print(aum_cols_貨幣)
    Total_Account['AUM(NMMF)'] = Total_Account[aum_cols_total] - Total_Account[aum_cols_貨幣]
    Total_Account_AUM  = Total_Account.groupby(['戶號','業務名稱'])['AUM(NMMF)'].agg('sum').reset_index()
    Total_Account_Date = Total_Account.groupby(['戶號','業務名稱'])[['最後申購日','最後買回日']].agg('max').reset_index()
    Total_Account      = pd.merge(Total_Account_AUM,Total_Account_Date)
    Total_Account

    df = Total_Account

    sales_1億以上_values    = df[df['AUM(NMMF)']  > 100000000]['AUM(NMMF)'].sum()
    sales_3千萬以上_values  = df[(df['AUM(NMMF)']  >= 30000000) & (df['AUM(NMMF)'] <= 100000000)]['AUM(NMMF)'].sum()
    sales_2千萬以上_values  = df[(df['AUM(NMMF)'] >= 20000000)  & (df['AUM(NMMF)']  < 30000000) ]['AUM(NMMF)'].sum()
    sales_1千萬以上_values  = df[(df['AUM(NMMF)'] >= 10000000)  & (df['AUM(NMMF)']  < 20000000) ]['AUM(NMMF)'].sum()

    sales_5百萬以上_values = df[ (df['AUM(NMMF)']  >= 5000000)  & (df['AUM(NMMF)'] < 10000000) ]['AUM(NMMF)'].sum()
    sales_3百萬以上_values = df[ (df['AUM(NMMF)']  >= 3000000)  & (df['AUM(NMMF)'] < 5000000)  ]['AUM(NMMF)'].sum()
    sales_1百萬以上_values = df[ (df['AUM(NMMF)']  >= 1000000)  & (df['AUM(NMMF)'] < 3000000)  ]['AUM(NMMF)'].sum()
    sales_1百萬以下_values = df[ (df['AUM(NMMF)']  <  1000000)  & (df['AUM(NMMF)'] > 0)        ]['AUM(NMMF)'].sum()



    columns            = Sales_AUM.columns 
    aum_section_values = ["","Total AUM(NMMF)",sales_1億以上_values,sales_3千萬以上_values,sales_2千萬以上_values,sales_1千萬以上_values,sales_5百萬以上_values,sales_3百萬以上_values,sales_1百萬以上_values,sales_1百萬以下_values,0,df['AUM(NMMF)'].sum(),"","","",""]
    aum_section_values = pd.Series(aum_section_values,index=columns)
    Sales_AUM = Sales_AUM.append(aum_section_values,ignore_index=True)


    return Sales_AUM

# Current Total Account 
Total_Onshore_Account , Total_Offshore_Account  = Onshre_Offshore_df(Open_Account,start,end,exchange_rate)
Sales_Customer_by_AUM = Sales_Customer(Total_Onshore_Account , Total_Offshore_Account)
Sales_AUM             = Sales_by_AUM(Sales_Customer_by_AUM)
Sales_AUM             = Total_AUM(Sales_AUM,Total_Onshore_Account,Total_Offshore_Account)
Sales_AUM



#%%

def Month_Year_df(Contact_Information):


    def MTD_Time(date):

        date = str(date)
        
        if "上午" in date : 

            index = date.index("上午")
            date  = date[:index] 

        elif "下午" in date : 
            index = date.index("下午")
            date  = date[:index] 

        date = pd.to_datetime(date)

        return date 

    def Calculate_Contact(Contact_Information):

        activity_df         = Contact_Information[(Contact_Information['活動代號'] == 'Existing VVIP') & (Contact_Information['企劃活動'] == 'C&I Joint Call') ]
        Contact_Information = Contact_Information[ Contact_Information['活動代號'] != 'Existing VVIP' ]
        Contact_Information = pd.concat([Contact_Information,activity_df])


        contact_df = Contact_Information.groupby('建檔業務姓名')
        output_df = pd.DataFrame()

        姓名 = []
        contact_numbers = []
        contact_clients = []
        join_call_numbers = []

        for name,df in contact_df :

            Sales_Names   = name
            聯絡次數       = df.shape[0]
            聯絡人數       = df.groupby('客戶ID')['聯絡狀態'].count().shape[0]
            Join_Call_次數 = df[df['企劃活動']=='C&I Joint Call'].shape[0]


            姓名.append(Sales_Names)
            contact_numbers.append(聯絡次數)
            contact_clients.append(聯絡人數)
            join_call_numbers.append(Join_Call_次數)



        output_df['姓名']    = 姓名
        output_df['聯絡次數'] = contact_numbers
        output_df['聯絡人數'] = contact_clients
        output_df['Join Call 次數'] = join_call_numbers
        output_df['聯絡次數/聯絡人數'] = output_df['聯絡次數'] / output_df['聯絡人數'] 

        return output_df

    Contact_Information = Contact_Information[Contact_Information['聯絡狀態'] =='連絡成功']
    Contact_Information['建檔日期'] = Contact_Information.apply(lambda x : MTD_Time(x['建檔日期']),axis=1 )
    mask = ( (Contact_Information['建檔日期'] >= start ) & (Contact_Information['建檔日期'] <= end ) ) 
    Month_Contact_Information   = Contact_Information.loc[mask]

    YTD_output_df = Calculate_Contact(Contact_Information)
    MTD_output_df = Calculate_Contact(Month_Contact_Information)

    return YTD_output_df , MTD_output_df

# YTD / MTD Change Columns 
def Given_Colume_Time_Range(YTD_output_df,MTD_output_df):

    YTD_Columns = []
    YTD_Columns_1 = [ str(col) for i,col in enumerate(YTD_output_df.columns) if i < 1  ]
    YTD_Columns_2 = ["YTD " +str(col) for i,col in enumerate(YTD_output_df.columns)  if i >= 1 ]
    YTD_Columns.extend(YTD_Columns_1)
    YTD_Columns.extend(YTD_Columns_2)

    YTD_output_df.columns = YTD_Columns

    MTD_Columns = []
    MTD_Columns_1 = [ str(col) for i,col in enumerate(MTD_output_df.columns) if i < 1  ]
    MTD_Columns_2 = ["MTD " +str(col) for i,col in enumerate(MTD_output_df.columns)  if i >= 1 ]
    MTD_Columns.extend(MTD_Columns_1)
    MTD_Columns.extend(MTD_Columns_2)

    MTD_output_df.columns = MTD_Columns

    return YTD_output_df, MTD_output_df

def Merge_df (YTD_output_df, MTD_output_df):

    MTD_output_df = pd.merge(業務名單_df,MTD_output_df,how='outer')
    MTD_output_df = MTD_output_df.groupby(['Section','Name']).sum().reset_index()
    MTD_output_df = pd.merge(業務名單_df[['Section','Name']],MTD_output_df,how='outer')
    MTD_output_df = MTD_output_df.drop_duplicates()

    
    YTD_output_df = pd.merge(業務名單_df,YTD_output_df,how='outer')
    YTD_output_df = YTD_output_df.groupby(['Section','Name']).sum().reset_index()
    YTD_output_df = pd.merge(業務名單_df[['Section','Name']],YTD_output_df,how='outer')
    YTD_output_df = YTD_output_df.drop_duplicates()

    return YTD_output_df, MTD_output_df

def Final_人數占比_Total_Account(Final_df,Sales_AUM):
    
    Sales_Total_Account = Sales_AUM[['Name','Section','Total']]
    Final_df = pd.merge(Final_df,Sales_Total_Account)
    Final_df['YTD 聯絡人數占比'] = Final_df['YTD 聯絡人數'] / Final_df['Total']
    Final_df['MTD 聯絡人數占比'] = Final_df['MTD 聯絡人數'] / Final_df['Total']
    Final_df = Final_df.fillna(value = 0)
    Final_df = Final_df[['Section', 'Name','Total','YTD 聯絡次數', 'YTD 聯絡人數', 'YTD Join Call 次數',
        'YTD 聯絡次數/聯絡人數','YTD 聯絡人數占比','MTD 聯絡次數', 'MTD 聯絡人數', 'MTD Join Call 次數',
        'MTD 聯絡次數/聯絡人數','MTD 聯絡人數占比']]

    Final_df = Final_df[(Final_df['Section']!='Institution') & (Final_df['Section']!='Institution SubTotal') & (Final_df['Section']!='Ins Total') & (Final_df['Section']!='DCB Total') & (Final_df['Section']!='DB Total')]
    
    return Final_df


YTD_output_df , MTD_output_df = Month_Year_df(Contact_Information)
YTD_output_df, MTD_output_df  = Given_Colume_Time_Range(YTD_output_df,MTD_output_df)
YTD_output_df, MTD_output_df  = Merge_df (YTD_output_df , MTD_output_df)
Final_df                      = pd.merge(YTD_output_df  , MTD_output_df)
Final_df                      = Final_人數占比_Total_Account(Final_df,Sales_AUM)

# 修正第二次 (排除 others 的聯絡次數)
Exclude_Other_Section_Index = [index for index,value in enumerate( Final_df['Name'].to_list() ) if "Other" not in value]
Final_df = Final_df.iloc[Exclude_Other_Section_Index]




#%%

output_path = r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\test\output\MTD_Monthly_Sales_Performance_Statistic_2022_06-02_May.xlsx'

excel = Excel_Work()
wb=excel.write_excel(df=Agent_AUM_df,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
wb.save(output_path)

excel.append_excel(excel_path=output_path , df=Agent_onshore_df                   , excel_sheet_name="Agent Onshore"                   )
excel.append_excel(excel_path=output_path , df=Agent_offshore_df                  , excel_sheet_name="Agent Offshore"                  )
excel.append_excel(excel_path=output_path , df=Offshore_Department_Product_Output , excel_sheet_name="Offshore By Company_Department"  )
excel.append_excel(excel_path=output_path , df=Fund_Flow                          , excel_sheet_name="Top 5 Fund Flow"                 )
excel.append_excel(excel_path=output_path , df=Focus_Fund_Flow                    , excel_sheet_name="Promotion and Focus Fund"        )

excel.append_excel(excel_path=output_path , df=New_open_account                   , excel_sheet_name="New open account"                )
excel.append_excel(excel_path=output_path , df=Sales_AUM                          , excel_sheet_name="customer by AUM"                 )
excel.append_excel(excel_path=output_path , df=Final_df                           , excel_sheet_name="聯繫紀錄"                         )

print('Task Complete !')


# %%
