#%%

# baisc
import os 
import re
import warnings

import numpy    as np 
import pandas   as pd
import datetime as dt
from tqdm       import tqdm
warnings.filterwarnings("ignore")

# excel 
from Module import Excel_Work




# ----- mask & exchange rate ----------

exchange_rate = 28

""" 

(Note) 樓上的 Exchange_Rate , 記得改
---------------------------------------------------------------------------------------------
報表 locate 的位置 : L:\Cross_Dept_Shared\AML&CTF\DB\Ken_Chiang\業務例行報表
---------------------------------------------------------------------------------------------
MTD_Python.py 吃 MTD onshore/offshore excel 
YTD_Python.py 吃 YTD onshore/offshore excel 
---------------------------------------------------------------------------------------------

(1.) Loading 5 Files : Onshore , Offshore , 申贖人數 , 聯繫紀錄 , Focus and Promotion Fund 

(2.) 計算個別 Sales 的 Indicator (Current AUM , AVG AUM ,....)  貨幣型的計算方法有點複雜 , 看不懂可以問 Joan

(3.) Top 5 Inflow/Outflow funds 

(4.) Focus and Promotion funds' Perforamnce 

(5.) Open Account and Current Total Account --> (申贖人數.xlsx)

(6.) Sales & Customer 聯繫紀錄 -->  (聯繫紀錄.xlsx)
---------------------------------------------------------------------------------------------

"""

# ------- Mapping Sales Excel Sheet ---------

業務名單_df         = pd.read_excel(r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\業務姓名.xlsx')
業務名單_English_df = 業務名單_df[['Name','Section']]

# ----- import os read file case --> locate file -------------

curr_path = os.getcwd()
data_path = os.path.join(curr_path , r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\input')
months    = os.listdir(data_path)

#　onshore / offshore / 開戶大表（申贖人數）／ 聯繫紀錄
for i in tqdm(range(len(months))) :

    excel_path=data_path+"\\"+str(months[i])
    
    if "YTD_Onshore" in excel_path :
        onshore_df  = pd.read_excel(excel_path,skiprows=3)
        onshore_df  = onshore_df[onshore_df['通路'] != 9]

    elif "YTD_Offshore" in excel_path:
        offshore_df = pd.read_excel(excel_path,skiprows=3)
        offshore_df = offshore_df[offshore_df['通路'] != 9]

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

    業務_list = 業務名單_df['姓名'].to_list()

    output_df = pd.DataFrame()
    agent_group = onshore_df.groupby('Sales姓名')
    Sales      = []
    Aum        = []
    Sales_Department = []
    Avg_AUM    = []
    一般申購    = []
    匯出_轉申購 = []
    淨流入     = []
    手續費     = []
    新錢       = []
    管理費     = []
    # current_year = np.sort(list(set(onshore_df['年月'].to_list())))[-1]

    for name,df in agent_group : 
        
        revenue_df   = df # revenue 要包含 貨幣型
        df           = df[df['股/債/貨幣/平衡']!='貨幣']
        
        df = df.reset_index(drop=True)
        # Current_df = df[df['年月']==current_year]

        Sales_Name       = name
        Sales_部門       = df['部門名稱'].to_list()[0]
        Current_AUM      = df['結存-月底AUM(只有迄月資料)'].sum()
        
        Current_AVG_AUM      = df['結存-月平均AUM(起迄月份合計後平均)'].sum()
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


        if channel == 7 : #　通路　9 包含在 DCB


            DCB_Onshore_DF  = onshore_df[ (onshore_df['通路']==channel)  ]
            DCB_Revenue_DF  = onshore_revenue_df[ (onshore_revenue_df['通路']==channel) ]
            
            # DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year ]
            # DCB_Current_DF  =  DCB_Current_DF[ ( DCB_Current_DF['通路']==channel)  ]
   
        elif channel :
            DCB_Onshore_DF  = onshore_df[ onshore_df['通路']==channel]
            # DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year]
            # DCB_Current_DF  =  DCB_Current_DF[ DCB_Current_DF['通路']==channel ]
            DCB_Revenue_DF  = onshore_revenue_df[ onshore_revenue_df['通路']==channel ]

        
        else :
            # DCB_Current_DF  = onshore_df[ onshore_df['年月']==current_year]
            DCB_Revenue_DF  = onshore_revenue_df
            DCB_Onshore_DF  = onshore_df

        # ----------  月底 AUM 抓最新的月份

        DCB_Total_Name     = name
        DCB_部門_Name      = name
        DCB_Onshore_DF_AUM =  DCB_Onshore_DF['結存-月底AUM(只有迄月資料)'].sum() 
        DCB_Current_AVG_AUM      =  DCB_Onshore_DF['結存-月平均AUM(起迄月份合計後平均)'].sum() 
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


    output_df = pd.DataFrame()
    agent_group = offshore_df.groupby('Sales姓名')
    Sales   = []
    Sales_Department = []
    Aum     = []
    Avg_AUM = []
    一般申購 = []
    匯出_轉申購 = []
    淨流入   = []
    新錢     = []
    手續費   = []
    管理費   = []
    # current_year = np.sort(list(set(offshore_df['年月'].to_list())))[-1]

    offshore_revenue_df = offshore_df
    offshore_df         = offshore_df[offshore_df['股/債/貨幣/平衡']!='貨幣']

    for name,df in agent_group : 
        
        revenue_df = df # revenue 要包含 貨幣型
        df         = df[df['股/債/貨幣/平衡']!='貨幣']
        # Current_df = df[df['年月']==current_year]

        Sales_Name     = name
        Sales_部門     = df['部門名稱'].to_list()[0]
        Current_AUM          = df['結存-月底AUM(只有迄月資料)'].sum() 
        Current_AVG_AUM      = df['結存-月平均AUM(起迄月份合計後平均)'].sum()
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
            # DCB_Current_DF  = offshore_df[  offshore_df['年月']==current_year]
            # DCB_Current_DF  = DCB_Current_DF[ ( DCB_Current_DF['通路']==channel)  ]
            DCB_Revenue_DF  = offshore_revenue_df[ (offshore_revenue_df['通路']==channel) ]

        elif channel :

            DCB_Onshore_DF  =  offshore_df[  offshore_df['通路']==channel]
            # DCB_Current_DF  =  offshore_df[  offshore_df['年月']==current_year]
            # DCB_Current_DF  =  DCB_Current_DF[DCB_Current_DF['通路']==channel]
            DCB_Revenue_DF  =  offshore_revenue_df[ (offshore_revenue_df['通路']==channel) ]


        else :
            # DCB_Current_DF  =  offshore_df[ offshore_df['年月']==current_year]
            DCB_Revenue_DF  =  offshore_revenue_df
            DCB_Onshore_DF  =  offshore_df



        DCB_Total_Name           = name
        DCB_部門_Name            = name
        DCB_Onshore_DF_AUM       =  DCB_Onshore_DF['結存-月底AUM(只有迄月資料)'].sum() 
        DCB_Current_AVG_AUM      =  DCB_Onshore_DF['結存-月平均AUM(起迄月份合計後平均)'].sum() 
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
    
    Section_Total(name="DCB Total",channel=7)
    Section_Total(name="C&I Total",channel=4)
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

    Other_index = []
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

    Agent_AUM_df = pd.DataFrame(Agent_onshore_df.values + Agent_offshore_df.values,columns=Agent_AUM_Columns )
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


    # Classification_list
    Foucs_Promotion_Df = pd.DataFrame()

    for i in range(len(Classification_list)):

        row_index       = [ index for index,value in enumerate(Fund_df['基金']) if value in Fund_list[i] ]
        onshore_groupby = Fund_df.iloc[row_index]

        Focus_Promotion_Fund_Flow           = Focus_Promot_Fund(onshore_groupby)
        Focus_Promotion_Fund_Flow['index']  = Classification_list[i]

        Foucs_Promotion_Df = Foucs_Promotion_Df.append(Focus_Promotion_Fund_Flow)

    Foucs_Promotion_Df = Foucs_Promotion_Df.reset_index(drop=True)
    #有的Foucs_Promotion_Df出來沒有資料，如果繼續往下會報錯，這邊做一個預防
    if len(Foucs_Promotion_Df)>0:
        Foucs_Promotion_Df['Net Flow'] = Foucs_Promotion_Df['Inflow'] - Foucs_Promotion_Df['Outflow']

        if Offshore:
                Foucs_Promotion_Df['Net Flow']   *= exchange_rate 
                Foucs_Promotion_Df['Inflow']     *= exchange_rate
                Foucs_Promotion_Df['Outflow']    *= exchange_rate
                Foucs_Promotion_Df['New Money']   *= exchange_rate

    return Foucs_Promotion_Df

Onshore_Focus_Fund_Flow  = Focus_Promotion(Onshore_Focus_Promotion_Groupby  , onshore_df  , exchange_rate                )
Offshore_Focus_Fund_Flow = Focus_Promotion(Offshore_Focus_Promotion_Groupby , offshore_df , exchange_rate ,Offshore=True )


Focus_Fund_Flow          = pd.concat([Onshore_Focus_Fund_Flow,Offshore_Focus_Fund_Flow])
Promotion_Index          = [index for index,value in enumerate(Focus_Fund_Flow['index'].to_list() ) if "Promotional" in value]
Focus_Index              = [index for index,value in enumerate(Focus_Fund_Flow['index'].to_list() ) if "Focus"       in value]


Promotion_df             = Focus_Fund_Flow.iloc[Promotion_Index].reset_index(drop=True)
Focus_df                 = Focus_Fund_Flow.iloc[Focus_Index].reset_index(drop=True)

Promotion_df['Type']     = "Promotion"
Focus_df['Type']         = "Focus"


Focus_Fund_Flow         = pd.concat([Promotion_df,Focus_df])
print('Focus_Fund_Flow:',Focus_Fund_Flow)
Focus_Fund_Flow_Groupby = Focus_Fund_Flow.groupby(['index','Type'])[['Inflow','Outflow','Net Flow','New Money']].agg('sum').sort_values(by='Type').reset_index()

list_5        = pd.DataFrame( [['Total','-','-','-','-','-','-']] , columns=Focus_Fund_Flow.columns )
sapce_df      = pd.DataFrame(np.zeros((2,Focus_Fund_Flow.shape[1])),columns=Focus_Fund_Flow.columns)
sapce_df      = sapce_df.replace(0,"")


Focus_Fund_Flow = pd.concat([Focus_Fund_Flow,sapce_df,list_5,Focus_Fund_Flow_Groupby]).fillna(value="-")
Focus_Fund_Flow 


#%%


output_path = r'D:\My Documents\andyhs\桌面\Andy\業務例行報表\output\YTD_Monthly_Sales_Performance_Statistic.xlsx'

excel = Excel_Work()
wb=excel.write_excel(df=Agent_AUM_df,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
wb.save(output_path)

excel.append_excel(excel_path=output_path , df=Agent_onshore_df                   , excel_sheet_name="Agent Onshore"                   )
excel.append_excel(excel_path=output_path , df=Agent_offshore_df                  , excel_sheet_name="Agent Offshore"                  )
excel.append_excel(excel_path=output_path , df=Offshore_Department_Product_Output , excel_sheet_name="Offshore By Company_Department"  )
excel.append_excel(excel_path=output_path , df=Fund_Flow                          , excel_sheet_name="Top 5 Fund Flow"                 )
print('Task Complete !')


# %%
