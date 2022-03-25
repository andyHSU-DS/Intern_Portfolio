import pandas as pd
import os 
from tqdm import tqdm
import numpy as np
def Joan_客戶歸戶():
    file_path = os.getcwd()+'/input'
    files     = os.listdir(file_path)
    for file in files:
        if "Onshore" in file:
            Onshore  = pd.read_excel(file_path+"/"+file,skiprows=3)
        elif "Offshore" in file:
            Offshore = pd.read_excel(file_path+"/"+file,skiprows=3)
    #需要的欄位
    col_need = ['通路','基金簡稱','客戶','客戶姓名','結存-月底AUM(只有迄月資料)','Sales姓名']

    Onshore   = Onshore[col_need]
    Onshore   = Onshore[Onshore['通路'] == 4]
    Onshore['Current_Rate']   = 1
    Offshore  = Offshore[col_need]
    Offshore  = Offshore[Offshore['通路'] == 4]
    Offshore['Current_Rate']   = 28
    
    Total_shore  = pd.concat([Onshore,Offshore],axis=0)

    Total_shore['結存-月底AUM(只有迄月資料)[NTD]'] = Total_shore['結存-月底AUM(只有迄月資料)']*Total_shore['Current_Rate']

    #做一個判斷是不是NMMF的function                                                                                                                                                                                                                  
    def NMMF(x):
        if "貨幣市場" in x:
            result = "MMF"
        else:
            result = "NMMF"
        return result

    def NMMF_grade(x):
        if ~np.isnan(x):
            if x<3000000:
                return '小於$3,000,000'
            elif 5000000>x>=300000:
                return '大於$3,000,000\n小於$5,000,000'
            elif 10000000>x>=500000:
                return '大於$5,000,000\n小於$10,000,000'
            elif 30000000>x>=1000000:
                return '大於$10,000,000\n小於$30,000,000'
            elif 100000000>x>=3000000:
                return '大於$30,000,000\n小於$100,000,000'
            else:
                return '大於$100,000,000'
        else:
            return '無NMMF'

    Total_shore['MMF/NMMF'] = Total_shore['基金簡稱'].apply(lambda x:NMMF(x))
    Total_shore['客戶'] = Total_shore['客戶'].astype('str')
    
    #客戶有包含_2
    Total_shore_2      = Total_shore[Total_shore['客戶'].str.contains('_2')]

    #客戶沒有包含_2
    Total_shore_normal = Total_shore[~Total_shore['客戶'].str.contains('_2')].reset_index(drop=True)
    Total_shore_normal['客戶'] = Total_shore_normal['客戶'].str.replace('_3','').reset_index(drop=True)

    for index,df in enumerate([Total_shore_normal, Total_shore_2]):
        df_gp = df.groupby(['客戶','客戶姓名','基金簡稱','Sales姓名','MMF/NMMF']).agg({'結存-月底AUM(只有迄月資料)[NTD]':'sum'})
        df_gp = df_gp.unstack()
        df_gp.reset_index(inplace=True)
        df_gp['NMMF等級'] = df_gp['結存-月底AUM(只有迄月資料)[NTD]']['NMMF'].apply(lambda x:NMMF_grade(x))
        客戶等級依照業務分類 = pd.concat([df_gp['Sales姓名'],df_gp['NMMF等級']],axis=1).groupby(['Sales姓名','NMMF等級']).agg({'NMMF等級':'count'})

        ###以下完全是為了產出csv好看的做法，要對產出內容作改變請寫在上面
        #戶號姓名都相同
        need_space_list = []
        for i in range(1,len(df_gp)):
            if df_gp.iloc[i,0] == df_gp.iloc[i-1,0] and df_gp.iloc[i,1] == df_gp.iloc[i-1,1]:
                need_space_list.append(i)
        #戶號相同，姓名不同
        need_account_number_space_list = []
        for i in range(1,len(df_gp)):
            if df_gp.iloc[i,0] == df_gp.iloc[i-1,0] and df_gp.iloc[i,1] != df_gp.iloc[i-1,1]:
                need_space_list.append(i)
        
        print(len(need_space_list))
        for i in tqdm(need_space_list):
            df_gp.iloc[i,[0,1]]=['','']

        print(len(need_account_number_space_list))
        for i in tqdm(need_account_number_space_list):
            df_gp.iloc[i,0]=['']
            
        NMMF_grade_stat = pd.DataFrame(df_gp['NMMF等級'].value_counts())

        if index == 1:
            df_gp.to_csv(r'output/戶號含_2的客戶資料.csv',encoding='utf-8-sig',index=True,header=True)
            NMMF_grade_stat.to_csv(r'output/戶號含_2的客戶NMMF等級分類.csv',encoding='utf-8-sig',index=True,header=True)
            客戶等級依照業務分類.to_csv(r'output/戶號含_2的客戶NMMF等級依業務分類.csv',encoding='utf-8-sig',index=True,header=True)
            
        else:
            df_gp.to_csv(r'output/正常客戶的歸戶資料.csv',encoding='utf-8-sig',index=True,header=True)
            NMMF_grade_stat.to_csv(r'output/正常客戶NMMF等級分類.csv',encoding='utf-8-sig',index=True,header=True)
            客戶等級依照業務分類.to_csv(r'output/正常客戶NMMF等級依業務分類.csv',encoding='utf-8-sig',index=True,header=True)
            

    



if __name__ == '__main__':
    Joan_客戶歸戶()



    

    
