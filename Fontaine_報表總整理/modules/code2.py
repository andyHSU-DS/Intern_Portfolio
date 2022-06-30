import os 
import numpy as np 
import pandas as pd 


def each_file(df):
    list_number = df["戶號"].to_list()
    list_number = list( set(list_number) )
    return list_number
## 算出所有df裡所有的戶號
def file_case(file_case_path):
    final_list=[]
    curr_path = os.getcwd()
    data_path = os.path.join( curr_path , file_case_path )
    months = os.listdir(data_path)
    for i in range( len(months) ) :
        excel_path=data_path+"\\"+str(months[i])
        df = pd.read_excel(excel_path)
        if ( ("關聯戶報表" in df.columns[0]) or ("（單一報表）" in df.columns[0])):
            df = pd.read_excel(excel_path,skiprows=4)
            if("查核項目" in df.columns[0] or "戶號" not in df.columns):
                pass
            else:
                final_list.extend( each_file(df) )
        else :
            if("查核項目" in df.columns[0] or "戶號" not in df.columns):
                pass
            else:
                final_list.extend( each_file(df) )
    final_list = list(set(final_list))
    number = len(final_list)
    return final_list,number 
## screening all the file and get total monitor transaction number
def total_交易監控數(file_case_path):
    """
    R_list      : 關聯戶 的FILE裡有多少筆交易監控數
    number_list : 單一帳戶 的FILE裡有多少筆交易監控數
    r_name      : 關聯戶的 file name 
    name        : 單一帳戶的 file name 
    要 name_list 是因為 要對 FILE , 看是哪裡數字有問題 

    [ Condition Explain ]

    因為 excel_file 每年可能長的不一樣.....
    所以條件設的有點多首先

    [第1個condition]
    if :
    ( "關聯戶報表" in df.columns[0] or "（單一報表）" in df.columns[0]) 
    --> 表示是新版的excel file , have to skip row * 4 去讀 excel file
    else :
    可以從第一行開始讀

    [第2個condition]
    分出 單一帳戶 和 關聯戶的 ---> 用filename 去判斷, 有沒有R而已
    
    [第3個condition]
    判斷內部是否有交易監控查核數字 , 有的append df.shape[0] , 沒有的append n=0
    
    """
    R_list = []
    number_list = []
    r_name=[]
    name=[]
    # 讀 certain year 裡資料夾內的所有檔案
    curr_path = os.getcwd()
    data_path = os.path.join(curr_path , file_case_path)
    months = os.listdir(data_path)
    for i in range(len(months)):
        excel_path=data_path+"\\"+str(months[i])
        df = pd.read_excel(excel_path)
        if ( "關聯戶報表" in df.columns[0] or "（單一報表）" in df.columns[0]):
            df = pd.read_excel(excel_path,skiprows=4)
            if ( "R" in str(months[i]) ):
                if ( ("查核原因說明" in df.columns) and ( "依查詢條件查無相關符合資料" in df["查核原因說明"].to_list() ) ):
                    n=0
                    R_list.append(n)
                elif( 'DTI_NO' in df.columns):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])
                elif ("查核項目" in df.columns[0] ):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])                   
                elif ( ("戶號" in df.columns) and ( df["戶號"].shape[0]==0 or np.isnan(df["戶號"][0]) == True  ) ):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])                 
                else:
                    n=df.shape[0]
                    R_list.append(n)
                    r_name.append(months[i])
            else :
                if ( ("查核原因說明" in df.columns) and ( "依查詢條件查無相關符合資料" in df["查核原因說明"].to_list() ) ):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                elif( 'DTI_NO' in df.columns):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                elif ("查核項目" in df.columns[0] ):
                    n=0
                    number_list.append(n) 
                    name.append(months[i])                  
                elif ( ("戶號" in df.columns) and ( df["戶號"].shape[0]==0 or np.isnan(df["戶號"][0]) == True  ) ):
                    n=0
                    number_list.append(n)    
                    name.append(months[i])               
                else:
                    n=df.shape[0]
                    number_list.append(n)
                    name.append(months[i])
        else:
            if ( "R" in str(months[i]) ):
                if ( ("查核原因說明" in df.columns) and ( "依查詢條件查無相關符合資料" in df["查核原因說明"].to_list() ) ):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])
                elif( 'DTI_NO' in df.columns):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])
                elif ("查核項目" in df.columns[0] ):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])
                elif (("戶號" in df.columns) and ( df["戶號"].shape[0]==0 or np.isnan(df["戶號"][0]) == True  )):
                    n=0
                    R_list.append(n)
                    r_name.append(months[i])
                else:
                    n=df.shape[0]
                    R_list.append(n)
                    r_name.append(months[i])
            else :
                if ( ("查核原因說明" in df.columns) and ( "依查詢條件查無相關符合資料" in df["查核原因說明"].to_list() ) ):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                elif( 'DTI_NO' in df.columns):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                elif ("查核項目" in df.columns[0] ):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                elif (("戶號" in df.columns) and ( df["戶號"].shape[0]==0 or np.isnan(df["戶號"][0]) == True  )):
                    n=0
                    number_list.append(n)
                    name.append(months[i])
                else:
                    n=df.shape[0]
                    number_list.append(n)
                    name.append(months[i])
    return R_list,number_list,r_name,name

def get_交易監控(file_path):
    """
    [1]
    list_1 ,list_2  -> 是單一帳戶,關聯戶的交易監控數 ,return list 
    ---> sum(list) 可以得到該 檔案 以及  該 year 的交易監控總數 
    [2]
    r_name , name   ->  單一帳戶,關聯戶所讀的檔案 , return list 
    ----> 方便看說有沒有錯 , 如果數字有問題
    [3]
    final_list -> 所有的客戶的編號,
    已經做過 list(set(list)) , 不會有重複的客戶 ,return list 
    final_number -> len(final_list) 
    得到該year  總共是有幾個客戶數
    """
    list_1, list_2 ,r_name, name = total_交易監控數(file_case_path = file_path)
    final_list , final_number = file_case(file_case_path = file_path)
    return list_1 , list_2 , r_name , name , final_list, final_number

def to_df(year,list_1,list_2,final_number):
    """
    整理成 df
    """
    df=pd.DataFrame()
    final_list = [sum(list_1),sum(list_2),sum(list_1)+sum(list_2) , final_number]
    df[str(year)] = final_list
    df_index = ["關聯戶統計","單一帳戶統計","交易間控警示總筆數","交易監控警示總戶數"]
    df.index =df_index
    k = df.loc["交易間控警示總筆數"].values / df.loc["交易監控警示總戶數"].values 
    df.loc["平均”每戶”出現的筆數"] = k
    df = df .reset_index()
    return df 
