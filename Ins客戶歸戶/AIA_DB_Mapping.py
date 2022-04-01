#%%

import os 
import re 
import numpy as np 
import pandas as pd 
import matplotlib.pyplot as plt
from module import Excel_Work 


# AIA-DB Mapping .
# 抓完月報資料後 , AIA List / top ILP madnate AUM 做個 AUM Mapping .
# 無法完全 Mapping , 抓完之後還是要手動調整 .


# %%

curr_path = os.getcwd()
data_path = os.path.join(curr_path , r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\Mapping_Data\DB_Data")
months    = os.listdir(data_path)
    
for i in range(len(months)):
    excel_path=data_path+"\\"+str(months[i])
    if "offshore" in months[i] : 
        offshore_account_data = pd.read_excel(excel_path,sheet_name="By Account")
        offshore_account_data = offshore_account_data[['客戶姓名','DB AUM','NAMU',"NN",'IAM']]
        
        offshore_account_data['DB AUM'] *= 28
        offshore_account_data['NAMU']   *= 28
        offshore_account_data['NN']     *= 28
        offshore_account_data['IAM']    *= 28
    elif "onshore" in months[i] :
        onshore_account_data = pd.read_excel(excel_path,sheet_name="By Account")
        onshore_account_data = onshore_account_data[['客戶姓名','DB AUM']]



onshore_account_data['NAMU']   = 0
onshore_account_data['NN']     = 0
onshore_account_data['IAM']    = 0

onshore_account_data['From'] = 'Onshore'
offshore_account_data['From'] = 'Offshore'

DB_Data = pd.concat([onshore_account_data,offshore_account_data])
DB_Data 

# %%

Manual_Data = pd.read_excel(r'D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\2022_Feb_ILP\Feb_Top49全委帳戶.xlsx',sheet_name='Account TOP 5 Holding')
Manual_Data = Manual_Data[['From','Account','目前規模(新台幣)']]
Manual_Data

# %%

DB_Acount_list    = DB_Data['客戶姓名'].to_list()
DB_Account_AUM    = DB_Data['DB AUM'].to_list()
DB_Account_Source = DB_Data['From'].to_list()
DB_Acount_NAMU    = DB_Data['NAMU'].to_list()
DB_Acount_NN      = DB_Data['NN'].to_list()
DB_Acount_IAM     = DB_Data['IAM'].to_list()



Manual_Acount_list   = Manual_Data['Account'].to_list()
Manual_Acount_AUM    = Manual_Data['目前規模(新台幣)'].to_list()
Manual_Acount_Source = Manual_Data['From'].to_list()


#%%


def process_account(account):

    
    # ----------------------------------------
    if "安聯人壽委託德盛安聯投信投資帳戶-台幣環球股債均衡組合(月撥回資產)" in account :
        account  = "安聯人壽委託安聯投信投資帳戶－台幣環球股債均衡組合"
        return account

    elif "三商美邦人壽鑫穩健投資帳戶(現金撥回)" in account :
        account  = "三商美邦人壽鑫穩健投資帳戶－全權委託安聯投信投資帳戶"
        return account

    elif "台灣人壽委託群益投信投資帳戶-安鑫增益投資帳戶(新臺幣)" in account :
        account  = "台灣人壽委託群益投信投資帳戶—安鑫增益投資帳戶（新台幣）"
        return account

    elif "**台灣人壽台幣代操帳戶(成長型)" in account :
        account  = "台灣人壽委託宏利投信—台幣投資帳戶（成長型）（一）"
        return account

    elif "安聯人壽委託復華投信投資帳戶-豐收得利2(月撥回資產)" in account :
        account  = "安聯人壽委託復華投信投資帳戶一豐收得利２（月撥回資產）"
        return account

    elif "**台灣人壽委託宏利投信-台幣投資帳戶(成長型)(二)" in account :
        account  = "台灣人壽委託宏利投信—台幣投資帳戶（成長型）（ＩＩ）"
        return account

    elif "*安聯人壽委託富蘭克林華美投信投資帳戶_新臺幣多元收益" in account :
        account  = "安聯人壽委託富蘭克林華美投信投資帳戶_新臺幣多元收益(月撥回)"
        return account

    elif "全球人壽優選樂退投資帳戶" in account :
        account  = "全球人壽優選樂退投資帳戶〈委託群益投信運用操作〉"
        return account

    elif "瀚亞新收益全權委託管理帳戶(美元)" in account :
        account  = "瀚亞新收益全權委託管理帳戶-中國人壽"
        return account

    # -----------------------
    
    elif "-" in account and "投資帳戶" in account and "(委託" in account :
        account = account.replace(" ","")
        index = account.index("-")
        account = account[:index]
        return account

    elif "*" in account and "-" in account and "投資帳戶" in account and "(委託" in account :
        account = account.replace(" ","")
        index_ = account.index("*")
        index = account.index("-")
        account = account[index_:index]
    
    elif "**" in account :
        account = account.replace(" ","")
        account = account[2:]
        return account

    elif "*" in account :
        account = account.replace(" ","")
        account = account[1:]
        return account
    
    elif "(現金撥回)" in account :
        account = account.replace(" ","")
        index = account.index("(現金撥回)")
        account  = account[:index]
        return account
    
    elif "月撥回" in account :
        account = account.replace(" ","")
        index = account.index("月撥回")
        account  = account[:index]
        return account

    elif "（月撥回資產）" in account :
        account = account.replace(" ","")
        index = account.index("（月撥回資產）" )
        account  = account[:index]
        return account


    elif "（新台幣）" in account :
        account = account.replace(" ","")
        index = account.index("（新台幣）")
        account  = account[:index]
        return account

    elif "－" in account :
        account = account.replace(" ","")
        account = account.replace("－","-")
        return account
        
    elif "一" in account :
        account = account.replace(" ","")
        account = account.replace("一","-")
        return account

    else:
        return account



# %%


from sklearn.feature_extraction.text import CountVectorizer 

def jaccard_similarity(s1, s2): 
    
    def add_space(s):
        return ' '.join(list(s)) # 將字中間加入空格

    s1, s2 = add_space(s1), add_space(s2) # 轉化為TF矩陣 
    
    cv = CountVectorizer(tokenizer=lambda s: s.split()) 
    
    corpus = [s1, s2]
    
    vectors = cv.fit_transform(corpus).toarray() # 獲取詞表內容 
    
    ret = cv.get_feature_names()
    
    # print(ret) # 求交集 
    
    numerator = np.sum(np.min(vectors, axis=0)) # 求並集 
    
    denominator = np.sum(np.max(vectors, axis=0)) # 計算傑卡德係數 
    
    return 1.0 * numerator / denominator 


output = Manual_Data

output["DB Account"] = ""
output["DB Account Size"] = 0
output['NAMU']   = 0
output['NN']     = 0
output['IAM']    = 0

for i in range(output.shape[0]):
    account = output['Account'][i]
    account_source = output['From'][i]
    if "onshore" in account_source:
        for j in range(len(DB_Acount_list)):
            db_account = DB_Acount_list[j]
            db_size = DB_Account_AUM [j]
            if jaccard_similarity(process_account(account),process_account(db_account) ) >= 0.85 :
                print(f"-----score : {jaccard_similarity(account, db_account )}---------")
                print(account)
                print(db_account)
                output["DB Account"][i] = db_account
                output["DB Account Size"][i] = db_size
            else:
                pass

    else:
        for j in range(len(DB_Acount_list)):
            db_account = DB_Acount_list[j]
            db_size = DB_Account_AUM [j]
            db_namu = DB_Acount_NAMU[j]
            db_nn = DB_Acount_NN[j]
            db_iam = DB_Acount_IAM[j]
            if jaccard_similarity(process_account(account),process_account(db_account) ) >= 0.85 :
                print(f"-----score : {jaccard_similarity(account, db_account )}---------")
                print(account)
                print(db_account)
                output["DB Account"][i] = db_account
                output["DB Account Size"][i] = db_size
                output['NAMU'][i]   = db_namu
                output['NN'][i]     = db_nn
                output['IAM'][i]    = db_iam
            else:
                pass

output['DB %'] = output['DB Account Size'] / output['目前規模(新台幣)'] *100
output['DB %'] = output['DB %'].apply(lambda x:str(round(x,2))+'%')
excel = Excel_Work()
wb=excel.write_excel(df=output,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
wb.save(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\mapping_output.xlsx")

# %%

