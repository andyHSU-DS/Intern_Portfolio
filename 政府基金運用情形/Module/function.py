import PyPDF2
from matplotlib.pyplot import get
import pdfplumber as pr
import pandas as pd
from openpyxl import load_workbook
import re
import os
from datetime import datetime

input_path = r'./input'

#抓內含.pdf的檔案
def get_PDF_files():
    PDF_files = []
    for x in os.listdir(input_path):
        if '.pdf' in x:
            PDF_files.append(x)
    return PDF_files

#讀這些檔案
def files_read(PDF_lists):
    for PDF_file in PDF_lists:
        if '勞工退休基金' in PDF_file:
            勞退_pdf = pr.open(input_path + '/' +PDF_file)
        elif '勞工保險基金' in PDF_file:
            勞保_pdf = pr.open(input_path + '/' +PDF_file)
        elif '國民年金保險基金' in PDF_file:
            國保_pdf = pr.open(input_path + '/' +PDF_file)
        elif '110年11月基金運用情形' in PDF_file:
            退撫_pdf = pr.open(input_path + '/' +PDF_file)
    return 勞退_pdf, 勞保_pdf, 國保_pdf, 退撫_pdf

#新舊index轉換
def get_index():
    for x in os.listdir(input_path):
        if '政府基金運用資訊' in x and '.xlsx' in x:
            GFI = pd.read_excel(input_path+'/'+x,sheet_name = 'Government Fund Info (Raw Data)',skiprows=2)
            ##把index處理一下
            new_index = []
            old_index = GFI.set_index('Unnamed: 0').index
            #print(old_index)
            for i in GFI.set_index('Unnamed: 0').index:
                #print(i)
                if re.match('[\u4e00-\u9fa5]+',i):
                    new_index.append(re.match('[\u4e00-\u9fa5]+',i)[0])
    return new_index, old_index

#輸入幫我們讀取所有的圖表
def get_all_table(file):
    tables_dict = {}
    ps = file.pages
    #print(len(ps))
    #print(ps)
    i = 0
    for p in range(len(ps)):
        #進去每一頁
        pg = ps[p]
        #找到有多少表格
        tables = pg.extract_tables()
        #一頁只有一個表格的狀況
        if len(tables) == 1:
            table = tables
            tables_dict['表格'+str(i)] = table
            i += 1
        #一頁有超過一個表格的狀況
        elif len(tables) > 1:
            for g in range(len(tables)):
                table = tables[g]
                tables_dict['表格'+str(i)] = table
                i += 1
    return tables_dict

#依照我們設定的行數來找尋目標table
def get_target_table(tables_dict,target_len):
    i = 0
    #用來放目標table的
    target_table_df = pd.DataFrame({})
    #用表格的長度來決定要抓什麼表格
    for table in list(tables_dict.values()):
        if len(table) == target_len:
            #print(i)
            #print(pd.DataFrame(table[1:],columns = table[0]))
            Table = pd.DataFrame(table[1:],columns = table[0])
            target_table_df = pd.concat([target_table_df,Table],axis=1)
            
        #這邊是發現有的表格會變成三維的，所以做特殊處理
        elif len(table[0]) == target_len:
            #print(i)
            #print(pd.DataFrame(table[0][1:],columns = table[0][0]))
            Table = pd.DataFrame(table[0][1:],columns = table[0][0])
            #因為項目我已經有了，所以要拿掉
            target_table_df = pd.concat([target_table_df,Table],axis=1)
        i += 1
    return target_table_df

#因為最後excel跟PDF的欄名可能不一樣，我們這邊用DICT的方式做一個轉換
def convert_勞退_dict(x):
    勞退_dict = {
    '公債、公司債、金融債券及特別股':'公債',
    '股票及受益憑證投資（含期貨）':'股票及受益憑證投資',
    '國內委託經營':'國內委託',
    '國外委託經營':'國外委託',
    '合           計':'金額總計'
    }
    if x in 勞退_dict:
        result = 勞退_dict[x]
    else:
        result = x
    return result

def convert_勞保_dict(x):
    勞保_dict = {
    '公債、公司債、金融債券及特別股':'公債',
    '股票及受益憑證投資（含期貨）':'股票及受益憑證投資',
    '國內委託經營':'國內委託',
    '國外委託經營':'國外委託',
    '合    計':'金額總計'
    }
    if x in 勞保_dict:
        result = 勞保_dict[x]
    else:
        result = x
    return result

def convert_國保_dict(x):
    國保_dict = {
    '合           計':'金額總計'
    }
    if x in 國保_dict:
        result = 國保_dict[x]
    else:
        result = x
    return result

def convert_退撫_dict(x):
    退撫_dict = {
    '自行運用小計':'自行運用',
    '國內委託經營':'國內委託',
    '國外委託經營':'國外委託',
    '委託經營小計':'委託經營',
    '合                 計':'金額總計'
}
    if x in 退撫_dict:
        result = 退撫_dict[x]
    else:
        result = x
    return result


#放到主要程式碼時，前面要加一個判斷這個檔案的名稱(勞退勞保or年金等)，來決定後續的動作
def handle_勞退_process(DataFrame):
    DataFrame.columns = ['項         目','勞退新制\nLPF','占基金運用比率（％）新','項         目','勞退舊制\nLRF','占基金運用比率（％）舊']
    #去掉欄重複的
    新舊制勞退_df = DataFrame.loc[:,~DataFrame.columns.duplicated()]
    新舊制勞退_df['項         目'] = 新舊制勞退_df['項         目'].apply(lambda x:convert_勞退_dict(x))   
    new_DF = pd.DataFrame(['-']*len(old_index),index = old_index,columns = ['對照欄位(之後要移除)'])
    new_DF['對照欄位(之後要移除)'] = new_index
    new_DF.reset_index()
    勞退完成_df = pd.merge(left = new_DF.reset_index(),right = 新舊制勞退_df,left_on = '對照欄位(之後要移除)',right_on = '項         目',how='outer')
    #這邊看之後能不能改的比較自動
    勞退完成_df = 勞退完成_df.iloc[[0,1,2,3,4,5,6,7,8,9,10,11,15,19,23,24,25,12,16,20,26],[0,3,4,5,6]]
    勞退完成_df.drop_duplicates(inplace=True)
    勞退完成_df.reset_index(drop=True,inplace=True)
    #加上%符號
    勞退完成_df['占基金運用比率（％）新'] = 勞退完成_df['占基金運用比率（％）新'] + '%'
    勞退完成_df['占基金運用比率（％）舊'] = 勞退完成_df['占基金運用比率（％）舊'] + '%'
    return 勞退完成_df

def handle_勞保_process(DataFrame):
    DataFrame.columns = ['項         目','勞保\nLIF','%']
    #去掉欄重複的
    #這邊不會用到
    勞保_df = DataFrame.loc[:,~DataFrame.columns.duplicated()]
    勞保_df['項         目'] = 勞保_df['項         目'].apply(lambda x:convert_勞保_dict(x))   
    new_DF = pd.DataFrame(['-']*len(old_index),index = old_index,columns = ['對照欄位(之後要移除)'])
    new_DF['對照欄位(之後要移除)'] = new_index
    new_DF.reset_index()
    勞保完成_df = pd.merge(left = new_DF.reset_index(),right = 勞保_df,left_on = '對照欄位(之後要移除)',right_on = '項         目',how='outer')
    #這邊看之後能不能改的比較自動
    勞保完成_df.drop_duplicates(inplace=True)
    勞保完成_df_1 = 勞保完成_df.reset_index(drop=True)
    勞保完成_df_1 = 勞保完成_df_1.iloc[[0,1,2,3,4,5,6,7,8,9,10,11,13,15,17,18,19,12,14,16,20],[0,3,4]]
    #加上%符號
    勞保完成_df_1['%'] = 勞保完成_df_1['%'] + '%'
    return 勞保完成_df_1

def handle_國保_process(DataFrame):
    DataFrame.columns = ['項         目','國保基金\nNPIF','%']
    #去掉欄重複的
    #這邊不會用到
    國保_df = DataFrame.loc[:,~DataFrame.columns.duplicated()]
    國保_df['項         目'] = 國保_df['項         目'].apply(lambda x:convert_國保_dict(x))  
    new_DF = pd.DataFrame(['-']*len(old_index),index = old_index,columns = ['對照欄位(之後要移除)'])
    new_DF['對照欄位(之後要移除)'] = new_index
    new_DF.reset_index()
    國保完成_df = pd.merge(left = new_DF.reset_index(),right = 國保_df,left_on = '對照欄位(之後要移除)',right_on = '項         目',how='outer')
    #這邊看之後能不能改的比較自動
    國保完成_df.drop_duplicates(inplace=True)
    國保完成_df_1 = 國保完成_df.reset_index(drop=True)
    國保完成_df_1 = 國保完成_df_1.iloc[[0,1,2,3,4,5,6,7,8,9,10,11,13,15,17,18,19,12,14,16,20],[0,3,4]]
    #加上%符號
    國保完成_df_1['%'] = 國保完成_df_1['%'] + '%'
    return 國保完成_df_1

def handle_退撫_process(DataFrame):
    DataFrame.columns = ['項         目','退撫基金\nPSPF','%']
    #去掉欄重複的
    #這邊不會用到
    退撫_df = DataFrame.loc[:,~DataFrame.columns.duplicated()]
    退撫_df['項         目'] = 退撫_df['項         目'].apply(lambda x:convert_退撫_dict(x))  
    new_DF = pd.DataFrame(['-']*len(old_index),index = old_index,columns = ['對照欄位(之後要移除)'])
    new_DF['對照欄位(之後要移除)'] = new_index
    new_DF.reset_index()
    退撫完成_df = pd.merge(left = new_DF.reset_index(),right = 退撫_df,left_on = '對照欄位(之後要移除)',right_on = '項         目',how='left')
    #這邊看之後能不能改的比較自動
    #這邊一定不要做drop_duplicates，因為都沒有資料，刪掉會出錯
    #退撫完成_df.drop_duplicates(inplace=True)
    退撫完成_df_1 = 退撫完成_df.reset_index(drop=True)
    退撫完成_df_1 = 退撫完成_df_1.iloc[:,[0,3,4]]
    #加上%符號
    退撫完成_df_1['%'] = 退撫完成_df_1['%'] + '%'
    return 退撫完成_df_1

#整合的function來結合get_all_table及後續的處理資料
def get_final_info(pdf_file,pdf_kind,target_table_line=0):
    print(pdf_kind)
    if target_table_line == 0:
        if '勞退' in pdf_kind:
            target_table_line = 17
        elif '勞保' in pdf_kind:
            target_table_line = 20
        elif '國保' in pdf_kind:
            target_table_line = 18
        elif '退撫' in pdf_kind:
            target_table_line = 14
    else:
        target_table_line = target_table_line
    #找到file內的所有table
    pdf_file_all_table = get_all_table(pdf_file)
    pdf_file_df  = get_target_table(pdf_file_all_table,target_table_line)
    #處理完後會用iloc是因為，要把不用的欄位去掉
    if '勞退' in pdf_kind:
        df_output = handle_勞退_process(pdf_file_df)
    elif '勞保' in pdf_kind:
        df_output = handle_勞保_process(pdf_file_df).iloc[:,1:]
    elif '國保' in pdf_kind:
        df_output = handle_國保_process(pdf_file_df).iloc[:,1:]
    elif '退撫' in pdf_kind:
        df_output = handle_退撫_process(pdf_file_df).iloc[:,1:]
    else:
        df_output = pd.DataFrame({})
    return df_output


if __name__ == '__main__':
    new_index, old_index = get_index()
    PDF_lists = get_PDF_files()
    勞退_pdf, 勞保_pdf, 國保_pdf, 退撫_pdf = files_read(PDF_lists)
    #勞退_pdf
    勞退_final_df = get_final_info(勞退_pdf,'勞退')
    #勞保_pdf
    勞保_final_df = get_final_info(勞保_pdf,'勞保')
    #國保_pdf
    國保_final_df = get_final_info(國保_pdf,'國保')
    #退撫_pdf
    退撫_final_df = get_final_info(退撫_pdf,'退撫')
    
    final_df = pd.concat([勞退_final_df,勞保_final_df,國保_final_df,退撫_final_df],axis = 1)
    output_path = r'./output'
    file_name_date = str(datetime.now().year)+' '+str(datetime.now().month)
    final_df.to_csv(output_path+'/政府基金資訊'+file_name_date+'.csv',encoding='utf-8-sig',index=False,header=True)
