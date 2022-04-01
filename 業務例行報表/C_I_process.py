import pandas as pd
import os
import numpy as np
import re
from Module import Excel_Work
###此份CODE
# 目的:讓C&I處理上更加方便
# 效果:C&I前幾分頁('月底 AUM','AVG AUM','一般申購','匯出/轉申購','淨流入','手續費','Revenue','管理費')
# note:
# 1. 聯繫紀錄和customer by AUM 不需要做整理
# 2. 先跑完YTD, MTD的巨集再執行，這很重要
def C_I_process():
    #找到ouput 資料夾
    file_path = os.getcwd() + '\output'
    files = os.listdir(file_path)
    print(file_path)
    for file in files:
        if 'MTD' in file:
            MTD = file_path + '/' +file
        elif 'YTD' in file:
            YTD = file_path + '/' + file
    
    def C_I_前8分頁month():

        MTD_page_1_Total    = pd.read_excel(MTD,sheet_name='Sheet').set_index('Department').loc['C&I Taipei 1':'C&I Total'].reset_index()

        MTD_page_1_Onshore  = pd.read_excel(MTD,sheet_name='Agent Onshore').set_index('Section').loc['C&I Taipei 1':'C&I Total'].reset_index() 

        MTD_page_1_Offshore = pd.read_excel(MTD,sheet_name='Agent Offshore').set_index('Section').loc['C&I Taipei 1':'C&I Total'].reset_index() 

        total_df = pd.DataFrame({})

        #將col name進行修改
        for sheet in [MTD_page_1_Total, MTD_page_1_Onshore, MTD_page_1_Offshore]:
            sheet_new_col = []
            for sheet_col in sheet.columns:
                if "onshore" in sheet_col or "offshore" in sheet_col:
                    sheet_col = sheet_col.replace('onshore ','').replace('offshore ','')
                elif "Section" in sheet_col:
                    sheet_col = sheet_col.replace('Section','Department')
                sheet_new_col.append(sheet_col)
            sheet.columns = sheet_new_col
            total_df      = total_df.append(sheet)

        return total_df
    
    def C_I_New_Open():
        MTD_page_New_Open_Total    = pd.read_excel(MTD,sheet_name='New open account').loc[:,['Total Individual','Total Corporate']].values
        MTD_page_New_Open_Onshore  = pd.read_excel(MTD,sheet_name='New open account').loc[:,['Onshore Individual','Onshore Corporate']].values
        MTD_page_New_Open_Offshore = pd.read_excel(MTD,sheet_name='New open account').loc[:,['Offshore Individual','Offshore Corporate']].values
        New_Open_分頁  = pd.DataFrame(np.vstack([MTD_page_New_Open_Total,MTD_page_New_Open_Onshore,MTD_page_New_Open_Offshore]),columns = ['Individual','corporate'])
        return New_Open_分頁

    def C_I_top5_Fund():
        #針對MTD top5 fund操作
        #col 1 是index，我不需要
        MTD_page_Top5_Fund    = pd.read_excel(MTD,sheet_name='Top 5 Fund Flow').iloc[:,1:].dropna(how='all',axis=0)
        MTD_page_Top5_Fund    = MTD_page_Top5_Fund[MTD_page_Top5_Fund['Top 5 Inflow 基金'] != '-']
        for i,row in enumerate(MTD_page_Top5_Fund.values):
            for z,row_each_value in enumerate(row):
                #找到Top 5開頭 前面加上 MTD
                if re.match('Top 5',str(row_each_value)):
                    MTD_page_Top5_Fund.values[i][z] = re.sub('\ATop 5','MTD Top 5',str(row_each_value))

        #針對YTD top5 fund操作
        #col 1 是index，我不需要
        YTD_page_Top5_Fund    = pd.read_excel(YTD,sheet_name='Top 5 Fund Flow').iloc[:,1:].dropna(how='all',axis=0)
        YTD_page_Top5_Fund    = YTD_page_Top5_Fund[YTD_page_Top5_Fund['Top 5 Inflow 基金'] != '-']
        #將col放進array這樣它會出現在row1而非col
        YTD_page_Top5_Fund_new = np.insert(YTD_page_Top5_Fund.values,0,YTD_page_Top5_Fund.columns.values,axis=0)
        for i,row in enumerate(YTD_page_Top5_Fund_new):
            for z,row_each_value in enumerate(row):
                #找到Top 5開頭 前面加上 YTD
                if re.match('Top 5',str(row_each_value)):
                    YTD_page_Top5_Fund_new[i][z] = re.sub('\ATop 5','YTD Top 5',str(row_each_value))

        total_top5_fund = np.vstack([MTD_page_Top5_Fund.values,YTD_page_Top5_Fund_new])
        top5_fund_df    = pd.DataFrame(total_top5_fund)
        return top5_fund_df 
        
    def C_I_Promo_and_Focus_Fund():
        MTD_page_Top5_Fund    = pd.read_excel(MTD,sheet_name='Promotion and Focus Fund')
        #因為我們只要total之後的資料，所以先找到total的index
        def get_index_after_total():
            for i in range(len(MTD_page_Top5_Fund)):
                if MTD_page_Top5_Fund.iloc[i,0] == 'Total':
                    break
            return i+1
        
        #我們要的index
        correct_index = get_index_after_total()
        correct_data  = MTD_page_Top5_Fund.iloc[correct_index:,[0,2,3,4,5]]
        return correct_data
    
    def three_to_eight_page_YTD():
        need_index=[0,1,4,5,6,7,8,9,10]
        YTD_page_1_Total    = pd.read_excel(YTD,sheet_name='Sheet').set_index('Department').loc['C&I Taipei 1':'C&I Total'].reset_index().iloc[:,need_index]

        YTD_page_1_Onshore  = pd.read_excel(YTD,sheet_name='Agent Onshore').set_index('Section').loc['C&I Taipei 1':'C&I Total'].reset_index().iloc[:,need_index]

        YTD_page_1_Offshore = pd.read_excel(YTD,sheet_name='Agent Offshore').set_index('Section').loc['C&I Taipei 1':'C&I Total'].reset_index().iloc[:,need_index] 

        total_df = pd.DataFrame({})

        #將col name進行修改
        for sheet in [YTD_page_1_Total, YTD_page_1_Onshore, YTD_page_1_Offshore]:
            sheet_new_col = []
            for sheet_col in sheet.columns:
                if "onshore" in sheet_col or "offshore" in sheet_col:
                    sheet_col = sheet_col.replace('onshore ','').replace('offshore ','')
                elif "Section" in sheet_col:
                    sheet_col = sheet_col.replace('Section','Department')
                sheet_new_col.append(sheet_col)
            sheet.columns = sheet_new_col
            total_df      = total_df.append(sheet)

        return total_df
        

    前8分頁所需要的month資料    = C_I_前8分頁month()
    三到八分頁所需要的year資料  = three_to_eight_page_YTD()
    New_Open_分頁              = C_I_New_Open()
    top_5_分頁                 = C_I_top5_Fund()
    Promo_and_Focus_分頁       = C_I_Promo_and_Focus_Fund()
    
    return 前8分頁所需要的month資料, 三到八分頁所需要的year資料, New_Open_分頁, top_5_分頁, Promo_and_Focus_分頁
    



if __name__ == '__main__':
    前8分頁所需要的month資料, 三到八分頁所需要的year資料, New_Open_分頁, top_5_分頁, Promo_and_Focus_分頁 = C_I_process()
    output_path = os.getcwd() + '\output\C&I剪貼資料.xlsx'

    excel = Excel_Work()
    wb=excel.write_excel(df=前8分頁所需要的month資料,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9")
    wb.save(output_path)


    excel.append_excel(excel_path=output_path , df=三到八分頁所需要的year資料                 , excel_sheet_name="三到八分頁所需要的year資料"   )
    excel.append_excel(excel_path=output_path , df=New_Open_分頁                             , excel_sheet_name="New_Open_分頁"              )
    excel.append_excel(excel_path=output_path , df=top_5_分頁                                , excel_sheet_name="top_5_分頁"                 )
    excel.append_excel(excel_path=output_path , df=Promo_and_Focus_分頁                      , excel_sheet_name="Promo_and_Focus_分頁"       )