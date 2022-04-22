#載入套件
import plotly
import plotly.express as px
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
pio.renderers.default = 'notebook'
from plotly.subplots import make_subplots
import threading
import sys
import os

#抓資料
def collect_data():
    abspath = 'D:/My Documents/andyhs/桌面/Andy/契約資料/Input/'
    files = os.listdir(r'Input')
    for file in files:
        if 'EC' in file:
            path = abspath + file
        elif '姓名' in file:
            sales_list = abspath + file
    return path, sales_list


#因為讀檔案有些問題所以用平行運算的方式開啟，如果之後遇到編碼的問題就從這邊修
def read_csv(File_Path,start,end,encoding='utf-8-sig'):
    i = 0
    try:
        with open(File_Path,'r',encoding=encoding) as file_reader:
            while True:
                #print('目前:',i)
                line = file_reader.readline().replace("\n",'').replace("\t",'').replace('晉達環球高收益債券基金 C 收益-3 股份 (南非幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-3 股份 (南非幣避險 IRD 月配)')\
                    .replace('晉達環球高收益債券基金 C 收益-2 股份 (澳幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-2 股份 (澳幣避險 IRD 月配)')\
                    .replace('晉達環球動力股息基金 C 收益-2 股份 (南非幣避險 IRD, 月配)','晉達環球動力股息基金 C 收益-2 股份 (南非幣避險 IRD 月配)').split(',')
                #如果在範圍內，就要抓每行
                if start<=i<=end:
                    if len(line) != 0:
                        Line_list.append(line)
                        i+=1
                    else:
                        break
                #如果比start小，就pass，但記得要讓索引+1不然會無限輪迴
                elif i<start:
                    i+=1
                    pass
                #超過end就break
                elif i>end:
                    break
    except:
        with open(File_Path,'r',encoding='big5') as file_reader:
            while True:
                #print('目前:',i)
                line = file_reader.readline().replace("\n",'').replace("\t",'').replace('晉達環球高收益債券基金 C 收益-3 股份 (南非幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-3 股份 (南非幣避險 IRD 月配)')\
                    .replace('晉達環球高收益債券基金 C 收益-2 股份 (澳幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-2 股份 (澳幣避險 IRD 月配)')\
                    .replace('晉達環球動力股息基金 C 收益-2 股份 (南非幣避險 IRD, 月配)','晉達環球動力股息基金 C 收益-2 股份 (南非幣避險 IRD 月配)').split(',')
                #如果在範圍內，就要抓每行
                if start<=i<=end:
                    if len(line) != 0:
                        Line_list.append(line)
                        i+=1
                    else:
                        break
                #如果比start小，就pass，但記得要讓索引+1不然會無限輪迴
                elif i<start:
                    i+=1
                    pass
                #超過end就break
                elif i>end:
                    break
    finally:
        print('編碼:',encoding)
        print(start,end)
    print(i)
    return '載入資料Done'


if __name__ == '__main__':
    path, sales_list = collect_data()
    #--------------------------read file--------------------------------#
    Line_list = []
    part1     = threading.Thread(target = read_csv,args=(path,0,10,),)
    part2     = threading.Thread(target = read_csv,args=(path,11,8190,),)
    part3     = threading.Thread(target = read_csv,args=(path,8191,100000),)
    part4     = threading.Thread(target = read_csv,args=(path,100001,150000),)
    part5     = threading.Thread(target = read_csv,args=(path,150001,200000),)
    part6     = threading.Thread(target = read_csv,args=(path,200001,300000),)
    part7     = threading.Thread(target = read_csv,args=(path,300001,400000),)
    part8     = threading.Thread(target = read_csv,args=(path,400001,600000),)
    part9     = threading.Thread(target = read_csv,args=(path,600001,800000),)
    part10     = threading.Thread(target = read_csv,args=(path,800001,1000000),)
    part11     = threading.Thread(target = read_csv,args=(path,1000001,1200000),)
    part12     = threading.Thread(target = read_csv,args=(path,1200001,1400000),)
    part13     = threading.Thread(target = read_csv,args=(path,1400001,1500000),)
    part14     = threading.Thread(target = read_csv,args=(path,1500001,1600000),)
    part15     = threading.Thread(target = read_csv,args=(path,1600001,1900000),)
    part16     = threading.Thread(target = read_csv,args=(path,1900001,2000000,),)
    part17     = threading.Thread(target = read_csv,args=(path,2000001,2100000,),)
    part18     = threading.Thread(target = read_csv,args=(path,2100001,2200000),)
    part19     = threading.Thread(target = read_csv,args=(path,2200001,2300000),)
    part20    = threading.Thread(target = read_csv,args=(path,2300001,2400000),)
    part21     = threading.Thread(target = read_csv,args=(path,2400001,2500000),)
    part22     = threading.Thread(target = read_csv,args=(path,2500001,2600000),)
    part23    = threading.Thread(target = read_csv,args=(path,2600001,2700000),)
    part24     = threading.Thread(target = read_csv,args=(path,2700001,2800000),)
    part25     = threading.Thread(target = read_csv,args=(path,2800001,2900000),)
    part26     = threading.Thread(target = read_csv,args=(path,2900001,3000000),)


    part1.start()
    part2.start()
    part3.start()
    part4.start()
    part5.start()
    part6.start()
    part7.start()
    part8.start()
    part9.start()
    part10.start()
    part11.start()
    part12.start()
    part13.start()
    part14.start()
    part15.start()
    part16.start()
    part17.start()
    part18.start()
    part19.start()
    part20.start()
    part21.start()
    part22.start()
    part23.start()
    part24.start()
    part25.start()
    part26.start()
    


    part1.join()
    part2.join()
    part3.join()
    part4.join()
    part5.join()
    part6.join()
    part7.join()
    part8.join()
    part9.join()
    part10.join()
    part11.join()
    part12.join()
    part13.join()
    part14.join()
    part15.join()
    part16.join()
    part17.join()
    part18.join()
    part19.join()
    part20.join()
    part21.join()
    part22.join()
    part23.join()
    part24.join()
    part25.join()
    part26.join()

    #處理一些欄位的問題
    read_df = pd.DataFrame(Line_list)
    read_df = read_df[read_df[0]!='']
    read_df.columns = ['交易類別','資料別','日期','扣款狀態','戶號','基金代碼','基金簡稱','申購幣別','金額(非台幣)','金額(台幣)',\
        '是否為台股基金','[股/債/貨幣/平衡]','Main_Code','EC_Code','Agent','[契約書號(For ROBO)]']
    #因為用平行運算，所以不會知道第一行(header)在哪裡，所以手動輸入column name並將原有header delete
    read_df = read_df[read_df['扣款狀態']!='扣款狀態']

    read_df['日期'] = pd.to_datetime(read_df['日期'])
    read_df['日期_M'] = read_df['日期'].dt.to_period('M')

    Month_data_dict = {}
    read_data_M = read_df.groupby(['日期_M'])
    for M, data in read_data_M:
        Month_data_dict[M] = data

    for month, data in Month_data_dict.items():
        print(month)
        print(len(data))
        data_path = r'每月資料' + '/' + str(month) + '.csv'
        data.to_csv(data_path,encoding='UTF-8-SIG',header=True,index=False)