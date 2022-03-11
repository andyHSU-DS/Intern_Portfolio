#載入套件
import plotly
import plotly.express as px
import pandas as pd
import plolty.graph_objects as go
import plotly.io as pio
pio.renders.default = 'notebook'
from plotly.subplots import make_subplots
import threading
import sys
import os


#抓資料
def collect_data():
    abspath = r'input/'
    files   = os.listdir(abspath)
    for file in files:
        if 'EC' in file:
            path       = abspath + file
        elif '姓名' in file:
            sales_list = abspath + file
    return path, sales_list


#先找我們需要的業務
def get_sales(sales_file):
    sales_name = pd.read_excel(sales_file)
    sales_need = sales_name['姓名'][sales_name['姓名'].apply( lambda x:len(x)<4 and len(x)>=2)]
    sales_df   = pd.DataFrame(sales_need,columns = ['姓名'])
    return sales_df

def read_csv(File_Path,start,end,encoding = 'utf-8-sig'):
    i = 0
    try:
        with open(File_Path, 'r', encoding = encoding) as file_reader:
            while True:
                line = file_reader.readline().replace('\n','').replace('\t','').replace('晉達環球高收益債券基金 C 收益-3 股份 （南非幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-3 股份 （南非幣避險 IRD 月配)')\
                    .replace('晉達環球高收益債券基金 C 收益-2 股份 （澳幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-2 股份 （澳幣避險 IRD 月配)')\
                    .replace('晉達環球動力股息基金 C 收益-2 股份 （南非幣避險 IRD, 月配)','晉達環球動力股息基金 C 收益-2 股份 （南非幣避險 IRD 月配)').split(',')
                #如果在範圍內就要抓每行
                if start <= i <= end:
                    if len(line) != 0 :
                        Line_list.append(line)   
                        i += 1
                    else:
                        break
                #如果比start小，就pass，記得索引要+1不然會無限輪迴
                elif i < start:
                    i += 1
                    pass
                elif i > end:
                    break
    except:
        with open(File_Path, 'r', encoding = 'big5') as file_reader:
            while True:
                line = file_reader.readline().replace('\n','').replace('\t','').replace('晉達環球高收益債券基金 C 收益-3 股份 （南非幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-3 股份 （南非幣避險 IRD 月配)')\
                    .replace('晉達環球高收益債券基金 C 收益-2 股份 （澳幣避險 IRD, 月配)','晉達環球高收益債券基金 C 收益-2 股份 （澳幣避險 IRD 月配)')\
                    .replace('晉達環球動力股息基金 C 收益-2 股份 （南非幣避險 IRD, 月配)','晉達環球動力股息基金 C 收益-2 股份 （南非幣避險 IRD 月配)').split(',')
                #如果在範圍內就要抓每行
                if start <= i <= end:
                    if len(line) != 0 :
                        Line_list.append(line)   
                        i += 1
                    else:
                        break
                #如果比start小，就pass，記得索引要+1不然會無限輪迴
                elif i < start:
                    i += 1
                    pass
                elif i > end:
                    break
    finally:
        print('編碼:',encoding)
        print(start,end)
    print(i)
    return '輸入資料Done'

#專資料型態及日期，預設period = 'M'
def column_convert(df,column_name='日期',period='M'):
    print('轉換欄位開始')
    #轉datetime形式
    df['日期'] = pd.to_datetime(df['日期'],format = '%Y-%m-%d')
    #轉成月份
    df['日期'] = df['日期'].df.to_period(period)
    for x in df.columns:
        if '金額' in x:
            print(x+'欄位屬性轉換')
            df[x] = df[x].astype('float')
    print('欄位屬性轉換完成')
    return df

#過濾想看的交易類別，目前預設RSP
def trade_type_filter(df,type='RSP'):
    df = df[df['交易類別'] == 'RSP']
    df.reset_index(drop=True,inplace=True)
    return df

#畫每個業務，包含Team
def make_all_agent_plot(df):
    df      = df.groupby(['Agent','日期']).agg({'金額(台幣)':'sum'})
    picture = px.line(x=df.index.get_level_values(1).to_timestamp(),y=df['金額(台幣)'],color=df.index.get_level_values(0),\
                      labels = {'x':'日期','y':'扣款金額(NTD)','color':'業務'},title='RSP扣款金額')
    print('繪圖完成(all agents)')
    picture.write_html(file = r'output/全體業務RSP表現.html')


#畫需要看業務，同時做一個csv放部分業務的資料
def make_agent_plot(df,sales_df):
    df               = df.groupby(['Agent','日期']).agg({'金額(台幣)':'sum'}).reset_index()
    #做過濾
    df_pic           = df[df['Agent'].isin(sales_df['姓名'].values)]
    df_pic_re_index  = df_pic.set_index('日期')
    picture = px.line(x=df_pic_re_index.index.to_timestamp(),y=df_pic_re_index['金額(台幣)'],color=df_pic_re_index['Agent'],\
                      labels = {'x':'日期','y':'扣款金額(NTD)','color':'業務'},title='RSP扣款金額')
    print('繪圖完成(part agents)')
    picture.write_html(file = r'output/部分業務RSP表現.html')
    #-------------------做csv------------------------#
    df_csv = df_pic.groupby(['Agent','日期']).agg({'金額(台幣)':'sum'}).unstack()
    df_csv.to_csv(r'output/部分業務個月資料.csv',index=True,header=True,encoding='utf-8-sig')




if __name__ == '__main__':
    path, sales_list = collect_data()
    #-------------------read file----------------#
    Line_list = []
    part1     = threading.Thread(target = read_csv, args = (path,0,10),)
    part2     = threading.Thread(target = read_csv, args = (path,11,8190),)