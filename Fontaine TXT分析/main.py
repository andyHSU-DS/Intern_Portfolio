##載入套件
import os 
import pandas as pd
import re

def TXT_process():

    def data_collection():
        current_path = os.getcwd()
        #我需要的資料夾
        folders = []
        for folder in os.listdir(current_path):
            if '.' not in folder:
                folders.append(folder)

        for folder in folders:
            print('-'*50)
            print(folder)
            #設定特定資料夾
            #判斷是不是整年的資料

            #2021年的資料
            #下面會有日報的資料，分成關聯非關聯
            if '2020' not in folder and '年' in folder:
                #particular_folder 是d:\My Documents\andyhs\桌面\Andy\Fontaine TXT分析\2020年交易監控月報
                particular_folder = os.getcwd() + '\\' + folder
                #找partiuclar_folder底下的資料夾
                months_folders = os.listdir(particular_folder)
                for month_folder in months_folders:
                    #進到每個月的資料夾
                    particular_path = particular_folder + '\\' + month_folder
                    files = os.listdir(particular_path)
                    for file in files:
                        print(file)
                        #找資料夾下面的.csv or .xlsx檔案
                        if re.match(r'TXT_[M|D]',file):
                            print(particular_path + '\\' + file)

                        elif '日報' in file and '.' not in file:
                            日報_folder = particular_path + '\\' + file
                            for 日報_file in os.listdir(日報_folder):
                                if re.match(r'TXT_[M|D]',日報_file):
                                    print(日報_folder + '\\' + 日報_file)
                        else:
                            pass
                    print('-'*50+month_folder+'-'*50)
            #這邊就是今年每月的了
            #下面會有日報的資料，分成關聯非關聯
            #跟月報
            #和2021得差別在於資料層數
            elif '2022' in folder:
                #particular_folder 是d:\My Documents\andyhs\桌面\Andy\Fontaine TXT分析\202201
                particular_folder = os.getcwd() + '\\' + folder
                #找partiuclar_folder底下的資料夾
                months_folders = os.listdir(particular_folder)
                for month_folder in months_folders:
                    if re.match(r'TXT_[M|D]',month_folder):
                        print(particular_path + '\\' + file)
                    elif '日報' in month_folder and '.' not in file:
                        日報_folder = particular_folder + '\\' + month_folder
                        for 日報_file in os.listdir(日報_folder):
                            if re.match(r'TXT_[M|D]',日報_file):
                                print(日報_folder + '\\' + 日報_file)
                    else:
                        pass

            elif '2020' in folder:
                #12月以前是舊資料
                #particular_folder 是d:\My Documents\andyhs\桌面\Andy\Fontaine TXT分析\2020年交易監控月報
                particular_folder = os.getcwd() + '\\' + folder
                #找partiuclar_folder底下的資料夾
                months_folders = os.listdir(particular_folder)
                for month_folder in months_folders:
                    if '12' not in month_folder:
                        #進到每個月的資料夾
                        particular_path = particular_folder + '\\' + month_folder
                        files = os.listdir(particular_path)
                        for file in files:
                            #找資料夾下面的.csv or .xlsx檔案
                            if re.match(r'TXT_[M|D]',file):
                                print(particular_path + '\\' + file)

                            elif 'DB' in file and '.' not in file:
                                日報_folder = particular_path + '\\' + file
                                for 日報_file in os.listdir(日報_folder):
                                    if re.match(r'\d+',日報_file):
                                        print(日報_folder + '\\' + 日報_file)
                            else:
                                pass
                    else: 
                        #進到每個月的資料夾
                        particular_path = particular_folder + '\\' + month_folder
                        files = os.listdir(particular_path)
                        for file in files:
                            print(file)
                            #找資料夾下面的.csv or .xlsx檔案
                            if re.match(r'\d+-\d+ #',file):
                                print(particular_path + '\\' + file)

                            elif 'DB' in file and '.' not in file:
                                DB_folder = particular_path + '\\' + file
                                for DB_file in os.listdir(DB_folder):
                                    if re.match(r'\d+-#\d+',DB_file) and '.xlsx' in DB_file:
                                        print(DB_folder + '\\' + DB_file)
                            elif '.' not in file:
                                next_particular_path = particular_path + "\\" + file
                                next_particular_files = os.listdir(next_particular_path)
                                for file in next_particular_files:
                                    if '.xlsx' in file:
                                        print(next_particular_path + file)
                                
                    print('-'*50+month_folder+'-'*50)