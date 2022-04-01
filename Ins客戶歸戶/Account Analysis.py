#%%

import os
import numpy as np 
import pandas as pd 

df = pd.read_excel(r'D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\2022_Feb_ILP\Feb_Top49全委帳戶.xlsx',sheet_name='Account TOP 5 Holding')


"""
Have to do Two Times 

(1.) offshore

(2.) onshore


"""

#df = df[ df['From'] == 'AIA List (offshore)' ]  # offshore
df = df[ df['From'] == 'AIA List (onshore)' ] # onshore
df = df.dropna()
df

#%%

df.head(50)


# %%

Total_list = []

Top_holding_list_1 = df['top holding-1'].to_list()
Top_holding_list_2 = df['top holding-2'].to_list()
Top_holding_list_3 = df['top holding-3'].to_list()
Top_holding_list_4 = df['top holding-4'].to_list()
Top_holding_list_5 = df['top holding-5'].to_list()



Total_list.extend(Top_holding_list_1)
Total_list.extend(Top_holding_list_2)
Total_list.extend(Top_holding_list_3)
Total_list.extend(Top_holding_list_4)
Total_list.extend(Top_holding_list_5)

Total_list

# %%
from collections import Counter

Analysis_df = pd.DataFrame()

list = Counter(Total_list).keys() # equals to list(set(words))
number = Counter(Total_list).values()

Analysis_df['funds'] = list
Analysis_df['account'] = number


Analysis_df = Analysis_df.sort_values(by='account',ascending=False).reset_index(drop=True).head(50)

Analysis_df
# %%


account_amount_df=pd.DataFrame()

df["Top-1 Amount"] = df['目前規模(新台幣)'] * df['top holding-1 百分比'] * 0.01
df["Top-2 Amount"] = df['目前規模(新台幣)'] * df['top holding-2 百分比'] * 0.01
df["Top-3 Amount"] = df['目前規模(新台幣)'] * df['top holding-3 百分比'] * 0.01
df["Top-4 Amount"] = df['目前規模(新台幣)'] * df['top holding-4 百分比'] * 0.01
df["Top-5 Amount"] = df['目前規模(新台幣)'] * df['top holding-5 百分比'] * 0.01


# account amount
Total_list_amount = []
Top_holding_amount_list_1 = df['Top-1 Amount'].to_list()
Top_holding_amount_list_2 = df['Top-2 Amount'].to_list()
Top_holding_amount_list_3 = df['Top-3 Amount'].to_list()
Top_holding_amount_list_4 = df['Top-4 Amount'].to_list()
Top_holding_amount_list_5 = df['Top-5 Amount'].to_list()

Total_list_amount.extend(Top_holding_amount_list_1)
Total_list_amount.extend(Top_holding_amount_list_2)
Total_list_amount.extend(Top_holding_amount_list_3)
Total_list_amount.extend(Top_holding_amount_list_4)
Total_list_amount.extend(Top_holding_amount_list_5)


# account list 
Total_list = []
Top_holding_list_1 = df['top holding-1'].to_list()
Top_holding_list_2 = df['top holding-2'].to_list()
Top_holding_list_3 = df['top holding-3'].to_list()
Top_holding_list_4 = df['top holding-4'].to_list()
Top_holding_list_5 = df['top holding-5'].to_list()

Total_list.extend(Top_holding_list_1)
Total_list.extend(Top_holding_list_2)
Total_list.extend(Top_holding_list_3)
Total_list.extend(Top_holding_list_4)
Total_list.extend(Top_holding_list_5)

account_amount_df['account'] = Total_list
account_amount_df['account_ammount'] = Total_list_amount

account_amount_df

# %%

output_df = account_amount_df.groupby('account')['account'].count().reset_index(name='Fund numbers')
output_df.index = output_df['account']
output_df = output_df.drop(['account'],axis=1)
output_df = pd.concat([output_df,account_amount_df.groupby('account')['account_ammount'].agg('sum')],axis=1)

output_df = output_df.sort_values(by='account_ammount',ascending=False)
output_df


# %%
output_df.to_excel(r"D:\My Documents\andyhs\桌面\Andy\Ins客戶歸戶\2022_Feb_ILP\Onshore_Account Analysis.xlsx")
# %%
