import pandas as pd

def TW_stock_process(file,target = 5):
    TW_stock        = pd.read_excel(file,skiprows=4)
    TW_stock_group  = TW_stock.groupby('名稱')
    top_5_component = pd.DataFrame({})
    #每個基金，做投資比率(%)找到前五
    for fund_name, fund_data in TW_stock_group:
        #print(fund_name)
        data = fund_data.sort_values(by = '投資比率(%)').nlargest(5,'投資比率(%)')
        top_5_component = top_5_component.append(data)
    #將包含Unnamed的欄位去掉
    top_5_component = top_5_component.drop(columns = [ col for col in top_5_component.columns if 'Unnamed' in col])
    top_5_component.to_csv(r'台股持股明細_2022_top5.csv',encoding='utf-8-sig',index=False,header=True)
    return None

if __name__ == "__main__":
    TW_stock_process(r'台股持股明細_2022.xlsx')
