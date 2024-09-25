import pandas as pd
import json

# 讀取 Excel 檔案
excel_file = '農業部112.xlsx'

# 載入 Excel 檔案的所有工作表
excel_data = pd.read_excel(excel_file, sheet_name=None)

# 儲存所有工作表的資料
all_data = []

# 逐個分頁處理
for sheet_name, data in excel_data.items():
    # 加入「112年」標註欄位
    data['標註'] = '112年'
    
    # 將 NaN 替換成空值
    data = data.fillna("")
    
    # 將資料加入總集合
    all_data.extend(data.to_dict(orient='records'))

# 將結果轉為 JSON 格式
json_data = json.dumps(all_data, ensure_ascii=False, indent=4)

# 印出結果
print(json_data)

# 如果需要將 JSON 儲存到檔案中，可以使用以下程式碼：
with open('output.json', 'w', encoding='utf-8') as f:
     json.dump(all_data, f, ensure_ascii=False, indent=4)
