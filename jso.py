import pandas as pd
import json

# 讀取 Excel 文件中的所有分頁
xls = pd.ExcelFile('cleaned_merged_ods.xlsx')

# 創建一個列表來儲存所有分頁的數據
merged_data = []

# 設置年份為 112 年
year = 112

# 金額欄位的名稱，這個根據你的實際欄位名稱來設置
amount_column_name = '金額'

# 遍歷所有分頁
for sheet_name in xls.sheet_names:
    # 讀取每個分頁的資料
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # 添加一列分頁名稱以標識數據來源
    df['SheetName'] = sheet_name
    
    # 添加一列年份欄位，值設為 112 年
    df['年份'] = year
    
    # 將該分頁的數據追加到列表中
    merged_data.append(df)

# 將所有分頁的數據合併成一個 DataFrame
merged_df = pd.concat(merged_data, ignore_index=True)

# 遍歷所有欄位進行處理
for col in merged_df.columns:
    # 處理金額欄位：轉換為數字格式，將無效數據填充為 0
    if col == amount_column_name:
        merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce').fillna(0)
    
    # 處理日期類型欄位：轉換為字串格式
    elif pd.api.types.is_datetime64_any_dtype(merged_df[col]):
        merged_df[col] = merged_df[col].dt.strftime('%Y-%m-%d')
    
    # 其他欄位：轉換為字串格式
    else:
        merged_df[col] = merged_df[col].astype(str)

# 將合併後的 DataFrame 轉換成字典形式
merged_dict = merged_df.to_dict(orient='records')

# 將字典轉換成 JSON 字符串
json_data = json.dumps(merged_dict, ensure_ascii=False, indent=4)

# 將 JSON 數據保存到文件
with open('merged_data.json', 'w', encoding='utf-8') as json_file:
    json_file.write(json_data)

print("所有分頁已合併、金額轉為數字，其他欄位轉為字串並保存為 'merged_data.json'.")
