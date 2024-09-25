import pandas as pd
import re

# 讀取所有分頁的試算表
xls = pd.ExcelFile('merged_ods.xlsx')

# 創建一個字典來儲存所有清洗後的 DataFrame
cleaned_sheets = {}

# 遍歷所有分頁名稱
for sheet_name in xls.sheet_names:
    # 讀取每個分頁的資料
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # 1. 遍歷資料，找到包含 "X月" 的列
    for index, row in df.iterrows():
        # 檢查 'Original_Sheet' 列中的內容是否包含 "X月" 或 "X月" 後面還有其他文字
        if re.match(r'\d+月', str(row['Original_Sheet'])):
            # 2. 檢查該月份倒數第3列第一欄的內容是否為 "填表說明："
            if index >= 2:  # 確保不會索引超出範圍
                if df.iloc[index - 2, 0] == '填表說明：':
                    # 3. 刪除該月份倒數前3列及後面的所有內容
                    df = df.drop(df.index[index - 2:])
                    print(f"Removed rows starting from the 3rd to last row for sheet {sheet_name} at index {index}")
                    break  # 假設只需要清洗一次

    # 將清洗後的 DataFrame 儲存到字典中
    cleaned_sheets[sheet_name] = df

# 將所有清洗過的分頁合併到一個 Excel 檔案中
with pd.ExcelWriter('cleaned_merged_ods.xlsx', engine='xlsxwriter') as writer:
    for sheet_name, cleaned_df in cleaned_sheets.items():
        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("All sheets have been processed, cleaned, and merged into 'cleaned_merged_ods.xlsx'.")
