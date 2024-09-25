import pandas as pd

# 讀取 Excel 文件中的所有分頁
xls = pd.ExcelFile('cleaned_merged_ods.xlsx')

# 創建一個字典來儲存每個分頁的欄位名稱
sheet_columns = {}

# 遍歷所有分頁
for sheet_name in xls.sheet_names:
    # 讀取每個分頁的前幾行（假設欄位名稱在第一列）
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # 儲存每個分頁的欄位名稱（第一列）
    sheet_columns[sheet_name] = set(df.columns)

# 比較所有分頁的共同欄位
common_columns = set.intersection(*sheet_columns.values())

# 找出每個分頁的獨有欄位
unique_columns = {sheet_name: columns - common_columns for sheet_name, columns in sheet_columns.items()}

# 輸出結果
print("共同的欄位名稱：")
print(common_columns)

print("\n各分頁獨有的欄位名稱：")
for sheet_name, columns in unique_columns.items():
    print(f"{sheet_name} 的獨有欄位名稱: {columns}")
