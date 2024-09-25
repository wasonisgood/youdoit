import pandas as pd
import pdfplumber
import logging

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def make_unique_columns(columns):
    seen = {}
    unique_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            unique_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique_columns.append(col)
    return unique_columns

def standardize_columns(columns):
    mapping = {
        '機關名稱': '機關名稱',
        '機關名稱_1': '機關名稱',
        '宣導項目': '宣導項目',
        '宣導項目_1': '宣導項目',
        '媒體類型': '媒體類型',
        '宣導期程': '宣導期程',
    }
    standardized = [mapping.get(col, col) for col in columns]
    return standardized

def adjust_columns(df):
    expected_columns = ['機關名稱', '宣導項目', '媒體類型', '宣導期程']
    existing_columns = df.columns.tolist()
    
    for col in expected_columns:
        if col not in existing_columns:
            df[col] = None
    
    df = df[expected_columns]
    return df

def extract_tables_with_pdfplumber(pdf_path):
    all_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            logging.info(f"處理第 {page_num} 頁")
            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 3,
                "edge_min_length": 3,
                "min_words_vertical": 1,
                "min_words_horizontal": 1
            })
            
            logging.info(f"檢測到 {len(tables)} 個表格")
            
            for table_num, table in enumerate(tables, start=1):
                if len(table) < 2:
                    logging.warning(f"第 {page_num} 頁的第 {table_num} 個表格行數不足，已跳過。")
                    continue
                
                headers = table[0]
                data_rows = table[1:]
                
                unique_headers = make_unique_columns(headers)
                standardized_headers = standardize_columns(unique_headers)
                
                df = pd.DataFrame(data_rows, columns=standardized_headers)
                all_data.append(df)
                
                logging.info(f"提取第 {page_num} 頁的第 {table_num} 個表格，行數：{len(df)}")
    
    if all_data:
        try:
            combined_df = pd.concat(all_data, ignore_index=True, sort=False)
            logging.info("成功合併所有表格。")
            return combined_df
        except pd.errors.InvalidIndexError as e:
            logging.error("合併 DataFrame 時發生錯誤:", exc_info=True)
            return pd.DataFrame()
    else:
        logging.warning("未檢測到任何表格。")
        return pd.DataFrame()

def clean_data(df):
    df = df.ffill()  # 使用 ffill() 代替 fillna(method='ffill')
    
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    
    df = df.drop_duplicates()
    
    df = df.apply(lambda x: x.map(lambda y: y.strip() if isinstance(y, str) else y))  # 使用 apply 替代 applymap
    
    df = adjust_columns(df)
    
    if '宣導期程' in df.columns:
        df['宣導期程'] = pd.to_datetime(df['宣導期程'], format='%Y-%m-%d', errors='coerce')  # 指定日期格式
    
    return df

def main(pdf_file_path):
    logging.info(f"開始處理 PDF 文件：{pdf_file_path}")
    
    combined_df = extract_tables_with_pdfplumber(pdf_file_path)
    
    if combined_df.empty:
        logging.warning("未檢測到表格或合併後的 DataFrame 為空。")
        return
    
    cleaned_df = clean_data(combined_df)
    
    output_file = 'output_pdfplumber.xlsx'
    cleaned_df.to_excel(output_file, index=False)
    logging.info(f"數據已保存到 {output_file}")
    
    print(f"數據已保存到 {output_file}")

if __name__ == "__main__":
    pdf_file_path = "1.pdf"
    main(pdf_file_path)
