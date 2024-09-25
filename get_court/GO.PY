import PyPDF2
import re
import json

def read_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

def parse_content(content):
    lines = content.split('\n')
    result = {}
    current_item = None
    current_category = None
    name_buffer = []

    for line in lines:
        line = line.strip()
        
        # 匹配主文項次
        item_match = re.search(r'第\s*(\S+)\s*項', line)
        if item_match:
            if current_item:
                result[current_item] = process_item(result.get(current_item, {}), name_buffer, current_category)
            current_item = item_match.group(1)
            result[current_item] = {}
            current_category = None
            name_buffer = []
            continue

        # 匹配分類
        if '全部同意：' in line:
            current_category = '全部同意'
            continue
        elif '部分同意' in line:
            current_category = '部分同意'
            continue
        elif '部分不同意' in line:
            current_category = '部分不同意'
            continue
        elif '全部不同意：' in line:
            current_category = '全部不同意'
            continue

        # 處理名單
        names = re.findall(r'(\w+)大法官', line)
        if names:
            name_buffer.extend(names)

        # 處理"無"的情況
        if '無' in line and current_item:
            result[current_item]['不同意'] = []

    # 處理最後一項
    if current_item:
        result[current_item] = process_item(result.get(current_item, {}), name_buffer, current_category)

    return result

def process_item(item_dict, names, category):
    if category == '全部同意':
        item_dict['同意'] = names
    elif category == '全部不同意':
        item_dict['不同意'] = names
    elif category == '部分同意':
        item_dict['部分同意'] = names
    elif category == '部分不同意':
        item_dict['部分不同意'] = names
    else:
        if '同意' not in item_dict:
            item_dict['同意'] = names
        elif '不同意' not in item_dict:
            item_dict['不同意'] = names
    return item_dict

def pdf_to_json(pdf_path):
    content = read_pdf(pdf_path)
    parsed_data = parse_content(content)
    return json.dumps(parsed_data, ensure_ascii=False, indent=2)

# 使用示例
pdf_path = "113年憲判字第8號判決主文立場表.pdf"
json_output = pdf_to_json(pdf_path)
print(json_output)

# 將結果保存到文件
with open("output.json", "w", encoding="utf-8") as f:
    f.write(json_output)
print("JSON output saved to output.json")