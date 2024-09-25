import requests
from bs4 import BeautifulSoup
import re
import os
import json
import time
from urllib.parse import urljoin
import random

# 基本設定
base_url = "https://cons.judicial.gov.tw"
main_url = "https://cons.judicial.gov.tw/judcurrentNew1.aspx?fid=38"

# 建立資料夾來存放下載的PDF文件
if not os.path.exists("pdf_downloads"):
    os.makedirs("pdf_downloads")

# 建立判決結果的JSON輸出資料夾
if not os.path.exists("json_outputs"):
    os.makedirs("json_outputs")

# 設置 headers
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}

def retry_request(url, max_retries=5, delay=10):
    for attempt in range(max_retries):
        try:
            print(f"Attempting to load {url} (Attempt {attempt+1}/{max_retries})")
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            print(f"Page {url} loaded successfully.")
            return response
        except requests.RequestException as e:
            print(f"Error occurred while trying to load {url}: {e}, retrying... ({attempt+1}/{max_retries})")
            time.sleep(delay + random.uniform(0, 5))
    print(f"Failed to load {url} after {max_retries} attempts.")
    return None

def download_file(url, filename, max_retries=5, delay=10):
    for attempt in range(max_retries):
        try:
            print(f"Attempting to download {filename} (Attempt {attempt+1}/{max_retries})")
            
            # 檢查文件是否已存在
            if os.path.exists(filename):
                print(f"File {filename} already exists. Skipping download.")
                return True
            
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()
            
            with open(filename, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            print(f"Downloaded: {filename}")
            return True
        except requests.RequestException as e:
            print(f"Error downloading {filename}: {e}, retrying... ({attempt+1}/{max_retries})")
            time.sleep(delay + random.uniform(0, 5))
    print(f"Failed to download {filename} after {max_retries} attempts.")
    return False

def extract_opinion_info(title):
    match = re.search(r"憲法法庭(\d+年憲判字第\d+號)判決([\u4e00-\u9fa5]+)大法官([\u4e00-\u9fa5]+)提出(.*?意見書)(.*)", title)
    if match:
        case_number = match.group(1)
        proposer = match.group(2) + "大法官" + match.group(3)
        opinion_type = match.group(4)
        others = match.group(5)
        
        joiners = []
        if "加入" in others:
            joiner_match = re.findall(r"([\u4e00-\u9fa5]+)大法官([\u4e00-\u9fa5]+)加入", others)
            joiners = [f"{j[0]}大法官{j[1]}" for j in joiner_match]
        
        return {
            "case_number": case_number,
            "proposer": proposer,
            "opinion_type": opinion_type,
            "joiners": joiners
        }
    return None

def process_ruling(url, index):
    try:
        print(f"\n[{index+1}] Processing ruling: {url}")
        
        response = retry_request(url)
        if not response:
            print(f"Failed to process ruling [{index+1}] due to connection issues.")
            return None

        soup = BeautifulSoup(response.content, 'html.parser')

        ruling_info = {
            'title': soup.find('h3', class_='h3_xline').text.strip(),
            'url': url,
            'pdfs': [],
            'opinions': []
        }

        # 處理立場表 PDF 下載
        pdf_links = soup.find_all('a', href=lambda href: href and '/download/download' in href and '立場表' in href.get('title', ''))
        print(f"Found {len(pdf_links)} PDF links for ruling.")
        
        for pdf_index, pdf_link in enumerate(pdf_links):
            pdf_title = pdf_link['title']
            pdf_url = urljoin(base_url, pdf_link['href'])
            pdf_filename = os.path.join("pdf_downloads", pdf_title.replace(" ", "_").replace("「", "").replace("」", "") + ".pdf")
            print(f"Downloading PDF [{pdf_index+1}]: {pdf_title}")
            
            if download_file(pdf_url, pdf_filename):
                ruling_info['pdfs'].append(pdf_title)

        # 處理大法官意見書，只包括部分不同意見書、不同意見書和協同意見書
        opinion_links = soup.find_all('a', href=lambda href: href and '/download/download' in href and ('部分不同意見書' in href.get('title', '') or '不同意見書' in href.get('title', '') or '協同意見書' in href.get('title', '')))
        print(f"Found {len(opinion_links)} relevant opinion links for ruling.")
        
        for opinion_index, opinion_link in enumerate(opinion_links):
            opinion_title = opinion_link['title']
            opinion_info = extract_opinion_info(opinion_title)
            
            if opinion_info:
                print(f"Opinion [{opinion_index+1}]: {opinion_info}")
                ruling_info['opinions'].append(opinion_info)
            else:
                print(f"Failed to extract information from opinion title: {opinion_title}")

        return ruling_info

    except Exception as e:
        print(f"Error processing ruling [{index+1}]: {e}")
        return None

# 主程序
try:
    print(f"Opening main page: {main_url}")
    main_response = retry_request(main_url)
    if not main_response:
        raise Exception("Failed to load main page")

    main_soup = BeautifulSoup(main_response.content, 'html.parser')
    ruling_links = main_soup.find_all('a', href=lambda href: href and '/docdata.aspx' in href and '憲判字' in href.get('title', ''))
    print(f"Found {len(ruling_links)} ruling links.")

    rulings_data = []

    for index, link in enumerate(ruling_links):
        ruling_url = urljoin(base_url, link['href'])
        ruling_info = process_ruling(ruling_url, index)
        if ruling_info:
            rulings_data.append(ruling_info)

        # 每處理5個判決後保存一次數據，以防中途失敗
        if (index + 1) % 5 == 0 or index == len(ruling_links) - 1:
            with open(f"json_outputs/rulings_data_{index + 1}.json", "w", encoding="utf-8") as f:
                json.dump(rulings_data, f, ensure_ascii=False, indent=2)
            print(f"Saved data for {index + 1} rulings.")

        time.sleep(random.uniform(2, 5))  # 在每次請求之間添加隨機延遲，以避免過於頻繁的請求

except Exception as e:
    print(f"An error occurred in the main process: {e}")

print("Script completed.")