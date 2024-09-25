import os
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Step 1: 抓取該頁面上的所有 ODS 檔案的鏈接
url = "https://www.mohw.gov.tw/cp-5297-75662-1.html"
base_url = "https://www.mohw.gov.tw"
response = requests.get(url)
soup = BeautifulSoup(response.content, "html.parser")

# 查找所有含有ODS文件鏈接的a標籤
ods_links = soup.find_all("a", href=True, title=True)

# 下載ODS文件存放目錄
download_folder = "ods1_files"
if not os.path.exists(download_folder):
    os.makedirs(download_folder)