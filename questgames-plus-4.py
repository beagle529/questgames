import schedule
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm  # 引入 tqdm 來顯示進度條
import pandas as pd
from datetime import datetime  # 引入 datetime 來處理日期
import os  # 引入 os 來處理文件路徑

def fetch_othello_data():
    # 設定 Selenium 的 Chrome 驅動
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # 啟用無頭模式
    chrome_options.add_argument("--disable-gpu")  # 禁用GPU硬件加速
    chrome_options.add_argument("--log-level=3")  # 只記錄嚴重錯誤信息
    chrome_options.add_argument("--disable-dev-shm-usage")  # 避免大量內存佔用
    chrome_options.add_argument("--no-sandbox")  # 解決DevToolsActivePort文件不存在的錯誤

    # 指定 Chromedriver 的路徑
    service = Service(executable_path='./chromedriver.exe')

    # 初始化 webdriver
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # 從 user.txt 文件中讀取用戶名單，轉為小寫並移除空白
    with open('./othello/user.txt', 'r') as file:
        usernames = [line.strip().replace(' ', '').lower() for line in file.readlines()]

    # 棋類型和相對應的網址
    game_types = {
        "5min": "http://questgames.net/reversi/#user/",
        "1min": "http://questgames.net/reversi1/#user/"
    }

    # 獲取當前的日期和時間，格式為 YYYYMMDDHHMM
    current_time = datetime.now().strftime("%Y%m%d%H%M")

    # 確保當前目錄的子目錄 files 存在
    if not os.path.exists('./othello/files'):
        os.makedirs('./othello/files')

    # 處理每種棋類型
    for game_type, url_prefix in game_types.items():
        data = []  # 初始化資料存儲列表
        # 在進度條下處理每個用戶
        for username in tqdm(usernames, desc=f"處理 {game_type} 用戶中"):
            attempts = 0
            success = False
            while attempts < 2 and not success:
                driver.get(f"{url_prefix}{username}")
                time.sleep(0.8)  # 等待內容加載
                try:
                    user_element = driver.find_element(By.CSS_SELECTOR, 'li.record')
                    rank = user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Rank"]]/td').text.strip()
                    data.append({
                        'Username': username,
                        'Rating': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Rating"]]/td').text,
                        'Rank': int(rank.split()[0]),  # 假設排名是文本的第一部分且為數字
                        'Win/Loss': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Win loss"]]/td').text,
                        'Streak': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Streak"]]/td').text
                    })
                    success = True
                except Exception as e:
                    print(f"第 {attempts+1} 次無法取得 {username} 的資料: {str(e)}")
                    attempts += 1

        # 將數據轉換成 DataFrame
        df = pd.DataFrame(data)
        # 按 'Rank' 升序排列 DataFrame
        df = df.sort_values(by='Rank')
        # 將 DataFrame 儲存到 Excel 文件，文件名包含當前日期和時間
        excel_filename = f'./othello/files/Reversi_{game_type}_data_{current_time}.xlsx'
        df.to_excel(excel_filename, index=False)
        print(f"數據已儲存至 {excel_filename} 並按排名排序。")

    # 關閉瀏覽器
    driver.quit()

# 初始執行
fetch_othello_data()

# 設定每六小時執行一次
schedule.every(6).hours.do(fetch_othello_data)

# 保持腳本運行
while True:
    schedule.run_pending()
    time.sleep(1)
