from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm  # 引入 tqdm 來顯示進度條
import pandas as pd
import time
from datetime import datetime  # 引入 datetime 來處理日期

# 設定 Selenium 的 Chrome 驅動
chrome_options = Options()
chrome_options.add_argument("--headless")  # 啟用無頭模式

# 指定 Chromedriver 的路徑
service = Service(executable_path='./chromedriver.exe')

# 初始化 webdriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# 用戶名單
usernames = ["formosa_slosi", "tetratio", "y86n2qc9dt", "takotime", "beagle529", "manymao", "yango_chen",
             "cgc", "afriedfish", "liaod", "cloud_strife", "diat", "chocobotang", "zeropower", "konran",
             "zoehuang","hilda1080","pinsi","alan1014","kmwei","victor0117","larryxo","quantina","paula2058",
             "oscar12346","xxcco1111","weiweili","chessraider","rody0803","cpho","phy202008","kumi824","kueichen",
             "ray0715","thomas0422","claire123456789","lonely1009","cgc003"]

# 棋類型和相對應的網址
game_types = {
    "5min": "http://questgames.net/reversi/#user/",
    "1min": "http://questgames.net/reversi1/#user/"
}

# 獲取今天的日期，格式為 YYYYMMDD
today = datetime.now().strftime("%Y%m%d")

# 處理每種棋類型
for game_type, url_prefix in game_types.items():
    data = []
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
                    'Rank': int(rank.split()[0]),
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
    # 將 DataFrame 儲存到 Excel 文件，文件名包含當天日期
    excel_filename = f'user_data-{game_type}-{today}.xlsx'
    df.to_excel(excel_filename, index=False)
    print(f"數據已儲存至 {excel_filename} 並按排名排序。")

# 關閉瀏覽器
driver.quit()
