# 抓取questgames黑白棋(一分)的紀錄
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm  # Import tqdm for progress bar
import pandas as pd
import time

# Selenium's Chrome driver setup
chrome_options = Options()
chrome_options.add_argument("--headless")  # for headless mode

# Specify the path to your Chromedriver
service = Service(executable_path='./chromedriver.exe')

# Initialize the webdriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# List of usernames 會員名單 ，ID前後請用雙引號並使用逗點區隔 ，英文請用小寫
usernames = ["formosa_slosi", "tetratio", "y86n2qc9dt", "takotime", "beagle529", "manymao", "yango_chen",
             "cgc", "afriedfish", "liaod", "cloud_strife", "diat", "chocobotang", "zeropower", "konran",
             "zoehuang","hilda1080","pinsi","alan1014","kmwei","victor0117","larryxo","quantina","paula2058",
             "oscar12346","xxcco1111","weiweili","chessraider","rody0803","cpho","phy202008","kumi824","kueichen",
             "ray0715","thomas0422","claire123456789","lonely1009","cgc003"]

# Store data in a list
data = []

# Loop through each username with a progress bar
for username in tqdm(usernames, desc="Processing Users"):
    driver.get(f"http://questgames.net/reversi1/#user/{username}")
    time.sleep(1)  # wait for dynamic content to load 在解析HTML之前等待1秒鐘
    try:
        user_element = driver.find_element(By.CSS_SELECTOR, 'li.record')
        rank = user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Rank"]]/td').text.strip()
        data.append({
            'Username': username,
            'Rating': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Rating"]]/td').text,
            'Rank': int(rank.split()[0]),  # Assumes rank is the first part of the text and is a number
            'Win/Loss': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Win loss"]]/td').text,
            'Streak': user_element.find_element(By.XPATH, './table/tbody/tr[th[text()="Streak"]]/td').text
        })
    except Exception as e:
        print(f"Failed to retrieve data for {username}: {str(e)}")

# Close the browser
driver.quit()

# Convert data into a DataFrame
df = pd.DataFrame(data)

# Sort the DataFrame by 'Rank' in ascending order
df = df.sort_values(by='Rank')

# Save the DataFrame to an Excel file
df.to_excel('user_data-1min.xlsx', index=False)
print("Data has been saved to user_data.xlsx and sorted by Rank.")