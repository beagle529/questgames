import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import re
import os
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
from tqdm import tqdm

# 確保目錄存在
if not os.path.exists('./othello/PNG'):
    os.makedirs('./othello/PNG')

# 上傳文件的路徑
file_directory = "./othello/files/"
file_paths = [os.path.join(file_directory, f) for f in os.listdir(file_directory) if f.endswith('.xlsx')]

# 用於存儲所有數據的列表
all_data = {
    "1min": [],
    "5min": []
}

# 設置中文字體
zh_font = fm.FontProperties(fname='./font/SimHei.ttf')  # 請將此路徑替換為您實際安裝的字體路徑

# 讀取每個文件並追加到對應的列表中
for file_path in tqdm(file_paths, desc='讀取文件'):
    try:
        df = pd.read_excel(file_path)
        # 提取日期信息
        date_str = re.search(r'\d{12}', file_path).group()
        df['Date'] = datetime.strptime(date_str, '%Y%m%d%H%M')
        # 提取當前分數 (Rating)
        df['CurrentRating'] = df['Rating'].apply(lambda x: int(re.search(r'^\d+', x).group()))
        # 提取歷史最高分 (MaxRating)
        df['MaxRating'] = df['Rating'].apply(lambda x: int(re.search(r'max:\s*(\d+)', x).group(1)))
        # 根據文件名中的棋局類型區分數據
        if "1min" in file_path:
            all_data["1min"].append(df)
        elif "5min" in file_path:
            all_data["5min"].append(df)
    except Exception as e:
        print(f"無法讀取文件 {file_path}：{str(e)}")

# 設置線條顏色的函數
def get_line_color(rating):
    if rating >= 2000:
        return (238/255, 0, 0)  # RGB(238,0,0)
    elif 1700 <= rating <= 1999:
        return (204/255, 170/255, 0)  # RGB(204,170,0)
    elif 1500 <= rating <= 1699:
        return (52/255, 52/255, 241/255)  # RGB(52,52,241)
    elif 1200 <= rating <= 1499:
        return (0, 169/255, 0)  # RGB(0,169,0)
    else:
        return (153/255, 153/255, 153/255)  # RGB(153,153,153)

# 分析每個棋局類型
for game_type, data_list in all_data.items():
    if data_list:
        # 將所有數據合併到一個DataFrame中
        combined_df = pd.concat(data_list, ignore_index=True)
        # 按日期和時間排序
        combined_df = combined_df.sort_values(by='Date')
        # 提取最新Rating並按降序排序
        latest_ratings = combined_df.groupby('Username').last().reset_index()
        sorted_users = latest_ratings.sort_values(by='CurrentRating', ascending=False)['Username']

        # 繪製十八宮格折線圖
        with tqdm(total=(len(sorted_users) + 17) // 18, desc=f'繪製 {game_type} 十八宮格折線圖') as pbar:
            for i in range(0, len(sorted_users), 18):
                top_users = sorted_users[i:i+18]
                fig, axes = plt.subplots(3, 6, figsize=(30, 15))
                axes = axes.flatten()

                for j, (ax, username) in enumerate(zip(axes, top_users), start=i+1):
                    user_data = combined_df[combined_df['Username'] == username].sort_values(by='Date')
                    color = get_line_color(user_data['CurrentRating'].iloc[-1])
                    max_rating = user_data['MaxRating'].max()
                    ax.plot(user_data['Date'], user_data['CurrentRating'], marker='o', label=f"{username}", color=color)
                    ax.set_title(f"{username} (MAX: {max_rating})", fontproperties=zh_font)
                    ax.set_xlabel('月-日', fontproperties=zh_font)
                    ax.set_ylabel('最新分數 (Rating)', fontproperties=zh_font)
                    ax.grid(True)
                    ax.legend()
                    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
                    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d', tz=None))
                    plt.setp(ax.get_xticklabels(), rotation=45, fontproperties=zh_font)

                for k in range(len(top_users), 18):
                    fig.delaxes(axes[k])

                plt.tight_layout()
                # 保存圖表，檔名中包含當前日期和時間
                current_time = datetime.now().strftime("%Y%m%d%H%M")
                plt.savefig(f'./othello/PNG/{game_type}_rating_changes_{i+1}_{i+18}_{current_time}.png')
                # plt.show()

                print(f"十八宮格折線圖已保存至 ./othello/PNG/{game_type}_rating_changes_{i+1}_{i+18}_{current_time}.png")
                pbar.update(1)
