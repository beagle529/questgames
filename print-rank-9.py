import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import re
import os
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
from tqdm import tqdm

# 确保目录存在
if not os.path.exists('./othello/PNG'):
    os.makedirs('./othello/PNG')

# 获取当前的日期和时间，格式为 YYYYMMDDHHMM
current_time = datetime.now().strftime("%Y%m%d%H%M")

# 上传文件的路径
file_directory = "./othello/files/"
file_paths = [os.path.join(file_directory, f) for f in os.listdir(file_directory) if f.endswith('.xlsx')]

# 用于存储所有数据的列表
all_data = {
    "1min": [],
    "5min": []
}

# 设置中文字体
zh_font = fm.FontProperties(fname='./font/SimHei.ttf')  # 请将此路径替换为您实际安装的字体路径

# 黑名单设置
blacklist = ["skyneko1224", "taiwanchan", "formosa_"]

# 定义解析 Win/Loss 栏位的函数
def parse_win_loss(win_loss):
    try:
        wins, losses, rest = win_loss.split('-')
        draws, win_rate = rest.split(' ')
        draws = draws.strip()
        win_rate = win_rate.strip('()')
        return int(wins), int(losses), int(draws), float(win_rate)
    except Exception as e:
        print(f"Error parsing win/loss: {win_loss} - {e}")
        return None, None, None, None

# 读取每个文件并追加到对应的列表中
for file_path in tqdm(file_paths, desc='读取文件'):
    try:
        df = pd.read_excel(file_path)
        # 提取日期信息
        date_str = re.search(r'\d{12}', file_path).group()
        df['Date'] = datetime.strptime(date_str, '%Y%m%d%H%M')
        # 提取当前分数 (Rating)
        df['CurrentRating'] = df['Rating'].apply(lambda x: int(re.search(r'^\d+', x).group()))
        # 提取历史最高分 (MaxRating)
        df['MaxRating'] = df['Rating'].apply(lambda x: int(re.search(r'max:\s*(\d+)', x).group(1)))
        # 解析 Win/Loss 栏位
        df[['Wins', 'Losses', 'Draws', 'Win Rate']] = df['Win/Loss'].apply(lambda x: pd.Series(parse_win_loss(x)))
        # 将胜场数、负场数和和局转换为整数类型
        df['Wins'] = df['Wins'].astype('Int64')
        df['Losses'] = df['Losses'].astype('Int64')
        df['Draws'] = df['Draws'].astype('Int64')
        # 根据文件名中的棋局类型区分数据
        if "1min" in file_path:
            all_data["1min"].append(df)
        elif "5min" in file_path:
            all_data["5min"].append(df)
    except Exception as e:
        print(f"无法读取文件 {file_path}：{str(e)}")

# 设置线条颜色的函数
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

# 分析每个棋局类型
for game_type, data_list in all_data.items():
    if data_list:
        # 将所有数据合并到一个DataFrame中
        combined_df = pd.concat(data_list, ignore_index=True)
        # 按日期和时间排序
        combined_df = combined_df.sort_values(by='Date')
        # 提取最新Rating并按降序排序
        latest_ratings = combined_df.groupby('Username').last().reset_index()
        sorted_users = latest_ratings.sort_values(by='CurrentRating', ascending=False)['Username']

        # 仅获取前18名用户，过滤掉黑名单中的用户
        top_users = [user for user in sorted_users if user not in blacklist][:18]  # 修改为前18名用户

        # 绘制六列3行的折线图
        fig, axes = plt.subplots(3, 6, figsize=(24, 18))  # 修改为6刖3行
        axes = axes.flatten()

        for ax, username in zip(axes, top_users):
            user_data = combined_df[combined_df['Username'] == username].sort_values(by='Date')
            color = get_line_color(user_data['CurrentRating'].iloc[-1])
            max_rating = user_data['MaxRating'].max()

            # 计算每段期间的比赛次数
            user_data['Total Matches'] = user_data['Wins'] + user_data['Losses'] + user_data['Draws']
            user_data['Match Increase'] = user_data['Total Matches'].diff().fillna(0).astype(int)

            # 数据清洗，排除异常值
            user_data = user_data[user_data['Match Increase'] >= 1]
            user_data = user_data[user_data['Match Increase'] < 1000]  # 设置合理的阀值，排除异常值

            ax1 = ax
            ax1.plot(user_data['Date'], user_data['CurrentRating'], marker='o', label=f"{username}", color=color)
            ax1.set_title(f"{username} (MAX: {max_rating})", fontproperties=zh_font)
            ax1.set_xlabel('月-日', fontproperties=zh_font)
            ax1.set_ylabel('最新分数 (Rating)', fontproperties=zh_font)
            ax1.grid(True)
            ax1.legend(loc='upper left')
            ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d', tz=None))
            plt.setp(ax1.get_xticklabels(), rotation=45, fontproperties=zh_font)

            # 在图表下方添加比赛次数的长条图
            ax2 = ax1.twinx()
            ax2.bar(user_data['Date'], user_data['Match Increase'], alpha=0.3, color='gray', width=0.8)
            ax2.set_ylabel('比赛次数', fontproperties=zh_font)

            match_increase_max = user_data['Match Increase'].max()
            if pd.isna(match_increase_max) or match_increase_max == 0:
                match_increase_max = 1  # 避免 y 轴为零或 NaN/Inf
            ax2.set_ylim(0, match_increase_max * 1.1)

        # 如果有效用户数量不足18，则删除多余的轴
        for j in range(len(top_users), 18):
            fig.delaxes(axes[j])

        plt.tight_layout()
        # 保存图表，档名中包含当前日期和时间
        plt.savefig(f'./othello/PNG/{game_type}_top_18_rating_changes_{current_time}.png')  # 修改为前18名
        # plt.show()

        print(f"六刖3行的折线图已保存至 ./othello/PNG/{game_type}_top_18_rating_changes_{current_time}.png")
