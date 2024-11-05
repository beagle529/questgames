import pandas as pd
import re
import os
from tqdm import tqdm

# 上傳文件的路徑
file_directory = "./othello/files/"
file_paths = [os.path.join(file_directory, f) for f in os.listdir(file_directory) if f.endswith('.xlsx')]

# 用於存儲所有重複資料的列表
duplicates = []

# 讀取每個文件並檢查重複資料
for file_path in tqdm(file_paths, desc='檢查文件'):
    try:
        df = pd.read_excel(file_path)
        # 提取當前分數 (Rating)
        df['CurrentRating'] = df['Rating'].apply(lambda x: int(re.search(r'^\d+', x).group()))
        # 提取歷史最高分 (MaxRating)
        df['MaxRating'] = df['Rating'].apply(lambda x: int(re.search(r'max:\s*(\d+)', x).group(1)))

        # 只保留 CurrentRating 在 1000 以上的資料
        df = df[df['CurrentRating'] >= 1000]

        # 檢查重複資料，包括 CurrentRating, MaxRating 和 Win/Loss
        duplicated = df[df.duplicated(subset=['CurrentRating', 'MaxRating', 'Win/Loss'], keep=False)]
        if not duplicated.empty:
            duplicates.append((file_path, duplicated))
            # 查找每個重複紀錄，確認前兩次是否一致
            for idx in duplicated.index:
                duplicate_records = df[(df['CurrentRating'] == df.loc[idx, 'CurrentRating']) &
                                       (df['MaxRating'] == df.loc[idx, 'MaxRating']) &
                                       (df['Win/Loss'] == df.loc[idx, 'Win/Loss'])]
                if len(duplicate_records) < 3:
                    # 如果不相符，刪除這些重複記錄中的該筆記錄
                    df = df.drop(idx)

        # 將修改後的 DataFrame 保存至原文件
        df.to_excel(file_path, index=False)
    except Exception as e:
        print(f"無法讀取文件 {file_path}：{str(e)}")

# 輸出重複資料
if duplicates:
    for file_path, duplicated in duplicates:
        print(f"文件 {file_path} 中有重複的資料：")
        print(duplicated.to_string(index=False))
else:
    print("沒有發現重複的資料。")
