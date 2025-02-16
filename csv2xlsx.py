import pandas as pd

# 读取 UTF-8 编码的 CSV 文件
csv_file = "rollingstone_best_albums_of_all_time_2023.csv"
df = pd.read_csv(csv_file, encoding="utf-8")

# 将数据写入 XLSX 文件
xlsx_file = "rollingstone_best_albums_of_all_time_2023.xlsx"
df.to_excel(xlsx_file, index=False, engine='openpyxl')