import csv
import json
import os

# 1. 设置 CSV 文件夹路径
csv_folder = "docs"   # 替换成你的 CSV 文件夹
json_folder = "json" # JSON 输出文件夹
os.makedirs(json_folder, exist_ok=True)  # 如果文件夹不存在就创建

# 2. 遍历 CSV 文件夹
for filename in os.listdir(csv_folder):
    if filename.endswith(".csv"):
        csv_path = os.path.join(csv_folder, filename)
        json_filename = os.path.splitext(filename)[0] + ".json"
        json_path = os.path.join(json_folder, json_filename)

        # 3. 读取 CSV 并转换成字典列表
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            data = list(reader)

        # 4. 写入 JSON 文件
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
