import html
import json
import re
import requests
from bs4 import BeautifulSoup
import csv
import pandas as pd
import os

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
}

link = "https://www.rollingstone.com/music/music-lists/500-greatest-albums-of-all-time-156826/"
filename = "rollingstone_best_albums_of_all_time_2003"

data = []

while link:
    print(f"Crawling {link} ...")

    response = requests.get(link, headers=headers)
    soup = BeautifulSoup(response.content, "html.parser")

    script = soup.select_one("#pmc-lists-front-js-extra")
    script_text = (
        script.get_text(strip=True)
        .replace("&amp;", "&")
        .replace("&#8216;", "‘")
        .replace("&#8217;", "’")
        .replace("&#8230;", "…")
        .replace("\u2008", " ")
    )
    pmcGalleryExports = json.loads(
        re.search(r"var pmcGalleryExports = (.*);", script_text).group(1)
    )

    for item in pmcGalleryExports["gallery"]:
        cover = item["image"].split("?")[0]
        rank = item["positionDisplay"]
        title = re.sub(r"\ufeff|<.*?>", "", item["title"])
        try:
            artist, album = title.split(", ", 1)
        except ValueError:
            try:
                artist, album = title.split(" ,", 1)
            except ValueError:
                artist, album = title.split("’ ", 1)
        album = album.strip("‘").strip("’")
        caption = item["caption"]

        try:
            parts = re.findall(r"(<p.*?>.*?</p>)", item["description"], re.DOTALL)

            # 检查 part[0] 是否为公司和年份信息
            match = re.search(r"<p><em>(.*?)<\/em></p>", parts[0])
            if match:
                company, year = match.group(1).split(", ", 1)
                description = html.unescape(
                    re.sub(
                        r"<.*?>",
                        "",
                        "\n".join(parts[1:]).strip(),
                    )
                )
            else:
                description = html.unescape(
                    re.sub(
                        r"<.*?>",
                        "",
                        "\n".join(parts).strip(),
                    )
                )
        except AttributeError:
            print(f"Error: {item['description']}")

        data.append([rank, cover, artist, album, caption, company, year, description])

    link = pmcGalleryExports.get("nextPageLink")

# 写入CSV文件
with open(filename + ".csv", mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(
        [
            "rank",
            "cover",
            "artist",
            "album",
            "caption",
            "company",
            "year",
            "description",
        ]
    )
    writer.writerows(data)

# 读取 UTF-8 编码的 CSV 文件
csv_file = filename + ".csv"
df = pd.read_csv(csv_file, encoding="utf-8")

# 将数据写入 XLSX 文件
xlsx_file = filename + ".xlsx"
df.to_excel(xlsx_file, index=False, engine="openpyxl")

cwd = os.getcwd()
print(f"Data saved to {cwd}\\{csv_file} and {cwd}\\{xlsx_file}.")
