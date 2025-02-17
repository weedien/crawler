import os
import re
import requests
from bs4 import BeautifulSoup
import pandas as pd
import concurrent.futures

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
}

tmpdir = "tmp-cannes"
awards_base_url = "https://www.festival-cannes.com/en/retrospective/{year}/awards/"
select_base_url = "https://www.festival-cannes.com/en/retrospective/{year}/selection/"

start_year = 1946
except_years = [1948, 1950]  # 1948, 1950年因为财政问题没有举办

filename = "cannes-festival"


def fetch_save_content(end_year):
    os.makedirs(tmpdir, exist_ok=True)

    while end_year >= start_year:
        if end_year in [1948, 1950]:
            continue
        target_url = awards_base_url.format(year=end_year)
        response = requests.get(target_url, headers=headers)
        response.raise_for_status()
        with open(
            f"./{tmpdir}/cannes-of-{end_year}-awards.html", "w", encoding="utf-8"
        ) as f:
            f.write(response.text)
            print(f"Successfully fetch {target_url}")

        target_url = select_base_url.format(year=end_year)
        response = requests.get(target_url, headers=headers)
        response.raise_for_status()
        with open(
            f"./{tmpdir}/cannes-of-{end_year}-selection.html", "w", encoding="utf-8"
        ) as f:
            f.write(response.text)
            print(f"Successfully fetch {target_url}")
        end_year -= 1


def parse_selection(year, edition, content: str) -> list:
    data = []

    # print(f"\033[34mWelcome to the {year} - {edition}th Cannes Festival!\033[0m")

    soup = BeautifulSoup(content, "html.parser")
    sections = soup.select("main .section", recursive=False)
    for section in sections:
        inner = section.select_one(".container__inner")
        title = inner.select_one("h2").text.strip()
        h3 = inner.select_one("h3")
        if h3:
            comment = h3.text.strip()
            title = f"{title} ({comment})"

        container = inner.select_one("div.list_container")
        items = container.select("div.list_item", recursive=False)
        for item in items:
            img = item.attrs["data-over-src"]
            item_content = item.select_one("div.list_item__content")
            cannes_link = item_content.select_one("a").attrs["href"]
            part1 = re.sub(r"\s+", " ", item_content.select_one("a").text.strip())
            span = item_content.select_one("span")
            part2 = ""
            if span:
                part2 = (
                    span.text.strip()
                    .replace("de ", "")
                    .replace("pour ", "")
                    .replace("– ", "")
                    .strip(". ")
                )

            data.append(
                {
                    "year": year,
                    "edition": edition,
                    "title": title,
                    "part1": part1,
                    "part2": part2,
                    "cannes_link": cannes_link,
                    "img": img,
                }
            )

    return data


def parse_awards(year, edition, content: str) -> list:
    data = []

    soup = BeautifulSoup(content, "html.parser")
    sections = soup.select("main .section", recursive=False)
    for section in sections:
        inner = section.select_one(".container__inner")
        title = inner.select_one("h2").text.strip()
        h3 = inner.select_one("h3")
        comment = ""
        if h3:
            comment = h3.text.strip()
            title = f"{title} ({comment})"

        container = inner.select_one("div.list_container")
        items = container.select("div.list_item", recursive=False)
        for item in items:
            img = item.attrs["data-over-src"]
            item_content = item.select_one("div.list_item__content")
            cannes_link = item_content.select_one("a").attrs["href"]
            award = item_content.select_one("div.list_item__award").text.strip()
            first_div = item_content.select_one("div.block", recursive=False)
            part1 = re.sub(r"\s+", " ", first_div.select_one("a").text.strip())
            span = first_div.select_one("span")
            part2 = ""
            if span:
                part2 = (
                    re.sub(r"\s+", " ", span.text.strip())
                    .replace("de ", "")
                    .replace("pour ", "")
                )

            data.append(
                {
                    "year": year,
                    "edition": edition,
                    "title": title,
                    "part1": part1,
                    "part2": part2,
                    "award": award,
                    "cannes_link": cannes_link,
                    "img": img,
                }
            )

            print(
                f"{year} - {edition}th {title}({comment}) - \033[32m{award}\033[0m: {part1} - {part2}"
            )
    return data


def fetch_and_parse(end_year):
    with open(
        f"./{tmpdir}/cannes-of-{end_year}-awards.html", "r", encoding="utf-8"
    ) as f:
        content = f.read()
    awards = parse_awards(end_year, end_year - start_year + 1, content)

    with open(
        f"./{tmpdir}/cannes-of-{end_year}-selection.html", "r", encoding="utf-8"
    ) as f:
        content = f.read()
    selection = parse_selection(end_year, end_year - start_year + 1, content)

    return awards, selection


def main():
    # fetch_save_content(1961)

    selection_of_all_year = []
    awards_of_all_year = []

    end_year = 2024

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(fetch_and_parse, year)
            for year in range(start_year, end_year + 1)
            if year not in except_years
        ]
        for future in concurrent.futures.as_completed(futures):
            awards, selection = future.result()
            awards_of_all_year.extend(awards)
            selection_of_all_year.extend(selection)

    selection_of_all_year.sort(key=lambda x: x["year"], reverse=True)
    awards_of_all_year.sort(key=lambda x: x["year"], reverse=True)

    save_to_excel(selection_of_all_year, awards_of_all_year, filename)


def save_to_excel(selection: list, awards: list, path: str) -> None:
    with pd.ExcelWriter(path + ".xlsx", engine="openpyxl") as writer:
        df = pd.DataFrame(selection)
        df.to_excel(writer, sheet_name="selection", index=False)

        df = pd.DataFrame(awards)
        df.to_excel(writer, sheet_name="awards", index=False)

        print(f"Data saved to {os.getcwd()}\\{path}.xlsx")


if __name__ == "__main__":
    main()
