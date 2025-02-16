import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook

start_year = 1929
base_url = "https://www.oscars.org/oscars/ceremonies/"
tmpdir = "tmp-oscars"

filename = "oscars-of-all-years"


def fetch_save_content():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    driver_path = "D:\\DevTools\\browserdriver\\chromedriver-win64\\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)

    os.makedirs(tmpdir, exist_ok=True)
    current_year = datetime.now().year
    while current_year >= start_year:
        target_url = base_url + str(current_year)
        driver.get(target_url)
        with open(
            f"./{tmpdir}/oscars-of-{current_year}.html", "w", encoding="utf-8"
        ) as f:
            f.write(driver.page_source)
            print(f"Successfully fetch {target_url}")
        current_year -= 1
    driver.quit()


def parse_content(year: int) -> list:
    with open(f"./{tmpdir}/oscars-of-{year}.html", "r", encoding="utf-8") as f:
        content = f.read()

    soup = BeautifulSoup(content, "html.parser")
    content = soup.find("div", id="tabSectionsContent")
    categories = content.find("div", class_="field--name-field-award-categories")
    items = categories.find_all("div", class_="field__item", recursive=False)

    data = []
    edition = year - start_year + 1  # 第几届奥斯卡

    print(f"\033[34mWelcome to the {year} - {edition}th Oscars!\033[0m")

    for item in items:
        category = item.find(
            "div", class_="field--name-field-award-category-oscars"
        ).text
        honorees_pdiv = item.find("div", class_="field--name-field-award-honorees")
        honorees = honorees_pdiv.find_all("div", class_="field__item", recursive=False)

        for honoree in honorees:
            win = False
            honoree_type = honoree.find("div", class_="field--name-field-honoree-type")
            if honoree_type is not None:
                win = honoree_type.text.strip() == "Winner"

            parts = honoree.find_all("div", class_="field__item")

            part1 = parts[0].text.strip()
            if len(parts) > 1:
                part2 = parts[1].text.strip()
            else:
                part2 = ""

            data.append(
                {
                    "year": year,
                    "edition": edition,
                    "category": category,
                    "part1": part1,
                    "part2": part2,
                    "win": win,
                }
            )

            print(
                f"{edition}th {category}: {part1} - {part2} ({'Winner' if win else 'Nominee'})"
            )
    return data


def save_to_excel(data: list, path: str) -> None:
    df = pd.DataFrame(data)
    df.to_excel(f"{path}.xlsx", sheet_name="oscars", index=False)

    wb = load_workbook(f"{path}.xlsx")

    wb.save(f"{path}.xlsx")
    print(f"Data saved to {os.getcwd()}\\{path}.xlsx")


def main():
    # fetch_save_content()
    current_year = datetime.now().year
    data = []
    while current_year >= start_year:
        data.extend(parse_content(current_year))
        current_year -= 1
    save_to_excel(data, filename)


if __name__ == "__main__":
    main()
