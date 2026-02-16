import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException


SECTORS = {
    1: "НЕФТЕГАЗ",
    2: "БАНКИ",
    14: "ФИНАНСЫ",
    3: "МЕТАЛЛУРГИЯ черн.",
    21: "МЕТАЛЛУРГИЯ цвет.",
    23: "ДРАГ.МЕТАЛЛЫ",
    18: "ГОРНОДОБЫВАЮЩИЕ",
    17: "ХИМИЯ удобрения",
    24: "ХИМИЯ разное",
    4: "Э/ГЕНЕРАЦИЯ",
    19: "ЭЛЕКТРОСЕТИ",
    20: "ЭНЕРГОСБЫТ",
    5: "РИТЕЙЛ",
    13: "ПОТРЕБ",
    26: "Агропром и Пищепром",
    27: "Промышленность разное",
    6: "ТЕЛЕКОМ",
    25: "ИНТЕРНЕТ",
    15: "HIGH TECH",
    28: "Производство Софта",
    29: "Фармацевтика",
    16: "МЕДИА",
    7: "ТРАНСПОРТ",
    8: "СТРОИТЕЛИ",
    9: "МАШИНОСТРОЕНИЕ",
    10: "ТРЕТИЙ ЭШЕЛОН"
}


BASE_URL_PE_YEARS = "https://smart-lab.ru/q/shares_fundamental4/?sector_id%5B%5D={}&field=p_e"
BASE_URL_PE_CURRENT = "https://smart-lab.ru/q/shares_fundamental2/?sector_id%5B%5D={}&field=p_e"


def parse_float(x):
    try:
        x = x.replace(" ", "").replace(",", ".")
        return float(x)
    except:
        return None


def safe_sheet_name(name):
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for ch in invalid_chars:
        name = name.replace(ch, ' ')
    return name[:30]


def load_sector_years(driver, sector_id, sector_name):
    print(f"\nЗагружаем сектор (история): {sector_name}")

    url = BASE_URL_PE_YEARS.format(sector_id)
    driver.get(url)

    try:
        table = WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CLASS_NAME, "simple-little-table"))
        )
    except TimeoutException:
        print("Таблица не найдена")
        return pd.DataFrame()

    rows = table.find_elements(By.TAG_NAME, "tr")
    header_cells = rows[0].find_elements(By.TAG_NAME, "th")
    headers = [h.text.strip() for h in header_cells]

    year_columns = headers[5:-2]

    data = []

    for row in rows[1:]:
        cells = [c.text.strip() for c in row.find_elements(By.TAG_NAME, "td")]
        if len(cells) < 6:
            continue

        name = cells[1]
        ticker = cells[2]

        values = []
        for cell in cells[5:5 + len(year_columns)]:
            values.append(parse_float(cell))

        row_dict = {
            "Название": name,
            "Тикер": ticker,
            "Сектор": sector_name
        }

        for year, val in zip(year_columns, values):
            row_dict[year] = val

        data.append(row_dict)

    return pd.DataFrame(data)


def load_pe_filter(driver, sector_id, sector_name):
    print(f"Проверяем P/E фильтр: {sector_name}")

    url = BASE_URL_PE_CURRENT.format(sector_id)
    driver.get(url)

    try:
        table = WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CLASS_NAME, "simple-little-table"))
        )
    except TimeoutException:
        print("Нет таблицы P/E")
        return pd.DataFrame()

    rows = table.find_elements(By.TAG_NAME, "tr")

    data = []
    sector_average = None

    for row in rows[1:]:
        cells = row.find_elements(By.TAG_NAME, "td")

        if not cells:
            continue

        name = cells[0].text.strip()

        # строка "Среднее"
        if "сред" in name.lower():
            if len(cells) > 1:
                sector_average = parse_float(cells[1].text)
                print("Средний P/E по сектору:", sector_average)
            break

        # обычные строки компаний
        if len(cells) < 6:
            continue

        company_name = cells[1].text.strip()
        ticker = cells[2].text.strip()
        pe = parse_float(cells[5].text)

        if pe is None:
            continue

        data.append({
            "Название": company_name,
            "Тикер": ticker,
            "Сектор": sector_name,
            "P/E": pe
        })

    if sector_average is None:
        print("Среднее не найдено")
        return pd.DataFrame()

    df = pd.DataFrame(data)
    print("Всего компаний:", len(df))

    filtered = df[
        (df["P/E"] > 0) &
        (df["P/E"] < 15) &
        (df["P/E"] < sector_average)
    ].copy()

    print("После фильтра:", len(filtered))

    filtered["Средний P/E по сектору"] = sector_average

    return filtered


def main():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    output_path = r"C:\Users\Irina\Desktop\pe_all_sectors.xlsx"

    pe_filtered_all = []

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # --- собираем P/E фильтр ---
        for sector_id, sector_name in SECTORS.items():
            df_filter = load_pe_filter(driver, sector_id, sector_name)
            if not df_filter.empty:
                pe_filtered_all.append(df_filter)

        # --- записываем первую вкладку ---
        if pe_filtered_all:
            final_filter_df = pd.concat(pe_filtered_all, ignore_index=True)
            final_filter_df.to_excel(writer, sheet_name="P_E < 15", index=False)

        # --- записываем исторические таблицы ---
        for sector_id, sector_name in SECTORS.items():
            df = load_sector_years(driver, sector_id, sector_name)

            if df.empty:
                continue

            year_columns = df.columns[3:]
            averages = df[year_columns].mean(numeric_only=True)

            avg_row = {
                "Название": "СРЕДНЕЕ ПО СЕКТОРУ",
                "Тикер": "",
                "Сектор": sector_name
            }

            for col in year_columns:
                avg_row[col] = averages.get(col, None)

            df_with_avg = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)

            sheet_name = safe_sheet_name(sector_name)
            df_with_avg.to_excel(writer, sheet_name=sheet_name, index=False)

    driver.quit()

    print("\nГотово")
    print("Файл:", output_path)


if __name__ == "__main__":
    main()
