import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


SECTORS = {
    1: "НЕФТЕГАЗ",
    2: "БАНКИ",
    14: "ФИНАНСЫ",
    3: "МЕТАЛЛУРГИЯ черн.",
    21: "МЕТАЛЛУРГИЯ цвет.",
    22: "МЕТАЛЛУРГИЯ разное",
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
    10: "ТРЕТИЙ ЭШЕЛОН",
    11: "НЕПУБЛИЧНЫЕ",
    12: "ДРУГОЕ"
}


BASE_URL_YEARS = "https://smart-lab.ru/q/shares_fundamental4/?sector_id%5B%5D={}&field=p_e"
BASE_URL_LTM = "https://smart-lab.ru/q/shares_fundamental2/?sector_id%5B%5D={}&field=p_e"


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


# -------- P/E по годам --------
def load_sector_years(driver, sector_id, sector_name):
    print(f"\nЗагружаем сектор (годы): {sector_name}")

    driver.get(BASE_URL_YEARS.format(sector_id))

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

        row_dict = {
            "Название": name,
            "Тикер": ticker,
            "Сектор": sector_name
        }

        values = [parse_float(x) for x in cells[5:5+len(year_columns)]]

        for year, val in zip(year_columns, values):
            row_dict[year] = val

        data.append(row_dict)

    return pd.DataFrame(data)


# -------- P/E LTM --------
def load_sector_pe_ltm(driver, sector_id, sector_name):
    print(f"Сканируем LTM P/E: {sector_name}")

    driver.get(BASE_URL_LTM.format(sector_id))

    try:
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "simple-little-table"))
        )
    except TimeoutException:
        return pd.DataFrame()

    rows = table.find_elements(By.TAG_NAME, "tr")

    data = []

    for row in rows[1:]:
        cells = row.find_elements(By.TAG_NAME, "td")

        if len(cells) < 6:
            continue

        name = cells[1].text.strip()

        if name in ["Всего:", "Среднее:"]:
            continue

        ticker = cells[2].text.strip()
        pe_value = parse_float(cells[5].text.strip())

        if pe_value is None:
            continue

        # ФИЛЬТР
        if 0 < pe_value < 15:
            data.append({
                "Название": name,
                "Тикер": ticker,
                "Сектор": sector_name,
                "P/E": pe_value
            })

    return pd.DataFrame(data)


# -------- MAIN --------
def main():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    output_path = r"C:\Users\Irina\Desktop\pe_all_sectors.xlsx"

    filtered_companies = []
    sector_data = {}

    # собираем данные
    for sector_id, sector_name in SECTORS.items():

        # данные по годам
        df_years = load_sector_years(driver, sector_id, sector_name)

        if df_years is not None and not df_years.empty:
            year_columns = df_years.columns[3:]
            averages = df_years[year_columns].mean(numeric_only=True)

            avg_row = {
                "Название": "СРЕДНЕЕ ПО СЕКТОРУ",
                "Тикер": "",
                "Сектор": sector_name
            }

            for col in year_columns:
                avg_row[col] = averages.get(col, None)

            df_years = pd.concat([df_years, pd.DataFrame([avg_row])], ignore_index=True)
            sector_data[sector_name] = df_years

        # LTM фильтр
        df_filtered = load_sector_pe_ltm(driver, sector_id, sector_name)

        if df_filtered is not None and not df_filtered.empty:
            filtered_companies.append(df_filtered)

    driver.quit()

    # -------- запись Excel --------
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # первая вкладка
        if filtered_companies:
            result_df = pd.concat(filtered_companies, ignore_index=True)
            result_df.to_excel(writer, sheet_name="P_E < 15", index=False)

        # вкладки по секторам
        for sector_name, df in sector_data.items():
            sheet_name = safe_sheet_name(sector_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("\nГотово ✅")
    print("Файл сохранён:", output_path)


if __name__ == "__main__":
    main()
