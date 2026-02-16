import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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


BASE_URL = "https://smart-lab.ru/q/shares_fundamental4/?sector_id%5B%5D={}&field=p_e"


def parse_float(x):
    try:
        x = x.replace(" ", "").replace(",", ".")
        return float(x)
    except:
        return None


from selenium.common.exceptions import TimeoutException


def load_sector(driver, sector_id, sector_name):
    print(f"\nЗагружаем сектор: {sector_name}")

    url = BASE_URL.format(sector_id)
    driver.get(url)

    try:
        table = WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CLASS_NAME, "simple-little-table"))
        )
    except TimeoutException:
        print("Таблица не найдена — сектор пропущен")
        return pd.DataFrame()

    rows = table.find_elements(By.TAG_NAME, "tr")

    if len(rows) < 2:
        print("Нет строк с данными")
        return pd.DataFrame()

    # Заголовки
    header_cells = rows[0].find_elements(By.TAG_NAME, "th")
    headers = [h.text.strip() for h in header_cells]

    year_columns = headers[5:-2]

    print("Годы:", year_columns)

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

    print("Компаний:", len(data))
    return pd.DataFrame(data)

def safe_sheet_name(name):
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for ch in invalid_chars:
        name = name.replace(ch, ' ')
    return name[:30]  # ограничение Excel


def main():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    output_path = r"C:\Users\Irina\Desktop\pe_all_sectors.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for sector_id, sector_name in SECTORS.items():
            df = load_sector(driver, sector_id, sector_name)

            if df is not None and not df.empty:
                sheet_name = safe_sheet_name(sector_name)
                print("Записываем в Excel:", sector_name)

                # --- расчет среднего по сектору ---
                year_columns = df.columns[3:]  # первые 3 колонки: Название, Тикер, Сектор

                averages = df[year_columns].mean(numeric_only=True)

                avg_row = {
                    "Название": "СРЕДНЕЕ ПО СЕКТОРУ",
                    "Тикер": "",
                    "Сектор": sector_name
                }

                for col in year_columns:
                    avg_row[col] = averages.get(col, None)   # ВАЖНО: безопасное получение

                df_with_avg = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)


                try:
                    df_with_avg.to_excel(writer, sheet_name=sheet_name, index=False)

                except Exception as e:
                    print("Ошибка записи листа:", sector_name, e)

    driver.quit()

    print("\nГотово ✅")
    print("Файл сохранён в:", output_path)



if __name__ == "__main__":
    main()
