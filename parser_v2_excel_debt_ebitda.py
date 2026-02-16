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


BASE_URL = "https://smart-lab.ru/q/shares_fundamental2/?sector_id%5B%5D={}&field=debt_ebitda"


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
    return name[:30]  # ограничение Excel


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

    data = []

    for row in rows[1:]:
        cells = row.find_elements(By.TAG_NAME, "td")

        # строки "Всего" и "Среднее" обычно имеют меньше колонок
        if len(cells) < 6:
            continue

        name = cells[1].text.strip()
        ticker = cells[2].text.strip()
        value = cells[5].text.strip()

        if name in ["Всего:", "Среднее:"]:
            continue

        value = parse_float(value)

        data.append({
            "Название": name,
            "Тикер": ticker,
            "Сектор": sector_name,
            "Debt/EBITDA": value
        })

    print("Компаний:", len(data))
    return pd.DataFrame(data)


def main():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    output_path = r"C:\Users\Irina\Desktop\debt_ebitda_all_sectors.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for sector_id, sector_name in SECTORS.items():
            df = load_sector(driver, sector_id, sector_name)

            if df is not None and not df.empty:
                sheet_name = safe_sheet_name(sector_name)
                print("Записываем в Excel:", sector_name)

                # считаем среднее по сектору
                avg_value = df["Debt/EBITDA"].mean()

                avg_row = {
                    "Название": "СРЕДНЕЕ ПО СЕКТОРУ",
                    "Тикер": "",
                    "Сектор": sector_name,
                    "Debt/EBITDA": round(avg_value, 2) if pd.notna(avg_value) else None
                }

                df_with_avg = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)

                df_with_avg.to_excel(writer, sheet_name=sheet_name, index=False)

    driver.quit()

    print("\nГотово ✅")
    print("Файл сохранён в:", output_path)


if __name__ == "__main__":
    main()
