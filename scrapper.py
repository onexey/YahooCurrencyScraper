import calendar
import os
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

from currency_data import CurrencyData
from excel_exporter import export


def scrap_currencies():
    currency_data = []  # this will store raw data scrapped from the page

    currencies_url = "https://finance.yahoo.com/currencies"
    driver = webdriver.Chrome()
    driver.get(currencies_url)

    currencies_table = driver.find_element(By.XPATH, '//*[@id="mrt-node-Lead-4-YFinListTable"]')
    currencies_rows = currencies_table.find_elements(By.CLASS_NAME, "simpTblRow")

    for currencies_row in currencies_rows:
        symbol = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label=Symbol]")
        name = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label=Name]")
        last_price = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label='Last Price']")
        change = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label=Change]")
        percent_change = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label='% Change']")

        data = CurrencyData(symbol.text, name.text, last_price.text, change.text, percent_change.text)
        currency_data.append(data)

    print("Currencies download completed")

    export(currency_data)

    # normally, I would do this in app level instead of in browser
    driver.find_element(By.XPATH, "//th[.='% Change']").click()  # sort asc
    driver.find_element(By.XPATH, "//th[.='% Change']").click()  # sort desc

    currencies_rows = currencies_table.find_elements(By.CLASS_NAME, "simpTblRow")
    for currencies_row in currencies_rows[:5]:
        symbol = currencies_row.find_element(By.CSS_SELECTOR, "[aria-label=Symbol]").text

        download_history(symbol)

    print("History download completed")


def download_history(symbol):
    today = datetime.today()
    start_date = calendar.timegm(datetime(today.year, today.month, 1).timetuple())
    end_date = calendar.timegm(datetime(today.year, today.month, today.day).timetuple())
    history_page = f"https://finance.yahoo.com/quote/{symbol}/history?period1={start_date}&period2={end_date}&interval=1d&filter=history&frequency=1d&includeAdjustedClose=true"

    download_path = os.path.abspath("Currencies")
    opts = webdriver.ChromeOptions()
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36")
    prefs = {"download.default_directory": download_path}
    opts.add_experimental_option("prefs", prefs)

    history_driver = webdriver.Chrome(options=opts)
    history_driver.get(history_page)
    delay = 10  # seconds

    popup = WebDriverWait(history_driver, delay).until(
        EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Close']")))
    popup.click()

    # I am suspecting the url subdomain may change.
    # That's why I am not directly building the url.
    download_link = history_driver.find_element(By.XPATH, "//span[.='Download']/a")
    download_link.click()

    # This is just to be make sure.
    # In real life scenario I would prefer more robust "IsDownloadCompleted" algorithm
    time.sleep(5)

    history_driver.quit()
