from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import os
import pandas as pd
from datetime import datetime, timedelta
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from dotenv import load_dotenv

load_dotenv()

password = os.getenv("password")
login = os.getenv("login")

# Задаем нужный диапазон дат
months_to_fill = pd.date_range(start="2025-01-01", end="2025-03-01", freq="MS")

new_data = []

save_path = os.getcwd()

options = webdriver.ChromeOptions()
prefs = {
        "download.default_directory": save_path,  
        "download.prompt_for_download": False,    
        "download.directory_upgrade": True,        
        "safebrowsing.enabled": True            
    }
options.add_experimental_option("prefs", prefs)

def autorisation(login, passsword, options=options):
        browser = webdriver.Chrome(options=options)
        browser.get('https://passport.yandex.ru/auth/add/login?retpath=https%3A%2F%2Fwordstat.yandex.ru%2F%3Fregion%3Dall%26view%3Dtable%26words%3Dtest')

        time.sleep(2.5)
        
        login_input = browser.find_element("id", "passp-field-login")
        login_input.send_keys(login)
        browser.find_element(By.XPATH, '//*[@id="passp:sign-in"]').click()
        time.sleep(0.5)

        login_input = browser.find_element("id", "passp-field-passwd")
        login_input.send_keys(passsword)
        browser.find_element(By.XPATH, '//*[@id="passp:sign-in"]').click()

        time.sleep(3)

        try:
          browser.find_element(By.CLASS_NAME, 'shepherd-cancel-icon').click()
        except:
          pass

        return browser

browser = autorisation(password, login)

def get_query_counts(keyword, browser = browser):

        renamed_files = []

        def rename_last_downloaded_file(save_path, new_name):
            time.sleep(1)
            files = os.listdir(save_path)
            files = [f for f in files if "wordstat_dynamic" in f]
            if files:
                latest_file = max(files, key=lambda x: os.path.getctime(os.path.join(save_path, x)))
                new_filename = f"{new_name}.csv"
                os.rename(os.path.join(save_path, latest_file), os.path.join(save_path, new_filename))
            return new_filename

        try:
            input_field = browser.find_element(By.CLASS_NAME, 'textinput__control')
        
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.DELETE)
        
            input_field.send_keys(keyword)
        
            search_button = browser.find_element(By.CLASS_NAME, 'wordstat__search-button')
        
            search_button.click()
            time.sleep(5)
            try:
                download_button = browser.find_element(By.CLASS_NAME, 'save-button')
            
                download_button.click()
                time.sleep(1)
                download_link = browser.find_element(By.CSS_SELECTOR, 'a[download="wordstat_dynamic"]')
                download_link.click()
                time.sleep(2)
                renamed_files.append(rename_last_downloaded_file(save_path, keyword))
            except:
                pass

            try:
                browser.find_element(By.CSS_SELECTOR, ".icon.icon_type_close").click()
            except:
                pass
        except:
            browser.refresh()
            time.sleep(2)
            input_field = browser.find_element(By.CLASS_NAME, 'textinput__control')
        
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.DELETE)
        
            input_field.send_keys(keyword)
        
            search_button = browser.find_element(By.CLASS_NAME, 'wordstat__search-button')
        
            search_button.click()
            time.sleep(5)
            try:
                download_button = browser.find_element(By.CLASS_NAME, 'save-button')
            
                download_button.click()
                time.sleep(1)
                download_link = browser.find_element(By.CSS_SELECTOR, 'a[download="wordstat_dynamic"]')
                download_link.click()
                time.sleep(1.5)
                renamed_files.append(rename_last_downloaded_file(save_path, keyword))
            except:
                pass

def convert_month_year_to_date(month_year_str):
    months = {
        "январь": "01", "февраль": "02", "март": "03", "апрель": "04",
        "май": "05", "июнь": "06", "июль": "07", "август": "08", 
        "сентябрь": "09", "октябрь": "10", "ноябрь": "11", "декабрь": "12"
    }
    
    month_name, year = month_year_str.lower().split()
    
    month_number = months.get(month_name, None)
    
    if month_number:
        return f"01.{month_number}.{year}"
    else:
        return None  
    
def get_quarter(date):
    month = date.month
    if month in [1, 2, 3]:
        return "Q1"
    elif month in [4, 5, 6]:
        return "Q2"
    elif month in [7, 8, 9]:
        return "Q3"
    else:
        return "Q4"

new_data = []

wordstat_queries = pd.read_excel("wordstat_queries.xlsx")
keywords = wordstat_queries['Запрос'].unique()

for keyword in keywords:
        keyword_counts = get_query_counts(keyword)
        
        try:
            word_df = pd.read_csv(f"{keyword}.csv", delimiter=';')
            os.remove(f"{keyword}.csv")  
        except:
            print(keyword)

        for month in months_to_fill:
            year = month.year
            quarter = get_quarter(month)

            for _, row in word_df.iterrows():
                period_date = convert_month_year_to_date(row['Период'])
                try:
                    count = int(row['Число запросов'].replace(" ", ""))
                except:
                    count = int(row['Число запросов'])
    
                if pd.to_datetime(period_date, dayfirst=True, errors="coerce") == month:
                    new_data.append([year, quarter, period_date, keyword, count])

new_df = pd.DataFrame(new_data, columns=["Год", "Квартал", "Месяц", "Запрос", "Количество"])

browser.quit()

new_df.to_excel("result.xlsx")