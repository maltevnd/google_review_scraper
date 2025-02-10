import pandas as pd
import urllib.parse as urlparse
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup

def open_chrome_scrape(name, street_address, postal_code, city, exceptions_list):
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    search_query = f"{name}, {street_address}, {postal_code}, {city}"
    google_search_url = f"https://www.google.com/maps/search/{urlparse.quote_plus(search_query)}"
    driver.get(google_search_url)
    time.sleep(2)

    try:
        driver.find_element(By.XPATH, "/html/body/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button").click()
        time.sleep(2)
    except Exception:
        pass

    try:
        driver.find_element(By.XPATH, "/html/body/div[2]/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[3]/div/div/button[2]").click()
        time.sleep(3)
    except NoSuchElementException:
        exceptions_list.append({
            "Name": name, "Street": street_address, "Postal Code": postal_code, "City": city,
            "Search URL": google_search_url, "Timestamp": datetime.now()
        })
        driver.quit()
        return None
    except Exception as e:
        exceptions_list.append({
            "Name": name, "Street": street_address, "Postal Code": postal_code, "City": city,
            "Search URL": google_search_url, "Timestamp": datetime.now()
        })
        driver.quit()
        return None

    try:
        reviews_container = driver.find_element(By.XPATH, "//div[contains(@class, 'm6QErb DxyBCb kA9KIf dS8AEf XiKgde')]")
    except NoSuchElementException:
        exceptions_list.append({
            "Name": name, "Street": street_address, "Postal Code": postal_code, "City": city,
            "Search URL": google_search_url, "Timestamp": datetime.now()
        })
        driver.quit()
        return None

    last_height = driver.execute_script("return arguments[0].scrollHeight", reviews_container)
    while True:
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", reviews_container)
        time.sleep(2)
        new_height = driver.execute_script("return arguments[0].scrollHeight", reviews_container)
        if new_height == last_height:
            break
        last_height = new_height

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()

    review_containers = soup.find_all('div', class_="jftiEf fontBodyMedium")
    reviews_data = [{
        "Reviewer Name": container.find('div', class_='d4r55').text.strip() if container.find('div', class_='d4r55') else None,
        "Review Date": container.find('span', class_='rsqaWe').text.strip() if container.find('span', class_='rsqaWe') else None,
        "Stars": container.find('span', class_='kvMYJc')['aria-label'] if container.find('span', class_='kvMYJc') else None,
        "Review Text": container.find('span', class_='wiI7pd').text.strip() if container.find('span', class_='wiI7pd') else None
    } for container in review_containers]

    df = pd.DataFrame(reviews_data)
    df["Name"] = name
    df["Agency Street"] = street_address
    df["Agency plz"] = postal_code
    df["Agency City"] = city
    df["Timestamp"] = datetime.now()

    return df

def scrape_all_agencies():
    file_path = 'your_input_file.xlsx'
    rand_file = pd.read_excel(file_path)
    all_reviews_df = pd.DataFrame()
    exceptions_list = []

    for idx, row in rand_file.iterrows():
        name = row['Name']
        street_address = row['Street']
        postal_code = row['PLZ']
        city = row['City']
        print(f"Scraping reviews for {name}, {street_address}, {postal_code}, {city}...")
        name_reviews_df = open_chrome_scrape(name, street_address, postal_code, city, exceptions_list)
        if name_reviews_df is not None:
            all_reviews_df = pd.concat([all_reviews_df, name_reviews_df], ignore_index=True)

    all_reviews_df.to_excel("result.xlsx", index=False)
    if exceptions_list:
        pd.DataFrame(exceptions_list).to_excel("error_log.xlsx", index=False)
    print("All data saved in 'result.xlsx'.")
    print("Exception log saved in 'error_log.xlsx'.")

scrape_all_agencies()