import subprocess
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
import time
import json

# Setup and initial configurations
URL = "https://www.google.com/maps"
service = "SERVICE"  # e.g. catering, events, etc.
location = "LOCATION"  # e.g. London, Germany, etc.

print("Starting the web scraping script...")

options = Options()
options.add_argument('--headless=new')
driver = webdriver.Chrome(options=options)
print(f"Accessing URL: {URL}")
driver.get(URL)

# Accept cookies
try:
    print("Looking for accept cookies button...")
    accept_cookies = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button'))
    )
    accept_cookies.click()
    print("Accepted cookies.")
except Exception:
    print("No accept cookies button found or already accepted.")

# Search for results
print(f"Searching for: {service} in {location}")
input_field = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="searchboxinput"]'))
)
input_field.send_keys(f"{service} {location}")
input_field.send_keys(Keys.ENTER)
print("Search submitted.")

# Wait for results to load
time.sleep(5)

# Scroll and collect clickable result cards
print("Scrolling to load all business listings...")
scrollable_div_xpath = '//div[@role="feed"]'
scrollable_div = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, scrollable_div_xpath))
)

last_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
while True:
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
    time.sleep(2)
    new_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
    if new_height == last_height:
        break
    last_height = new_height

print("Finished scrolling. Collecting business profiles...")
business_cards = driver.find_elements(By.CLASS_NAME, "Nv2PK")

data = []

for i, card in enumerate(business_cards):
    try:
        print(f"Processing business {i+1}/{len(business_cards)}")
        driver.execute_script("arguments[0].scrollIntoView();", card)
        time.sleep(1)
        card.click()
        time.sleep(3)  # Let the profile panel load

        # Collect business data from the profile panel
        try:
            name = driver.find_element(By.CLASS_NAME, "DUwDvf").text
        except:
            name = "N/A"

        try:
            address = driver.find_element(By.XPATH, "//button[contains(@data-item-id, 'address')]//div[2]/div[1]").text
        except:
            address = "N/A"

        try:
            phone = driver.find_element(By.XPATH, "//button[contains(@data-item-id, 'phone')]//div[2]/div[1]").text
        except:
            phone = "N/A"

        try:
            website = driver.find_element(By.XPATH, "//a[contains(@data-item-id, 'authority')]" ).get_attribute("href")
        except:
            website = "N/A"

        try:
            stars = driver.find_element(By.CLASS_NAME, "F7nice").text
        except:
            stars = "N/A"

        try:
            reviews = driver.find_element(By.CLASS_NAME, "UY7F9").text.strip("()")
        except:
            reviews = "N/A"

        data.append({
            'Business Name': name,
            'Address': address,
            'Stars': stars,
            'Number of Reviews': reviews,
            'Phone Number': phone,
            'Website': website,
            'Email': ' ',
        })

        # Back to the results panel
        time.sleep(2)
        back_button = driver.find_element(By.CLASS_NAME, "RVQdVd")
        if back_button:
            back_button.click()
            time.sleep(2)

    except Exception as e:
        print(f"Error processing card {i+1}: {e}")
        continue

# Save to Excel
excel_file = f'{location}_{service}.xlsx'
df = pd.DataFrame(data)
df.to_excel(excel_file, index=False)

print(f"Data has been saved to {excel_file}")

# Save config
with open('config.json', 'w') as config_file:
    json.dump({ 'excel_file': excel_file }, config_file)
print("Configuration file created: config.json")

# Run the email extraction script
print("Calling the email extraction script...")
subprocess.run(['python', 'email_extraction_script.py'])
print("Email extraction script completed.")
