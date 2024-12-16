# -*- coding: utf-8 -*-
"""
Created on Wed Dec  4 10:36:41 2024

@author: boutahirya
"""

import requests
from bs4 import BeautifulSoup
from datetime import datetime
import openpyxl
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Set up Selenium WebDriver
options = Options()
options.add_argument('--headless')  # Run in headless mode
service = Service(r'chromedriver-win64\chromedriver.exe')  # Replace with your chromedriver path
driver = webdriver.Chrome(service=service, options=options)

# Open the webpage
url = 'https://fr.investing.com/currencies/eur-usd-news'
driver.get(url)

# Wait for the element to load and extract the EUR/USD rate
try:
    eur_usd_element = driver.find_element(By.CSS_SELECTOR, "div[data-test='instrument-price-last']")
    print(f"EUR/USD Rate: {eur_usd_element.text}")
except Exception as e:
    print("EUR/USD rate not found:", e)


# Step 2: Prepare the data
if eur_usd_element.text:
    today_date = datetime.now().strftime('%Y-%m-%d')  # Get today's date in YYYY-MM-DD format
    data_to_add = [today_date, eur_usd_element.text]

    # Step 3: Update the Excel file
    file_name = r'DataBI\euro_to_dollar.xlsx'

    # Check if the Excel file already exists
    if os.path.exists(file_name):
        # Load the existing workbook
        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active  # Use the active sheet
    else:
        # Create a new workbook and add headers
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['Date', 'Rate'])  # Add headers

    # Append the new data (Date in column A, Rate in column B)
    sheet.append(data_to_add)

    # Save the workbook
    workbook.save(file_name)
    print(f"Data successfully added to '{file_name}': {data_to_add}")
    
# Close the driver
driver.quit()