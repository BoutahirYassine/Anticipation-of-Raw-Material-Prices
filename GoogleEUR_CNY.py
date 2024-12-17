# -*- coding: utf-8 -*-
"""
Created on Mon Dec 16 16:17:17 2024

@author: boutahirya
"""

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
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


# Set up Selenium WebDriver
options = Options()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
options.add_argument("--start-maximized")
service = Service('chromedriver-win64/chromedriver.exe')  # Replace with your chromedriver path
driver = webdriver.Chrome(service=service, options=options)

# Open the webpage
url = 'https://www.google.com/finance/quote/EUR-CNY?sa=X&sqi=2&ved=2ahUKEwiDr5HNxayKAxUCaqQEHbUrG0cQmY0JegQIIhAu'
driver.get(url)

print(driver.page_source)



# # Pause pour s'assurer que la page est entièrement chargée
# time.sleep(5)

# # Localiser le taux de change en utilisant la classe ou l'attribut qui contient la valeur
# try:
#     exchange_rate = WebDriverWait(driver, 10).until(
#         EC.visibility_of_element_located((By.CSS_SELECTOR, "div.YMlKec.fxKbKc"))
#     ).text
#     print(f"Taux de change EUR/CNY : {exchange_rate}")

# except Exception as e:
#     print(f"Erreur : {e}")

# Fermer le navigateur
driver.quit()


# # Step 2: Prepare the data
# if rate_element.text:
#     today_date = datetime.now().strftime('%Y-%m-%d')  # Get today's date in YYYY-MM-DD format
#     data_to_add = [today_date, rate_element.text]

#     # Step 3: Update the Excel file
#     file_name = r'DataBI/euro_to_dollar.xlsx'

#     # Check if the Excel file already exists
#     if os.path.exists(file_name):
#         # Load the existing workbook
#         workbook = openpyxl.load_workbook(file_name)
#         sheet = workbook.active  # Use the active sheet
#     else:
#         # Create a new workbook and add headers
#         workbook = openpyxl.Workbook()
#         sheet = workbook.active
#         sheet.append(['Date', 'Rate'])  # Add headers

#     # Append the new data (Date in column A, Rate in column B)
#     sheet.append(data_to_add)

#     # Save the workbook
#     workbook.save(file_name)
#     print(f"Data successfully added to '{file_name}': {data_to_add}")
    
# Close the driver
# driver.quit()
