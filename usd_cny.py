from bs4 import BeautifulSoup
import requests
from datetime import datetime
import openpyxl
import os

# Step 1: Scrape the conversion rate
url = "https://www.x-rates.com/calculator/?from=USD&to=CNY&amount=1"
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Locate the element containing the conversion rate
conversion_rate = soup.find('span', class_='ccOutputRslt')

if conversion_rate:
    rate = conversion_rate.text.split()[0]  # Extract only the numeric part
    print(f"Conversion Rate (1 USD to CNY): {rate}")
else:
    print("Conversion rate not found.")
    rate = None

# Step 2: Prepare the data
if rate:
    today_date = datetime.now().strftime('%Y-%m-%d')  # Get today's date in YYYY-MM-DD format
    data_to_add = [today_date, rate]

    # Step 3: Update the Excel file
    file_name = 'DataBI/dollar_to_yuan.xlsx'

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
