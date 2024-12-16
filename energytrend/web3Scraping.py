# -*- coding: utf-8 -*-
"""
Created on Mon Oct 21 10:16:35 2024

@author: boutahirya
"""

from bs4 import BeautifulSoup
import pandas as pd
import requests
import os

url = 'https://www.energytrend.com/solar-price.html'
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Define a list to store the items and prices
items_prices = []

# Function to extract prices from a table
def extract_prices(table):
    rows = table.find_all('tr')
    for row in rows:
        columns = row.find_all('td')
        if len(columns) >= 4:  # Ensure we have enough columns
            item = columns[1].text.strip()
            high_price = columns[2].text.strip()
            low_price = columns[3].text.strip()
            avg_price = columns[4].text.strip()
            chg = columns[5].text.strip() if len(columns) > 5 else "N/A"
            items_prices.append({'Item': item, 'High': high_price, 'Low': low_price, 'Avg': avg_price, 'Chg': chg})

# Find all relevant tables in the HTML
tables = soup.find_all('table')

# Extract data from each table
for table in tables:
    extract_prices(table)

# Convert the extracted data into a DataFrame
df_prices = pd.DataFrame(items_prices)

import openpyxl
from datetime import datetime

# Obtenir la date actuelle sous forme de chaîne au format "YYYY-MM-DD"
current_date = datetime.now().strftime("%Y-%m-%d")
# Nom du fichier avec la date
filename = f'files\energytrend_{current_date}.xlsx'

if os.path.isfile(filename):
    print("fichier existant")
else:
    # Saving the DataFrame to an Excel file
    df_prices.to_excel(filename, index=False)
    
    print("SAVED 3")