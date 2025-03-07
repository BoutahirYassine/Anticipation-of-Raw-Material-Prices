# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 10:09:13 2025

@author: boutahirya
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Jan 16 13:39:43 2025

@author: boutahirya
"""

from bs4 import BeautifulSoup
import requests
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os

import requests
from bs4 import BeautifulSoup
from datetime import datetime
import os
from openpyxl import Workbook, load_workbook

def fetch_price(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        if url == "https://www.metal.com/Solar/202303220001":
            brut_section = soup.find('div', text='Original')
            if brut_section:              
                price_element = brut_section.find_next('div', class_='price___2mpJr')               # price WAFER
                
                
        elif url == "https://www.metal.com/en/prices/202210280001":
            brut_section = soup.find('div', text='Original')
            if brut_section:
                price_element = brut_section.find_next('div', class_='PriceDisplay_price__VYiMd')  #price CELL
                
                
        elif url == "https://www.metal.com/Solar/202112230003":
            brut_section = soup.find('div', text='Original')
            if brut_section:
                price_element = brut_section.find_next('div', class_='price___2mpJr')               #Silver Rear side
                
                
        elif url == "https://www.metal.com/Solar/202112230004":
            brut_section = soup.find('div', text='Original')
            if brut_section:
                price_element = brut_section.find_next('div', class_='price___2mpJr')              # Silver busbar front-side
                
                
        elif url == "https://www.metal.com/Solar/202112230005":
            brut_section = soup.find('div', text='Original')
            if brut_section:
                price_element = brut_section.find_next('div', class_='price___2mpJr')              #Silver Finger Front-side 
                
                
        # USD/CNY
        elif url == "https://www.metal.com/exchange-rate/200002250101":
            price_element1 = soup.find("span", class_="strong___Js3_I priceDown___2TbRQ")
            price_element2 = soup.find("span", class_="strong___Js3_I priceUp___3Mgsl")
            price_element3 = soup.find("span", class_="strong___3sC58 priceUp___3Mgsl")
            
            if price_element1:
                price_element = price_element1
            elif price_element3:
                price_element = price_element3
            else:
                price_element = price_element2
            
        # Find the parent div and extract the price MODULE
        elif url == "https://www.metal.com/Solar/202403260002":
            price_element = soup.find("span", class_="strong___3sC58")
        else:
            brut_section = soup.find('div', text='Original')
            if brut_section:
                price_element = brut_section.find_next('div', class_='price___2mpJr') 
                
        if price_element:
            return price_element.text.strip()
    return None

def update_excel(sheet_name, url, workbook):
    price = fetch_price(url)
    if price:
        today_date = datetime.now().strftime("%Y-%m-%d")
        
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
        else:
            sheet = workbook.create_sheet(sheet_name)
            sheet.append(["Date", "Price"])
        
        date_found = False
        for idx, row in enumerate(sheet.iter_rows(min_row=2, max_col=2, values_only=True), start=2):
            if row[0] == today_date:
                sheet.cell(row=idx, column=2, value=price)
                date_found = True
                print(f"Updated the price for {sheet_name} on {today_date}")
                break
        
        if not date_found:
            sheet.append([today_date, price])
            print(f"Added new data for {sheet_name} on {today_date}")
    else:
        print(f"Failed to fetch price for {sheet_name}")

def main():
    excel_file = "DataBI/SMM/Solar_Prices.xlsx"
    #excel_file = "Solar_Prices2.xlsx"
    if os.path.exists(excel_file):
        workbook = load_workbook(excel_file)
    else:
        workbook = Workbook()
        workbook.remove(workbook.active)  # Remove default sheet
    
    update_excel("N-Dense", "https://www.metal.com/Solar/202501060001", workbook)
    update_excel("N-Type", "https://www.metal.com/Solar/202501060003", workbook)
    update_excel("N-type Wafer", "https://www.metal.com/Solar/202303220001", workbook) ##
    update_excel("Cell Topcon 183mm", "https://www.metal.com/en/prices/202210280001", workbook)
    update_excel("Module Topcon 183mm", "https://www.metal.com/Solar/202403260002", workbook)
    update_excel("Silver Rear_side", "https://www.metal.com/Solar/202112230003", workbook)
    update_excel("Silver Busbar front-side", "https://www.metal.com/Solar/202112230004", workbook)
    update_excel("Silver finger front-side", "https://www.metal.com/Solar/202112230005", workbook)
    update_excel("USD_CNY", "https://www.metal.com/exchange-rate/200002250101", workbook)
    
    workbook.save(excel_file)
    print(f"Data saved to {excel_file}")

if __name__ == "__main__":
    main()