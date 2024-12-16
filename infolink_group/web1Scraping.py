import requests
from bs4 import BeautifulSoup
from web1excel import *
# URL de la page à scraper
url = 'https://www.infolink-group.com/spot-price/'

response = requests.get(url, timeout=10, verify=False)

# Vérifier si la requête a réussi
if response.status_code == 200:
    # Parser le contenu HTML avec BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Trouver toutes les balises <td> avec align="right"
    td_elements = soup.find_all('td', attrs={'align': 'right'})
    
    # Extraire les valeurs numériques présentes dans les balises <td>
    values = [td.get_text(strip=True) for td in td_elements]
    to_excel_combined(values)

else:
    print(f"Erreur lors de la requête : {response.status_code}")
