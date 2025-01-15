from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Chemin vers votre ChromeDriver
chrome_driver_path = "chromedriver-linux64/chromedriver"
service = Service(chrome_driver_path)
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")  # Enlever "--headless" pour déboguer visuellement
driver = webdriver.Chrome(service=service, options=options)

try:
    # Charger la page
    url = "https://www.bullionbypost.fr/cours-dargent/an/grammes/USD/"
    driver.get(url)

    # Attendre que l'élément soit visible
    price_element = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='current-price']/span[@name='current_price_field']"))
    )
    
    # Extraire le texte
    current_price = price_element.text.strip()
    print(f"Prix actuel : {current_price}")

except Exception as e:
    print(f"Une erreur s'est produite : {e}")
finally:
    # Fermer le navigateur
    driver.quit()
