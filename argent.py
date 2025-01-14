from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

# Configuration de Selenium
service = Service("chromedriver-linux64/chromedriver")  # Remplacez par le chemin de chromedriver
options = Options()
options.add_argument("--headless")  # Optionnel : mode sans interface graphique
options.add_argument("--disable-dev-shm-usage")  # Réduction des problèmes de mémoire
options.add_argument("--no-sandbox")  # Pour certains environnements

# Initialisation du WebDriver
driver = webdriver.Chrome(service=service, options=options)

try:
    # Charger la page cible
    url = "https://www.bullionbypost.fr/cours-dargent/5ans/grammes/USD/"
    driver.get(url)

    # Attendre un délai initial pour s'assurer que la page se charge
    time.sleep(3)

    # Attendre que l'élément contenant le prix actuel soit visible
    wait = WebDriverWait(driver, 15)
    price_element = wait.until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='current-price']/span"))
    )

    # Extraire le texte de l'élément
    current_price = price_element.text.strip()
    print(f"Prix actuel : {current_price}")

except Exception as e:
    print(f"Erreur : {e}")

finally:
    # Fermer le navigateur
    driver.quit()
