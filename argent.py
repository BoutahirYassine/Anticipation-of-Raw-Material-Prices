from bs4 import BeautifulSoup
import requests
import openpyxl
import os
from datetime import datetime

# URL cible
url = "https://www.bdor.fr/cours-argent"
file_name = 'DataBI/Argent_EUR.xlsx'

# Définir les cookies extraits
cookies = {
    "PHPSESSID": "95vbhpjinfj59jlg2grb182qg35",
    "_gcl_au": "1.1.511198851.1736945792",
    "_pk_id.6.a280": "f08a824fdca2523a.1736945792.1.1736945792.1736945792.",
    "_pk_ref.6.a280": "%5B%22%22%2C%22%22%2C1736945792%2C%22https%3A%2F%2Fgoogle.com%22%5D",
    "_pk_ses.6.a280": "1",
    "ms_conseiller_id102348": "9",
    "ms_consultationIdPageAEnregistrer": "172",
    "ms_consultationIdProduitAEnregistrer": "0",
    "ms_consultationIdRubriqueAEnregistrer": "0",
    "ms_idPanier": "5758104",
    "ms_timeConsultationAEnregistrer": "2025-01-15%2014%3A22%3A40",
    "msvd_idCookieUser": "NjExMTA1MzczOTg3OD==",
    "ssm_au_c": "k6/seAFFKjAwpB9qfnQ...",
}

# Fonction pour mettre à jour ou ajouter des données dans le fichier Excel
def update_or_replace_data(sheet, data_to_add):
    date_to_update, argent_gramme = data_to_add
    
    # Parcourir les lignes pour vérifier si la date existe déjà
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):  # Commence à partir de la 2ème ligne
        existing_date = row[0]  # Valeur dans la colonne A
        if existing_date.value and existing_date.value.date() == date_to_update.date():
            # Mettre à jour les colonnes correspondantes
            sheet.cell(row=row_idx, column=2, value=argent_gramme)  # Colonne B
            print(f"Mise à jour des données pour {date_to_update} : {argent_gramme}")
            return
    # Si la date n'existe pas, ajouter une nouvelle ligne
    sheet.append(data_to_add)
    print(f"Ajout des nouvelles données pour {date_to_update} : {argent_gramme}")

# Envoyer la requête avec les cookies
response = requests.get(url, cookies=cookies)

# Vérifier le statut de la requête
if response.status_code == 200:
    # Analyser le HTML avec BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')

    # Localiser le tableau contenant les données des métaux précieux
    tableau = soup.find("table", class_="contentTBC")
    if tableau:

        argent_gramme_price = None


        argent_row_gramme = tableau.find("td", text="1 Gramme")
        if argent_row_gramme:
            argent_price_cell = argent_row_gramme.find_next_sibling("td", class_="cellArgent")
            if argent_price_cell:
                argent_gramme_price = argent_price_cell.text.strip()
                argent_gramme_price = argent_gramme_price.replace('€', '').strip()  # Retire le symbole "€"
                print(argent_gramme_price)  # Résultat : "0.94"

        # Vérifier et afficher les prix extraits
        if argent_gramme_price:

            # Convertir les données en un format exploitable
            today = datetime.now()
            data_to_add = [today, argent_gramme_price]

            # Vérifier si le fichier Excel existe déjà
            if os.path.exists(file_name):
                # Charger le fichier Excel existant
                workbook = openpyxl.load_workbook(file_name)
                sheet = workbook.active  # Utiliser la feuille active
            else:
                # Créer un nouveau fichier Excel avec des en-têtes
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.append(['Date', '1 Once (€)'])  # Ajouter les en-têtes

            # Ajouter ou mettre à jour les données
            update_or_replace_data(sheet, data_to_add)

            # Sauvegarder le fichier Excel
            workbook.save(file_name)
            print(f"Les prix ont été enregistrés dans le fichier '{file_name}'.")
        else:
            print("Impossible d'extraire les prix de l'argent.")
    else:
        print("Impossible de trouver le tableau contenant les données.")
else:
    print(f"Erreur lors de la requête : {response.status_code}")
