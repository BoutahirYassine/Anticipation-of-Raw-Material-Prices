import pandas as pd

# Charger les fichiers Excel
argent_file = "DataBI/N_type_Polysilicon.xlsx"
conversion_file = "DataBI/USD_CNY.xlsx"


argent_data = pd.read_excel(argent_file)
conversion_data = pd.read_excel(conversion_file)


# Renommer les colonnes pour éviter les problèmes d'espaces ou caractères spéciaux
argent_data.columns = ["date", "Silver price in grams (dollar)"]
conversion_data.columns = ["date", "Conversion rate (dollar to euro)"]

# Convertir les colonnes 'date' en type datetime pour faciliter le croisement
argent_data['date'] = pd.to_datetime(argent_data['date'], format="%d/%m/%Y")
conversion_data['date'] = pd.to_datetime(conversion_data['date'], format="%d/%m/%Y")

# Joindre les deux DataFrames sur la colonne 'date'
merged_data = pd.merge(argent_data, conversion_data, on="date", how="inner")

# Enregistrer dans un nouveau fichier Excel
merged_data.to_excel("Po_Conversion.xlsx", index=False)

print("Les données fusionnées ont été enregistrées dans 'Argent_Conversion.xlsx'.")