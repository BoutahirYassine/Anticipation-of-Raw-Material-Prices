import pandas as pd

# Charger les fichiers Excel
argent_file = "DataBI/Conversionn.xlsx"
conversion_file = "DataBI/EUR_USD.xlsx"

argent_data = pd.read_excel(argent_file)
conversion_data = pd.read_excel(conversion_file)

# Renommer les colonnes pour éviter les problèmes d'espaces ou caractères spéciaux
argent_data.columns = ["date", "Silver price in grams (dollar)"]
conversion_data.columns = ["date", "Conversion rate (dollar to euro)"]

# Convertir les colonnes 'date' en type datetime pour faciliter le croisement et ne garder que la date courte
argent_data['date'] = pd.to_datetime(argent_data['date'], format="%d/%m/%Y", errors='coerce').dt.normalize()
conversion_data['date'] = pd.to_datetime(conversion_data['date'], format="%d/%m/%Y", errors='coerce').dt.normalize()

# Retirer les lignes avec des dates invalides si des erreurs sont survenues
argent_data = argent_data.dropna(subset=['date'])
conversion_data = conversion_data.dropna(subset=['date'])

# Joindre les deux DataFrames sur la colonne 'date'
merged_data = pd.merge(argent_data, conversion_data, on="date", how="inner")

# Enregistrer dans un nouveau fichier Excel
output_file = "DataBI/Conversionn.xlsx"
merged_data.to_excel(output_file, index=False)

print(f"Les données fusionnées ont été enregistrées dans '{output_file}'.")