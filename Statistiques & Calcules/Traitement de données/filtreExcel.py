import pandas as pd

# Charger le fichier Excel
file_path = "BI_Prices_Data.xlsx"  # Remplacer par le chemin vers ton fichier
df = pd.read_excel(file_path, sheet_name="Component Prices Data")

# Filtrer les lignes dont la source est 'EnergyTrend' (insensible à la casse)
df_filtered = df[df['Source'].str.lower() == 'energytrend']

# Enregistrer le résultat dans un nouveau fichier Excel
output_path = "Filtered_EnergyTrend_Data.xlsx"
df_filtered.to_excel(output_path, index=False)

print("Fichier filtré enregistré sous :", output_path)
