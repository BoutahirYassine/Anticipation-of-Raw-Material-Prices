# -*- coding: utf-8 -*-
"""
Created on Thu Mar 27 11:13:35 2025

@author: boutahirya
"""

import pandas as pd

# Charger le fichier Excel filtré sur la source EnergyTrend
file_path = "Filtered_EnergyTrend_Data.xlsx"
df = pd.read_excel(file_path)

# Filtrer les lignes correspondant aux deux descriptions d'intérêt
df_filtered = df[
    df['Description'].str.contains('n-type TOPCon M10', case=False, na=False) |
    ((df['Description'].str.contains('n-type TOPCon', case=False, na=False)) & (df['Item'] == 'Module'))
]

# Créer un tableau avec une ligne par date et une colonne par type de produit
df_pivot = df_filtered.pivot_table(
    index='Date',
    columns='Description',
    values='Base Price',
    aggfunc='first'
).reset_index()

# Renommer les colonnes pour plus de clarté
df_pivot.columns.name = None
df_pivot = df_pivot.rename(columns={
    'n-type TOPCon M10': 'Prix Cellule n-type TOPCon M10',
    'n-type TOPCon': 'Prix Module n-type TOPCon'
})

# Enregistrer le résultat dans un nouveau fichier Excel
output_file = "Prix_Cellule_Module_TOPCon.xlsx"
df_pivot.to_excel(output_file, index=False)

print(f"Fichier enregistré sous : {output_file}")
