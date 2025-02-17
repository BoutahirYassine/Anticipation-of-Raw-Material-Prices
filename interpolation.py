# -*- coding: utf-8 -*-
"""
Created on Mon Feb 17 11:58:18 2025

@author: boutahirya
"""

import pandas as pd
import numpy as np
from scipy import interpolate
import matplotlib.pyplot as plt

# Charger le fichier Excel
def load_excel(file_path):
    return pd.read_excel(file_path)

# Interpolation des données
def interpolate_data(df):
    # Convertir la colonne 'Date' en format datetime
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Créer un nouvel index avec toutes les dates du premier au dernier jour
    full_dates = pd.date_range(start=df['Date'].min(), end=df['Date'].max(), freq='D')
    
    # Interpolation des valeurs pour chaque jour
    interp_func = interpolate.interp1d(pd.to_numeric(df['Date']), df['Value'], kind='linear', fill_value="extrapolate")
    interpolated_values = interp_func(pd.to_numeric(full_dates))
    
    # Créer un nouveau DataFrame avec les dates et les valeurs interpolées
    result_df = pd.DataFrame({
        'Date': full_dates,
        'Interpolated Value': interpolated_values
    })
    
    return result_df

# Sauvegarder le DataFrame interpolé dans un nouveau fichier Excel
def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)

# Exemple d'utilisation
file_path = 'DataBI/wafer_m10.xlsx'  # Remplacez par le chemin de votre fichier
output_path = 'output_interpolé.xlsx'  # Chemin du fichier de sortie

# Charger les données depuis le fichier Excel
df = load_excel(file_path)

# Appliquer l'interpolation
interpolated_df = interpolate_data(df)

# Sauvegarder le résultat dans un nouveau fichier Excel
save_to_excel(interpolated_df, output_path)

# Afficher quelques résultats pour vérifier
print(interpolated_df.head())

