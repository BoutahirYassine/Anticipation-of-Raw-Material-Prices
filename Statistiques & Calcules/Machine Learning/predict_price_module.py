# -*- coding: utf-8 -*-
"""
Created on Mon Jul 21 16:35:39 2025

@author: boutahirya
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jul 11 11:50:27 2025

@author: boutahirya
"""

import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from sklearn.svm import SVR
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error

# Chargement
df = pd.read_excel("DATA_solar_module.xlsx")  # adapte si CSV

# Colonnes à convertir (virgules → points)
colonnes_a_convertir = ['3-Module_USD', 'N-Type', 'N-type Wafer','3-Cell_RMB',
                        'Silver Rear_side', 'Silver Busbar front-side', 'Silver finger front-side']
for col in colonnes_a_convertir:
    df[col] = df[col].astype(str).str.replace(',', '.').astype(float)

# Création de la cible à J+12
df['Module_Price_t_plus_15'] = df['3-Module_USD'].shift(-30)


# Suppression des lignes incomplètes
df = df.dropna()

# Variables d'entrée
X = df[['N-Type', 'N-type Wafer','3-Cell_RMB', 'Silver Rear_side',
        'Silver Busbar front-side', 'Silver finger front-side']]
y = df['Module_Price_t_plus_15']

# Split train/test temporel (80/20)
split = int(0.8 * len(df))
X_train, X_test = X.iloc[:split], X.iloc[split:]
y_train, y_test = y.iloc[:split], y.iloc[split:]

# Normalisation pour SVR
scaler = StandardScaler()
X_train_svr = scaler.fit_transform(X_train)
X_test_svr = scaler.transform(X_test)

# Modèles
rf = RandomForestRegressor(n_estimators=100, random_state=42)
xgb = XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42)
svr = SVR(kernel='rbf', C=10, epsilon=0.01)

# Entraînement
rf.fit(X_train, y_train)
xgb.fit(X_train, y_train)
svr.fit(X_train_svr, y_train)

# Évaluation
print("Évaluation des modèles :")
for model, name, xtest in zip([rf, xgb, svr],
                               ['Random Forest', 'XGBoost', 'SVR'],
                               [X_test, X_test, X_test_svr]):
    y_pred = model.predict(xtest)
    mae = mean_absolute_error(y_test, y_pred)
    rmse = mean_squared_error(y_test, y_pred, squared=False)
    print(f"{name} → MAE : {mae:.4f} | RMSE : {rmse:.4f}")
