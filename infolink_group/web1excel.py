import pandas as pd
from datetime import datetime
import os

# Obtenir la date actuelle sous forme de chaîne au format "YYYY-MM-DD"
current_date = datetime.now().strftime("%Y-%m-%d")

# Nom du fichier avec la date
filename = f'files/infolink_{current_date}.xlsx'

def to_excel_combined(values):
    # Créer une liste pour stocker tous les DataFrames
    all_data = []

    # Définir chaque section avec un nom et les valeurs
    sections = {
        'Polysilicon': {
            'Item': ['Polysilicon Price - Chunk polysilicon (USD)', 'Polysilicon Price - Chunk polysilicon (RMB)', 'Polysilicon Price - Granular polysilicon (RMB)'],
            'High': [values[0], values[5], values[10]],
            'Low': [values[1], values[6], values[11]],
            'Average price': [values[2], values[7], values[12]],
            'Change(%)': [values[3], values[8], values[13]],
            'Change($)': [values[4], values[9], values[14]],
            'Price prediction for next week': [None, None, None]
        },
        'P Type Wafer': {
            'Item': ['Mono P Type Wafer - 182-183.75mm / 150µm (USD)', 'Mono P Type Wafer - 182-183.75mm / 150µm (RMB)', 'Mono P Type Wafer - 210mm / 150µm (USD)', 'Mono P Type Wafer - 210mm / 150µm (RMB)'],
            'High': [values[15], values[20], values[25], values[30]],
            'Low': [values[16], values[21], values[26], values[31]],
            'Average price': [values[17], values[22], values[27], values[32]],
            'Change(%)': [values[18], values[23], values[28], values[33]],
            'Change($)': [values[19], values[24], values[29], values[34]],
            'Price prediction for next week': [None, None, None, None]
        },
        'N Type Wafer': {
            'Item': ['Mono N Type Wafer - 182-183.75mm / 130µm (RMB)', 'Mono N Type Wafer - 182*210mm / 130µm (RMB)', 'Mono N Type Wafer - 210mm / 130µm (RMB)'],
            'High': [values[35], values[40], values[45]],
            'Low': [values[36], values[41], values[46]],
            'Average price': [values[37], values[42], values[47]],
            'Change(%)': [values[38], values[43], values[48]],
            'Change($)': [values[39], values[44], values[49]],
            'Price prediction for next week': [None, None, None]
        },
        'P Type Cell': {
            'Item': ['Mono PERC Cell - 182-183.75mm / 23.1%+ (USD)', 'Mono PERC Cell - 182-183.75mm / 23.1%+ (RMB)', 'Mono PERC Cell - 210mm / 23.1%+ (USD)', 'Mono PERC Cell - 210mm / 23.1%+ (RMB)'],
            'High': [values[50], values[55], values[60], values[65]],
            'Low': [values[51], values[56], values[61], values[66]],
            'Average price': [values[52], values[57], values[62], values[67]],
            'Change(%)': [values[53], values[58], values[63], values[68]],
            'Change($)': [values[54], values[59], values[64], values[69]],
            'Price prediction for next week': [None, None, None, None]
        },
        'N Type Cell': {
            'Item': ['TOPCon Cell - 182-183.75mm / 24.9%+ (USD)', 'TOPCon Cell - 182-183.75mm / 24.9%+ (RMB)', 'TOPCon Cell - 182*210mm / 24.9%+ (RMB)', 'TOPCon Cell - 210mm / 24.9%+ (RMB)'],
            'High': [values[70], values[75], values[80], values[85]],
            'Low': [values[71], values[76], values[81], values[86]],
            'Average price': [values[72], values[77], values[82], values[87]],
            'Change(%)': [values[73], values[78], values[83], values[88]],
            'Change($)': [values[74], values[79], values[84], values[89]],
            'Price prediction for next week': [None, None, None, None]
        },
        'Glass Bifacial PERC': {
            'Item': ['182*182-210mm Mono PERC Module (USD)', '182*182-210mm Mono PERC Module (RMB)', '210mm Mono PERC Module (USD)', '210mm Mono PERC Module (RMB)'],
            'High': [values[90], values[95], values[100], values[105]],
            'Low': [values[91], values[96], values[101], values[106]],
            'Average price': [values[92], values[97], values[102], values[107]],
            'Change(%)': [values[93], values[98], values[103], values[108]],
            'Change($)': [values[94], values[99], values[104], values[109]],
            'Price prediction for next week': [None, None, None, None]
        },
        'Glass Bifacial NType': {
            'Item': ['182*182-210mm Mono TOPCon Module (USD)', '182*182-210mm Mono TOPCon Module (RMB)', '210mm Mono HJT Module (USD)', '210mm Mono HJT Module (RMB)'],
            'High': [values[110], values[115], values[120], values[125]],
            'Low': [values[111], values[116], values[121], values[126]],
            'Average price': [values[112], values[117], values[122], values[127]],
            'Change(%)': [values[113], values[118], values[123], values[128]],
            'Change($)': [values[114], values[119], values[124], values[129]],
            'Price prediction for next week': [None, None, None, None]
        },
        'China Projects': {
            'Item': ['182*182-210mm/210mm TOPCon Module - Ground-mounted project', '182*182-210mm/210mm TOPCon Module - Distributed project'],
            'High': [values[130], values[135]],
            'Low': [values[131], values[136]],
            'Average price': [values[132], values[137]],
            'Change(%)': [values[133], values[138]],
            'Change($)': [values[134], values[139]],
            'Price prediction for next week': [None, None]
        },
        'Region Module': {
            'Item': ['182*182-210mm Mono PERC - India made (USD)', '182*182-210mm/210mm Mono TOPCon - India (USD)', '182*182-210mm/210mm Mono PERC - US (USD)', '182*182-210mm/210mm Mono TOPCon - US (USD)', '182*182-210mm/210mm Mono PERC - EU (USD)', '182*182-210mm/210mm Mono TOPCon - EU (USD)'],
            'High': [values[140], values[145], values[150], values[155], values[160], values[165]],
            'Low': [values[141], values[146], values[151], values[156], values[161], values[166]],
            'Average price': [values[142], values[147], values[152], values[157], values[162], values[167]],
            'Change(%)': [values[143], values[148], values[153], values[158], values[163], values[168]],
            'Change($)': [values[144], values[149], values[154], values[159], values[164], values[169]],
            'Price prediction for next week': [None] * 6
        },
        'Module BOM Materials': {
            'Item': ['3.2mm Coating PV Glass(RMB)', '2.0mm Coating PV Glass(RMB)'],
            'High': [values[170], values[175]],
            'Low': [values[171], values[176]],
            'Average price': [values[172], values[177]],
            'Change(%)': [values[173], values[178]],
            'Change($)': [values[174], values[179]],
            'Price prediction for next week': [None, None]
        }
    }

    # Convertir chaque section en DataFrame et ajouter une colonne pour le nom de la section
    for section_name, data in sections.items():
        df = pd.DataFrame(data)
        df['Section'] = section_name
        all_data.append(df)

    # Combiner tous les DataFrames en un seul
    combined_df = pd.concat(all_data, ignore_index=True)

    # Enregistrer dans un seul fichier Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='All Data', index=False)

    print("SAVED 1")

# Appel de la fonction avec vos données
# Exemple : to_excel_combined(values)
