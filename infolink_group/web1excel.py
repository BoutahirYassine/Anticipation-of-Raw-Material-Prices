import pandas as pd
from datetime import datetime

current_date = datetime.now().strftime("%Y-%m-%d")
filename = f'files/infolink_{current_date}.xlsx'

def to_excel_combined(values):
    all_data = []

    sections = {
        'Polysilicon': {
            'Item': [
                'Polysilicon Price - Chunk polysilicon (USD)',
                'Polysilicon Price - Chunk polysilicon (RMB)',
                'Polysilicon Price - Granular polysilicon (RMB)'
            ],
            'High': [values[0], values[5], values[10]],
            'Low': [values[1], values[6], values[11]],
            'Average price': [values[2], values[7], values[12]],
            'Change(%)': [values[3], values[8], values[13]],
            'Change($)': [values[4], values[9], values[14]],
            'Price prediction for next week': [None, None, None]
        },
        'P Type Wafer': {
            'Item': [
                'Mono P Type Wafer - 182-183.75mm / 150µm (USD)',
                'Mono P Type Wafer - 182-183.75mm / 150µm (RMB)'
            ],
            'High': [values[15], values[20]],
            'Low': [values[16], values[21]],
            'Average price': [values[17], values[22]],
            'Change(%)': [values[18], values[23]],
            'Change($)': [values[19], values[24]],
            'Price prediction for next week': [None, None]
        },
        'N Type Wafer': {
            'Item': [
                'Mono N Type Wafer - 182-183.75mm / 130µm (USD)',
                'Mono N Type Wafer - 182-183.75mm / 130µm (RMB)',
                'Mono N Type Wafer - 182*210mm / 130µm (RMB)',
                'Mono N Type Wafer - 210mm / 130µm (RMB)'
            ],
            'High': [values[35], values[40], values[45], values[50]],
            'Low': [values[36], values[41], values[46], values[51]],
            'Average price': [values[37], values[42], values[47], values[52]],
            'Change(%)': [values[38], values[43], values[48], values[53]],
            'Change($)': [values[39], values[44], values[49], values[54]],
            'Price prediction for next week': [None, None, None, None]
        },
        'P Type Cell': {
            'Item': [
                'Mono PERC Cell - 182-183.75mm / 23.1%+ (USD)',
                'Mono PERC Cell - 182-183.75mm / 23.1%+ (RMB)'
            ],
            'High': [values[55], values[60]],
            'Low': [values[56], values[61]],
            'Average price': [values[57], values[62]],
            'Change(%)': [values[58], values[63]],
            'Change($)': [values[59], values[64]],
            'Price prediction for next week': [None, None]
        },
        'N Type Cell': {
            'Item': [
                'TOPCon Cell - 182-183.75mm / 24.9%+ (USD)',
                'TOPCon Cell - 182-183.75mm / 24.9%+ (RMB)',
                'TOPCon Cell - 182*210mm / 24.9%+ (RMB)',
                'TOPCon Cell - 210mm / 24.9%+ (RMB)'
            ],
            'High': [values[65], values[70], values[75], values[80]],
            'Low': [values[66], values[71], values[76], values[81]],
            'Average price': [values[67], values[72], values[77], values[82]],
            'Change(%)': [values[68], values[73], values[78], values[83]],
            'Change($)': [values[69], values[74], values[79], values[84]],
            'Price prediction for next week': [None, None, None, None]
        },
        'Glass Bifacial PERC Module': {
            'Item': [
                '182*182-210mm  Mono PERC Module (USD)',
                '182*182-210mm  Mono PERC Module (RMB)'
            ],
            'High': [values[85], values[90]],
            'Low': [values[86], values[91]],
            'Average price': [values[87], values[92]],
            'Change(%)': [values[88], values[93]],
            'Change($)': [values[89], values[94]],
            'Price prediction for next week': [None, None]
        },
        'Region Module': {
            'Item': [
                '182*182-210mm Mono PERC - India made (USD)',
                '182*182-210mm/210mm Mono TOPCon - India (USD)',
                '182*182-210mm/210mm Mono PERC - US (USD)',
                '182*182-210mm/210mm Mono TOPCon - US (USD)',
                '182*182-210mm/210mm Mono PERC - EU (USD)',
                '182*182-210mm/210mm Mono TOPCon - EU (USD)'
            ],
            'High': [values[95], values[100], values[105], values[110], values[115], values[120]],
            'Low': [values[96], values[101], values[106], values[111], values[116], values[121]],
            'Average price': [values[97], values[102], values[107], values[112], values[117], values[122]],
            'Change(%)': [values[98], values[103], values[108], values[113], values[118], values[123]],
            'Change($)': [values[99], values[104], values[109], values[114], values[119], values[124]],
            'Price prediction for next week': [None] * 6
        }
    }

    for section_name, data in sections.items():
        df = pd.DataFrame(data)
        df['Section'] = section_name
        all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='All Data', index=False)

    print("SAVED 1")
