import json
import pandas as pd
import os

# Ruta de la carpeta con JSONs
folder_path = './DataNasdaq100JSON'

all_data = []

# Recorrer los archivos de la carpeta
for file_name in os.listdir(folder_path):
    if file_name.endswith('.json'):  # Filtrar solo archivos JSON
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
            all_data.append(data)
            #print(f"Contenido de {file_name}:", data)
            
            
#Convert JSON data to pandas dataframe
df = pd.DataFrame(all_data)
df.to_excel('NASDAQ100.xlsx', index=False)