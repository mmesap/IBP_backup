import pandas as pd
from utils_normalize import normalizar_pais, normalizar_rango_empleados, normalizar_industria, normalice_company_name


archivo_destino = './Data/GlobalDataUpdated (version 2) 05-11-2024.xlsx'  # Cambia esto a la ruta de tu archivo principal
df = pd.read_excel(archivo_destino)

# Aplicar la función de normalización a la columna
df['NÚMERO DE EMPLEADOS'] = df['NÚMERO DE EMPLEADOS'].apply(normalizar_rango_empleados)
df['PAÍS'] = df['PAÍS'].apply(normalizar_pais)
df['INDUSTRIA'] = df['INDUSTRIA'].apply(normalizar_industria)
df['NOMBRE COMERCIAL EMPRESA'] = df['NOMBRE COMERCIAL EMPRESA'].apply(normalice_company_name)

# Guardar el archivo de Excel con la columna normalizada
df.to_excel(archivo_destino, index=False)
print("Listo Normalizado")
