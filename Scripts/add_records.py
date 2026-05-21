import os
import pandas as pd

# Ruta de la carpeta con los archivos de Excel de origen
carpeta_origen = './Contactabilidad/Data/campanas enero'  # Cambia esto a la ruta de tu carpeta
# Ruta del archivo de Excel principal donde se agregarán los datos
archivo_destino = './Data/GlobalDataUpdated-17-2-2025.xlsx'  # Cambia esto a la ruta de tu archivo principal

# Leer el archivo principal o crear un DataFrame vacío si no existe
if os.path.exists(archivo_destino):
    df_destino = pd.read_excel(archivo_destino)
else:
    df_destino = pd.DataFrame()

# Recorrer todos los archivos en la carpeta de origen
for archivo in os.listdir(carpeta_origen):
    if archivo.endswith('.xlsx'):
        # Leer cada archivo de Excel
        ruta_archivo = os.path.join(carpeta_origen, archivo)
        df_origen = pd.read_excel(ruta_archivo)
        
        # Ajustar las columnas del archivo de origen para que coincidan con las del archivo de destino
        columnas_comunes = df_destino.columns.intersection(df_origen.columns)
        df_origen = df_origen[columnas_comunes]  # Seleccionar solo las columnas comunes
        df_origen = df_origen.reindex(columns=df_destino.columns)  # Reordenar las columnas en el mismo orden

        # Agregar los datos del archivo al archivo principal
        df_destino = pd.concat([df_destino, df_origen], ignore_index=True)

# Guardar los datos actualizados en el archivo de destino
df_destino.to_excel(archivo_destino, index=False)
print("Todos los datos han sido agregados al archivo de destino.")

