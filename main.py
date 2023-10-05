import pandas as pd
from sqlalchemy import create_engine
import os

# Nombre del archivo Excel
nombre_archivo = 'PMI_Spiral Histórico (Rappi).xlsx'

# Ruta completa del archivo Excel dentro de la carpeta fileAProcesar
ruta_excel = os.path.join(os.getcwd(), 'fileToProcess', nombre_archivo)

# Crear un diccionario de mapeo entre los nombres de columnas del Excel y los de la tabla SQL
mapeo_columnas = {
    'Retailer': 'Retailer',
    'Tienda': 'Tienda',
    'IDSalesforce': 'IDSalesforce',
    'Dirección': 'Direccion',
    'Zonas': 'Zonas',
    'SKU': 'SKU',
    'Marca': 'Marca',
    'Empresa': 'Empresa',
    'Nombre PMI': 'NombrePMI',
    'Imagén PMI': 'ImagenPMI',
    'Nombre en tienda': 'NombreEnTienda',
    'Imagen tienda': 'ImagenTienda',
    'Precio': 'Precio',
    'PVP': 'PVP',
    'Respeta precio': 'RespetaPrecio',
    'Respeta imagen': 'RespetaImagen',
    'Respeta nombre': 'RespetaNombre',
    'Fecha': 'Fecha',
    'posis_code': 'PosisCode'
}

# Leer el archivo Excel en un DataFrame de pandas
df_excel = pd.read_excel(ruta_excel)

# Renombrar las columnas del DataFrame de acuerdo al mapeo
df_excel.rename(columns=mapeo_columnas, inplace=True)

engine = create_engine('mssql+pyodbc://PMIAROLIAPP01/InfoRP?trusted_connection=yes&driver=SQL+Server',fast_executemany=True)

# Insertar los datos del DataFrame en la tabla de la base de datos
try:
    df_excel.to_sql('Data', engine, if_exists='append', index=False)
    print("Datos insertados correctamente.")
except Exception as ex:
    print(f"Error en la conexión o inserción: {ex}")
