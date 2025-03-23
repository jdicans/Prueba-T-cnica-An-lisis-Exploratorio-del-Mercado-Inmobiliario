import pandas as pd
import sqlite3

# Ruta del archivo normalizado
archivo = r"c:\Users\cdjua\OneDrive\Desktop\Prueba T\Data_Limpia.xlsx"

# Leer todas las hojas del archivo
sheets = pd.read_excel(archivo, sheet_name=None)

# Conectar a la base de datos SQLite
conexion = sqlite3.connect(r"C:\Users\cdjua\OneDrive\Desktop\Prueba T\Inmobiliaria\inmobiliaria.db")
cursor = conexion.cursor()

# Eliminar las tablas existentes antes de insertar los datos nuevos
for nombre_hoja in sheets.keys():
    cursor.execute(f"DROP TABLE IF EXISTS {nombre_hoja}")

# Insertar cada hoja en la base de datos
for nombre_hoja, df in sheets.items():
    print(f"📥 Insertando '{nombre_hoja}' en SQLite...")
    df.to_sql(nombre_hoja, conexion, if_exists="replace", index=False)

# Cerrar conexión
conexion.close()
print("Datos exportados a SQLite correctamente.")
