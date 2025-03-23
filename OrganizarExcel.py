import pandas as pd

def modificar_excel(ruta_entrada, ruta_salida):
    # Cargar el archivo Excel
    xls = pd.ExcelFile(ruta_entrada)
    
    # Verificar si las hojas 'Contrato' y 'Tipo_de_inmueble' existen
    if 'Contrato' in xls.sheet_names and 'Tipo_de_inmueble' in xls.sheet_names:
        df_contrato = pd.read_excel(xls, 'Contrato')
        df_tipo_inmueble = pd.read_excel(xls, 'Tipo_de_inmueble')
    else:
        print("Error: Las hojas 'Contrato' o 'Tipo_de_inmueble' no se encuentran en el archivo Excel.")
        print("Hojas disponibles en el archivo Excel:", xls.sheet_names)
        return

    # Verificar los nombres de las columnas en df_contrato y df_tipo_inmueble
    print("Columnas en df_contrato:", df_contrato.columns)
    print("Columnas en df_tipo_inmueble:", df_tipo_inmueble.columns)

    # Mapear NOMBRE_INMUEBLE a su nuevo id
    tipo_inmueble_dict = df_tipo_inmueble.set_index('NOMBRE_INMUEBLE')['ID_INMUEBLE'].to_dict()
    df_contrato['ID_INMUEBLE'] = df_contrato['ID_INMUEBLE'].map(tipo_inmueble_dict)

    # Guardar el archivo modificado
    with pd.ExcelWriter(ruta_salida, engine='xlsxwriter') as writer:
        df_contrato.to_excel(writer, sheet_name='Contrato', index=False)
        df_tipo_inmueble.to_excel(writer, sheet_name='Tipo_de_inmueble', index=False)

    print("Modificación completada. Archivo guardado en:", ruta_salida)

# Rutas de los archivos
ruta_entrada = r"c:\Users\cdjua\OneDrive\Desktop\Prueba T\Data_limpia.xlsx"
ruta_salida = "Data_Modificada.xlsx"

# Ejecutar la función
modificar_excel(ruta_entrada, ruta_salida)
