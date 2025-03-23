import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF

# Cargar el archivo de Excel
archivo = r"c:\Users\cdjua\OneDrive\Desktop\Prueba T\Data (1).xlsx"
if os.path.exists(archivo):
    print("Archivo encontrado, cargando datos...")
    df_contrato = pd.read_excel(archivo, sheet_name="Contrato")
    df_sectores = pd.read_excel(archivo, sheet_name="Sectores")
    df_clientes = pd.read_excel(archivo, sheet_name="Clientes")
else:
    print("Error: Archivo no encontrado, verifica la ruta.")
    exit()

# Revisar valores nulos
print("\nValores nulos en los datos de contratos:")
print(df_contrato.isnull().sum())

print("\nValores nulos en los datos de sectores:")
print(df_sectores.isnull().sum())

print("\nValores nulos en los datos de clientes:")
print(df_clientes.isnull().sum())

# Revisar duplicados
print("\nDuplicados en los datos de contratos:")
print(df_contrato.duplicated().sum())

print("\nDuplicados en los datos de sectores:")
print(df_sectores.duplicated().sum())

print("\nDuplicados en los datos de clientes:")
print(df_clientes.duplicated().sum())

# Revisar tipos de datos
print("\nTipos de datos en los datos de contratos:")
print(df_contrato.dtypes)

print("\nTipos de datos en los datos de sectores:")
print(df_sectores.dtypes)

print("\nTipos de datos en los datos de clientes:")
print(df_clientes.dtypes)

# Análisis exploratorio de datos
# Resumen de los datos
print("\nResumen de los datos de contratos:")
print(df_contrato.describe())

print("\nResumen de los datos de sectores:")
print(df_sectores.describe())

print("\nResumen de los datos de clientes:")
print(df_clientes.describe())

# Principales insights
print("\nPrincipales insights:")

# Tendencias identificadas
print("\nTendencias identificadas:")
# Verificar si la columna 'Canon' existe
if 'Canon' in df_contrato.columns:
    print("Distribución de precios de alquiler (Canon):")
    print(df_contrato['Canon'].describe())
else:
    print("La columna 'Canon' no existe en los datos de contratos.")

# Anomalías detectadas
print("\nAnomalías detectadas:")
# Verificar si la columna 'Canon' existe antes de buscar anomalías
if 'Canon' in df_contrato.columns:
    anomalies = df_contrato[(df_contrato['Canon'] < 100) | (df_contrato['Canon'] > 10000)]
    print("Contratos con precios anómalos:")
    print(anomalies)
else:
    anomalies = pd.DataFrame()
    print("No se pueden detectar anomalías porque la columna 'Canon' no existe en los datos de contratos.")

# Verificar inconsistencias en los datos de sectores
print("\nInconsistencias en los datos de sectores:")
inconsistencias_poblado = df_sectores[(df_sectores['Zona'] == 'El Poblado') & (df_sectores['Ciudad'] != 'Medellín')]
inconsistencias_laureles = df_sectores[(df_sectores['Zona'] == 'Laureles') & (df_sectores['Ciudad'] != 'Medellín')]
inconsistencias = pd.concat([inconsistencias_poblado, inconsistencias_laureles])
print(inconsistencias)

# Impacto en la estrategia de alquiler
print("\nImpacto en la estrategia de alquiler:")
# Ejemplo: Ajustes en la estrategia de precios
if not anomalies.empty:
    print("Se recomienda revisar los contratos con precios anómalos para ajustar la estrategia de precios.")

# Análisis adicional
# Tendencias de alquiler en los últimos años
print("\nTendencias de alquiler en los últimos años:")
if 'Fecha cierre' in df_contrato.columns:
    df_contrato['Fecha cierre'] = pd.to_datetime(df_contrato['Fecha cierre'])
    contratos_por_anio = df_contrato.groupby(df_contrato['Fecha cierre'].dt.year).size()
    print(contratos_por_anio)
else:
    print("La columna 'Fecha cierre' no existe en los datos de contratos.")

# Patrones en los precios de alquiler por zona y tipo de inmueble
print("\nPatrones en los precios de alquiler por zona y tipo de inmueble:")
if 'Zona' in df_sectores.columns and 'Canon' in df_contrato.columns:
    precios_por_zona = df_contrato.groupby(df_sectores['Zona'])['Canon'].describe()
    print(precios_por_zona)
else:
    print("Las columnas 'Zona' o 'Canon' no existen en los datos.")

# Sectores más rentables
print("\nSectores más rentables:")
if 'Ingresos mensuales' in df_clientes.columns and 'Zona' in df_sectores.columns:
    ingresos_por_zona = df_clientes.groupby(df_sectores['Zona'])['Ingresos mensuales'].sum()
    print(ingresos_por_zona)
else:
    print("Las columnas 'Ingresos mensuales' o 'Zona' no existen en los datos.")

# Impacto de los ingresos mensuales de los clientes en los precios que pagan
print("\nImpacto de los ingresos mensuales de los clientes en los precios que pagan:")
if 'Ingresos mensuales' in df_clientes.columns and 'Canon' in df_contrato.columns:
    correlacion = df_clientes['Ingresos mensuales'].corr(df_contrato['Canon'])
    print(f"Correlación entre ingresos y valor del alquiler: {correlacion}")
else:
    print("Las columnas 'Ingresos mensuales' o 'Canon' no existen en los datos.")

# Perfil de los clientes según edad e ingresos
print("\nPerfil de los clientes según edad e ingresos:")
if 'Edad' in df_clientes.columns and 'Ingresos mensuales' in df_clientes.columns:
    perfil_clientes = df_clientes.groupby(['Edad'])['Ingresos mensuales'].describe()
    print(perfil_clientes)
else:
    print("Las columnas 'Edad' o 'Ingresos mensuales' no existen en los datos.")

# Anomalías o valores extremos en los contratos de alquiler
print("\nAnomalías o valores extremos en los contratos de alquiler:")
if 'Canon' in df_contrato.columns:
    valores_extremos = df_contrato[(df_contrato['Canon'] < 100) | (df_contrato['Canon'] > 10000)]
    print(valores_extremos)
else:
    print("La columna 'Canon' no existe en los datos de contratos.")

# Ocupación de los inmuebles por sector y tipo de inmueble
print("\nOcupación de los inmuebles por sector y tipo de inmueble:")
if 'Zona' in df_sectores.columns and 'Tipo de Inmuble' in df_contrato.columns:
    ocupacion_por_zona = df_contrato.groupby([df_sectores['Zona'], 'Tipo de Inmuble']).size()
    print(ocupacion_por_zona)
else:
    print("Las columnas 'Zona' o 'Tipo de Inmuble' no existen en los datos.")

# Factores que afectan la duración de los contratos
print("\nFactores que afectan la duración de los contratos:")
if 'Duracion' in df_contrato.columns and 'Zona' in df_sectores.columns:
    duracion_por_zona = df_contrato.groupby(df_sectores['Zona'])['Duracion'].describe()
    print(duracion_por_zona)
else:
    print("Las columnas 'Duracion' o 'Zona' no existen en los datos.")

# Generar gráficos
# Histograma de precios de alquiler
if 'Canon' in df_contrato.columns:
    plt.figure(figsize=(10, 6))
    sns.histplot(df_contrato['Canon'], bins=30, kde=True)
    plt.title('Histograma de precios de alquiler (Canon)')
    plt.xlabel('Canon')
    plt.ylabel('Frecuencia')
    plt.savefig('histograma_precios_alquiler.png')
    plt.show()

# Gráfico de evolución del número de contratos
if 'Fecha cierre' in df_contrato.columns:
    plt.figure(figsize=(10, 6))
    contratos_por_anio.plot(kind='bar')
    plt.title('Evolución del número de contratos por año')
    plt.xlabel('Año')
    plt.ylabel('Número de contratos')
    plt.savefig('evolucion_contratos.png')
    plt.show()

# Mapa de calor de zonas con mayor cantidad de contratos
if 'Zona' in df_sectores.columns:
    contratos_por_zona = df_contrato.groupby(df_sectores['Zona']).size().reset_index(name='Cantidad')
    if not contratos_por_zona.empty:
        contratos_por_zona_pivot = contratos_por_zona.pivot(index='Zona', columns='Cantidad', values='Cantidad')
        if not contratos_por_zona_pivot.empty:
            plt.figure(figsize=(10, 6))
            sns.heatmap(contratos_por_zona_pivot, annot=True, fmt=".2f", cmap="YlGnBu")
            plt.title('Mapa de calor de zonas con mayor cantidad de contratos')
            plt.xlabel('Cantidad')
            plt.ylabel('Zona')
            plt.savefig('mapa_calor_zonas.png')
            plt.show()
        else:
            print("El DataFrame pivotado está vacío, no se puede generar el mapa de calor.")
    else:
        print("No hay datos suficientes para generar el mapa de calor.")
else:
    print("La columna 'Zona' no existe en los datos de sectores.")

# Matriz de correlación
print("\nMatriz de correlación:")
if 'Canon' in df_contrato.columns and 'Ingresos mensuales' in df_clientes.columns and 'Edad' in df_clientes.columns:
    df_merged = df_contrato.merge(df_clientes, on='IdCliente')
    correlacion_matriz = df_merged[['Canon', 'Ingresos mensuales', 'Edad']].corr()
    plt.figure(figsize=(10, 6))
    sns.heatmap(correlacion_matriz, annot=True, cmap="coolwarm")
    plt.title('Matriz de correlación')
    plt.savefig('matriz_correlacion.png')
    plt.show()
    print(correlacion_matriz)
else:
    print("No se pueden generar correlaciones porque faltan columnas necesarias.")

# Generar un informe en un archivo de texto
with open("informe_exploratorio.txt", "w") as file:
    file.write("Informe Exploratorio de Datos\n")
    file.write("============================\n\n")
    file.write("Resumen de los datos de contratos:\n")
    file.write(df_contrato.describe().to_string())
    file.write("\n\nResumen de los datos de sectores:\n")
    file.write(df_sectores.describe().to_string())
    file.write("\n\nResumen de los datos de clientes:\n")
    file.write(df_clientes.describe().to_string())
    file.write("\n\nPrincipales insights:\n")
    # Añadir insights específicos aquí
    file.write("\n\nTendencias identificadas:\n")
    if 'Canon' in df_contrato.columns:
        file.write("Distribución de precios de alquiler (Canon):\n")
        file.write(df_contrato['Canon'].describe().to_string())
    else:
        file.write("La columna 'Canon' no existe en los datos de contratos.\n")
    file.write("\n\nAnomalías detectadas:\n")
    if not anomalies.empty:
        file.write("Contratos con precios anómalos:\n")
        file.write(anomalies.to_string())
    else:
        file.write("No se pueden detectar anomalías porque la columna 'Canon' no existe en los datos de contratos.\n")
    file.write("\n\nInconsistencias en los datos de sectores:\n")
    file.write(inconsistencias.to_string())
    file.write("\n\nImpacto en la estrategia de alquiler:\n")
    if not anomalies.empty:
        file.write("Se recomienda revisar los contratos con precios anómalos para ajustar la estrategia de precios.\n")
    # Añadir análisis adicional al informe
    file.write("\n\nTendencias de alquiler en los últimos años:\n")
    if 'Fecha cierre' in df_contrato.columns:
        file.write(contratos_por_anio.to_string())
    else:
        file.write("La columna 'Fecha cierre' no existe en los datos de contratos.\n")
    file.write("\n\nPatrones en los precios de alquiler por zona y tipo de inmueble:\n")
    if 'Zona' in df_sectores.columns and 'Canon' in df_contrato.columns:
        file.write(precios_por_zona.to_string())
    else:
        file.write("Las columnas 'Zona' o 'Canon' no existen en los datos.\n")
    file.write("\n\nSectores más rentables:\n")
    if 'Ingresos mensuales' in df_clientes.columns and 'Zona' in df_sectores.columns:
        file.write(ingresos_por_zona.to_string())
    else:
        file.write("Las columnas 'Ingresos mensuales' o 'Zona' no existen en los datos.\n")
    file.write("\n\nImpacto de los ingresos mensuales de los clientes en los precios que pagan:\n")
    if 'Ingresos mensuales' in df_clientes.columns and 'Canon' in df_contrato.columns:
        file.write(f"Correlación entre ingresos y valor del alquiler: {correlacion}\n")
    else:
        file.write("Las columnas 'Ingresos mensuales' o 'Canon' no existen en los datos.\n")
    file.write("\n\nPerfil de los clientes según edad e ingresos:\n")
    if 'Edad' in df_clientes.columns and 'Ingresos mensuales' in df_clientes.columns:
        file.write(perfil_clientes.to_string())
    else:
        file.write("Las columnas 'Edad' o 'Ingresos mensuales' no existen en los datos.\n")
    file.write("\n\nAnomalías o valores extremos en los contratos de alquiler:\n")
    if 'Canon' in df_contrato.columns:
        file.write(valores_extremos.to_string())
    else:
        file.write("La columna 'Canon' no existe en los datos de contratos.\n")
    file.write("\n\nOcupación de los inmuebles por sector y tipo de inmueble:\n")
    if 'Zona' in df_sectores.columns and 'Tipo de Inmuble' in df_contrato.columns:
        file.write(ocupacion_por_zona.to_string())
    else:
        file.write("Las columnas 'Zona' o 'Tipo de Inmuble' no existen en los datos.\n")
    file.write("\n\nFactores que afectan la duración de los contratos:\n")
    if 'Duracion' in df_contrato.columns and 'Zona' in df_sectores.columns:
        file.write(duracion_por_zona.to_string())
    else:
        file.write("Las columnas 'Duracion' o 'Zona' no existen en los datos.\n")

# Crear una clase para el PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Informe Exploratorio de Datos', 0, 1, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()

# Crear el PDF
pdf = PDF()
pdf.add_page()

# Añadir contenido al PDF
pdf.chapter_title('Resumen de los datos de contratos:')
pdf.chapter_body(df_contrato.describe().to_string())

pdf.chapter_title('Resumen de los datos de sectores:')
pdf.chapter_body(df_sectores.describe().to_string())

pdf.chapter_title('Resumen de los datos de clientes:')
pdf.chapter_body(df_clientes.describe().to_string())

pdf.chapter_title('Principales insights:')
# Añadir insights específicos aquí

pdf.chapter_title('Tendencias identificadas:')
if 'Canon' in df_contrato.columns:
    pdf.chapter_body("Distribución de precios de alquiler (Canon):\n" + df_contrato['Canon'].describe().to_string())
else:
    pdf.chapter_body("La columna 'Canon' no existe en los datos de contratos.")

pdf.chapter_title('Anomalías detectadas:')
if not anomalies.empty:
    pdf.chapter_body("Contratos con precios anómalos:\n" + anomalies.to_string())
else:
    pdf.chapter_body("No se pueden detectar anomalías porque la columna 'Canon' no existe en los datos de contratos.")

pdf.chapter_title('Inconsistencias en los datos de sectores:')
pdf.chapter_body(inconsistencias.to_string())

pdf.chapter_title('Impacto en la estrategia de alquiler:')
if not anomalies.empty:
    pdf.chapter_body("Se recomienda revisar los contratos con precios anómalos para ajustar la estrategia de precios.")

pdf.chapter_title('Tendencias de alquiler en los últimos años:')
if 'Fecha cierre' in df_contrato.columns:
    pdf.chapter_body(contratos_por_anio.to_string())
else:
    pdf.chapter_body("La columna 'Fecha cierre' no existe en los datos de contratos.")

pdf.chapter_title('Patrones en los precios de alquiler por zona y tipo de inmueble:')
if 'Zona' in df_sectores.columns and 'Canon' in df_contrato.columns:
    pdf.chapter_body(precios_por_zona.to_string())
else:
    pdf.chapter_body("Las columnas 'Zona' o 'Canon' no existen en los datos.")

pdf.chapter_title('Sectores más rentables:')
if 'Ingresos mensuales' in df_clientes.columns and 'Zona' in df_sectores.columns:
    pdf.chapter_body(ingresos_por_zona.to_string())
else:
    pdf.chapter_body("Las columnas 'Ingresos mensuales' o 'Zona' no existen en los datos.")

pdf.chapter_title('Impacto de los ingresos mensuales de los clientes en los precios que pagan:')
if 'Ingresos mensuales' in df_clientes.columns and 'Canon' in df_contrato.columns:
    pdf.chapter_body(f"Correlación entre ingresos y valor del alquiler: {correlacion}")
else:
    pdf.chapter_body("Las columnas 'Ingresos mensuales' o 'Canon' no existen en los datos.")

pdf.chapter_title('Perfil de los clientes según edad e ingresos:')
if 'Edad' in df_clientes.columns and 'Ingresos mensuales' in df_clientes.columns:
    pdf.chapter_body(perfil_clientes.to_string())
else:
    pdf.chapter_body("Las columnas 'Edad' o 'Ingresos mensuales' no existen en los datos.")

pdf.chapter_title('Anomalías o valores extremos en los contratos de alquiler:')
if 'Canon' in df_contrato.columns:
    pdf.chapter_body(valores_extremos.to_string())
else:
    pdf.chapter_body("La columna 'Canon' no existe en los datos de contratos.")

pdf.chapter_title('Ocupación de los inmuebles por sector y tipo de inmueble:')
if 'Zona' in df_sectores.columns and 'Tipo de Inmuble' in df_contrato.columns:
    pdf.chapter_body(ocupacion_por_zona.to_string())
else:
    pdf.chapter_body("Las columnas 'Zona' o 'Tipo de Inmuble' no existen en los datos.")

pdf.chapter_title('Factores que afectan la duración de los contratos:')
if 'Duracion' in df_contrato.columns and 'Zona' in df_sectores.columns:
    pdf.chapter_body(duracion_por_zona.to_string())
else:
    pdf.chapter_body("Las columnas 'Duracion' o 'Zona' no existen en los datos.")

# Guardar el PDF
pdf.output('informe_exploratorio.pdf')

