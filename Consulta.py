import sqlite3
import pandas as pd
from fpdf import FPDF

# Conectar a la base de datos
db_path = r"c:\Users\cdjua\OneDrive\Desktop\Prueba T\Inmobiliaria\inmobiliaria.db"  # Ruta del archivo SQLite
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Crear una clase para el PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Resultados de Consultas', 0, 1, 'C')

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

# Función para ejecutar consultas y mostrar resultados en DataFrame
def ejecutar_consulta(query, mensaje):
    try:
        df = pd.read_sql_query(query, conn)
        print(f"\n🔹 {mensaje}")
        print(df)
        pdf.chapter_title(mensaje)
        pdf.chapter_body(df.to_string(index=False))
    except Exception as e:
        print(f"Error en la consulta: {e}")

# Verificar datos en la tabla clientes
ejecutar_consulta("""
    SELECT * FROM clientes LIMIT 5;
""", "Datos de ejemplo en la tabla clientes")

# Verificar datos en la tabla contrato
ejecutar_consulta("""
    SELECT * FROM contrato LIMIT 5;
""", "Datos de ejemplo en la tabla contrato")

# Verificar datos en la tabla sectores
ejecutar_consulta("""
    SELECT * FROM sectores LIMIT 5;
""", "Datos de ejemplo en la tabla sectores")

# 🔍 5️⃣ Promedio de edad de los clientes en ENVIGADO e ITAGUI
ejecutar_consulta("""
    SELECT s.ciudad AS CIUDAD, ROUND(AVG(cl.edad), 0) AS promedio_edad
    FROM clientes cl
    JOIN contrato c ON cl.id_cliente = c.id_cliente
    JOIN sectores s ON c.id_sector = s.id_sector
    WHERE s.ciudad IN ('ENVIGADO', 'ITAGUI')
    GROUP BY s.ciudad;
""", "Promedio de edad de los clientes en ENVIGADO e ITAGUI")

# 🔍 1️⃣ Top 3 sectores con mayor cantidad de alquileres en el último año
ejecutar_consulta("""
    SELECT s.sector AS sector, COUNT(c.id_contrato) AS cantidad_alquileres
    FROM contrato c
    JOIN sectores s ON c.id_sector = s.id_sector
    WHERE c.fecha_cierre >= date('now', '-1 year')
    GROUP BY s.sector
    ORDER BY cantidad_alquileres DESC
    LIMIT 3;
""", "Top 3 sectores con mayor cantidad de alquileres en el último año")

# 🔍 2️⃣ Zonas con menor valor de canon arrendado en cada año
ejecutar_consulta("""
    SELECT strftime('%Y', c.fecha_cierre) AS ano, s.zona, MIN(c.canon) AS menor_valor_canon
    FROM contrato c
    JOIN sectores s ON c.id_sector = s.id_sector
    GROUP BY ano, s.zona
    ORDER BY ano, menor_valor_canon
    LIMIT 2;
""", "Las 2 zonas con menor valor de canon arrendado en cada año")

# 🔍 3️⃣ Valor promedio de alquiler por tipo de inmueble en cada zona
ejecutar_consulta("""
    SELECT t.nombre_inmueble, s.zona, ROUND(AVG(c.canon), 2) AS promedio_alquiler
    FROM contrato c
    JOIN tipo_de_inmueble t ON c.id_inmueble = t.id_inmueble
    JOIN sectores s ON c.id_sector = s.id_sector
    GROUP BY t.nombre_inmueble, s.zona;
""", "Valor promedio de alquiler por tipo de inmueble en cada zona")

# 🔍 4️⃣ Ingreso mensual promedio por zona en los últimos 6 meses
ejecutar_consulta("""
    SELECT s.zona, ROUND(AVG(c.canon), 2) AS ingreso_mensual_promedio
    FROM contrato c
    JOIN sectores s ON c.id_sector = s.id_sector
    WHERE c.fecha_cierre >= date('now', '-6 months')
    GROUP BY s.zona;
""", "Ingreso mensual promedio por zona en los últimos 6 meses")

# 🔍 6️⃣ Porcentaje de contratos cerrados por tipo de inmueble
ejecutar_consulta("""
    SELECT t.nombre_inmueble, 
           ROUND(COUNT(c.id_contrato) * 100.0 / (SELECT COUNT(*) FROM contrato), 2) AS porcentaje_contratos
    FROM contrato c
    JOIN tipo_de_inmueble t ON c.id_inmueble = t.id_inmueble
    GROUP BY t.nombre_inmueble;
""", "Porcentaje de contratos cerrados por tipo de inmueble")

# 🔍 7️⃣ Top 5 clientes con mayores ingresos mensuales que han firmado contratos en 2024
ejecutar_consulta("""
    SELECT cl.id_cliente, cl.ingresos_mensuales
    FROM clientes cl
    JOIN contrato c ON cl.id_cliente = c.id_cliente
    WHERE strftime('%Y', c.fecha_cierre) = '2024'
    ORDER BY cl.ingresos_mensuales DESC
    LIMIT 5;
""", "Top 5 clientes con mayores ingresos mensuales que han firmado contratos en 2024")

# Guardar el PDF
pdf.output("Resultados_Consultas.pdf")

# Cerrar conexión
conn.close()
