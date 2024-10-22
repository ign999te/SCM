import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import time

# Configuración de Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r'C:\Users\lild\Desktop\Proyectos\SCM\credentials\credenciales.json',
    scope
)
client = gspread.authorize(creds)

# Abre la hoja de cálculo
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1FMYazlh51vOWSt32Q5bOD_6WpjvhAwIx_aRLYNcZfUk/edit#gid=0'
sheet = client.open_by_url(spreadsheet_url).sheet1

# Cargar datos de vendedores
vendedores_df = pd.read_excel(r'C:\Users\lild\Desktop\Proyectos\SCM\Archivo base\Compras_agiles.xlsx', sheet_name='Vendedores')

region_to_vendors = {}

# Crear un diccionario de regiones a vendedores y proporciones
for _, row in vendedores_df.iterrows():
    region = row.iloc[0]  # Columna A (índice 0): Región
    vendors = row.iloc[1].split(';')  # Columna B (índice 1): Vendedores

    if len(vendors) > 1:
        proportions = list(map(int, row.iloc[2].split(';')))  # Columna C (índice 2): Proporciones
    else:
        proportions = [1]  # Solo un vendedor, asignar 1 por defecto

    region_to_vendors[region] = list(zip(vendors, proportions))

# Obtener datos existentes de Google Sheets
existing_data = sheet.get_all_values()[1:]  # Ignorar encabezados
updates = []

# Asignar vendedores según la región
current_region = None
vendor_cycle = []
total_proportions = 0
cycle_index = 0

for row in existing_data:
    region = row[7]  # Columna H (índice 7)
    current_vendor = row[3]  # Columna D (índice 3)

    # Verificar si hay datos en la columna H y si la columna D está vacía
    if region and not current_vendor:  # Solo procesar si D está vacío
        if region != current_region:
            # Si cambiamos de región, reiniciar el ciclo
            current_region = region
            if region in region_to_vendors:
                vendors_info = region_to_vendors[region]
                total_proportions = sum(proportion for _, proportion in vendors_info)

                vendor_cycle = []
                for vendor, proportion in vendors_info:
                    vendor_cycle.extend([vendor] * proportion)  # Crear un ciclo de vendedores según la proporción
                cycle_index = 0  # Reiniciar el índice del ciclo
            else:
                vendor_cycle = []  # No hay vendedores para esta región

        # Asignar vendedor a la compra actual
        if vendor_cycle:
            updates.append((row[0], row[1], row[2], vendor_cycle[cycle_index]))  # Asignar el vendedor a la columna D
            cycle_index = (cycle_index + 1) % total_proportions  # Avanzar el índice en el ciclo
        else:
            updates.append((row[0], row[1], row[2], ''))  # Sin vendedor asignado
    else:
        updates.append((row[0], row[1], row[2], current_vendor))  # Mantener el valor actual en la columna D

# Preparar las actualizaciones para la columna D
data = []
for idx, (col_a, col_b, col_c, vendor) in enumerate(updates, start=2):  # start=2 para ignorar encabezados
    if vendor != current_vendor:  # Solo preparar si el vendedor a asignar es diferente
        range_string = f"Hoja 1!D{idx}"  # Cambiar "Sheet1" a "Hoja 1"
        print(f"Preparando actualización para: {range_string} con vendedor: {vendor}")  # Imprimir el rango y el vendedor
        data.append({
            "range": range_string,
            "values": [[vendor]],
        })

# Realizar actualizaciones en bloque
if data:
    body = {
        "valueInputOption": "USER_ENTERED",
        "data": data
    }
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = sheet.spreadsheet.id

    while True:
        try:
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            print("Actualizaciones de vendedores completadas.")
            break  # Salir del bucle si la actualización fue exitosa
        except Exception as e:
            if "429" in str(e):  # Si es error 429, espera y reintenta
                print("Quota exceeded, waiting 60 seconds...")
                time.sleep(60)
            else:
                print(f"Error: {e}")
                break
else:
    print("No hay actualizaciones para realizar.")
