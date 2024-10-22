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

# Abre la hoja de cálculo en Google Sheets
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1hhW4m8J9uVLLhT4SqpUFtKWo0OOsqA-IrHwomLYN1l4/edit'
sheet = client.open_by_url(spreadsheet_url).sheet1  # Mantener la hoja original en Google Sheets

# Cargar datos de vendedores desde la hoja "Licitaciones" del Excel
vendedores_df = pd.read_excel(r'C:\Users\lild\Desktop\Proyectos\SCM\Archivo base\Compras_agiles.xlsx', sheet_name='Licitaciones')

# Crear un diccionario de vendedores y proporciones
vendor_to_proportions = {}

for _, row in vendedores_df.iterrows():
    vendor = row.iloc[0]  # Columna A (índice 0): Vendedor
    proportion = row.iloc[1]  # Columna B (índice 1): Proporción
    vendor_to_proportions[vendor] = proportion

# Obtener datos existentes de Google Sheets
existing_data = sheet.get_all_values()[1:]  # Ignorar encabezados
updates = []

# Asignar vendedores según la proporción
vendor_cycle = []
total_proportions = 0
cycle_index = 0

# Preparar un ciclo de vendedores con base en proporciones
for vendor, proportion in vendor_to_proportions.items():
    vendor_cycle.extend([vendor] * proportion)
total_proportions = len(vendor_cycle)

# Verificar que el ciclo de vendedores no esté vacío
if total_proportions == 0:
    print("No hay vendedores disponibles.")
    exit()

# Asignar vendedores a la columna D
for row in existing_data:
    current_vendor = row[3]  # Columna D (índice 3)

    # Solo procesar si la columna D está vacía
    if not current_vendor:
        # Asignar vendedor a la compra actual
        updates.append((row[0], row[1], row[2], vendor_cycle[cycle_index]))  # Asignar el vendedor a la columna D
        cycle_index = (cycle_index + 1) % total_proportions  # Avanzar el índice en el ciclo
    else:
        updates.append((row[0], row[1], row[2], current_vendor))  # Mantener el valor actual en la columna D

# Preparar las actualizaciones para la columna D
data = []
for idx, (col_a, col_b, col_c, vendor) in enumerate(updates, start=2):  # start=2 para ignorar encabezados
    if vendor != current_vendor:  # Solo preparar si el vendedor a asignar es diferente
        range_string = f'Hoja 1!D{idx}'  # Cambiar "Hoja 1" a la hoja correcta
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
