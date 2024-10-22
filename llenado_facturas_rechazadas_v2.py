import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuración de credenciales y acceso a Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets"]
creds = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lild/Desktop/Proyectos/SCM/credentials/credenciales.json', scope)
client = gspread.authorize(creds)

# Leer el archivo Excel
excel_path = 'C:/Users/lild/Desktop/Proyectos/SCM/Facturas Rechazadas/Facturas_excel.xlsx'
excel_data = pd.read_excel(excel_path)

# Filtrar los datos: columna de índice 0 debe ser igual a 33 y columna de índice 3 no debe estar vacía
filtered_data = excel_data[(excel_data.iloc[:, 0] == 33) & (excel_data.iloc[:, 3].notna())]

# Asegúrate de que filtered_data no contenga filas vacías
filtered_data = filtered_data.dropna(how='all')

# Convertir la columna 'Folio' a string para evitar conflictos de tipo
folio_index_excel = 6  # Índice para 'Folio' en Excel
filtered_data.iloc[:, folio_index_excel] = filtered_data.iloc[:, folio_index_excel].astype(str)

# Abrir Google Sheets
sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/19eseyq45ZP4VNdhFHYl4dA34fuO6hhwdgqe8UPEw_FY/edit')
worksheet = sheet.get_worksheet(0)  # Obtener la primera hoja

# Obtener datos existentes en Google Sheets
google_data = pd.DataFrame(worksheet.get_all_records(head=1))

# Verificar si google_data tiene columnas
if not google_data.empty and google_data.shape[1] > 5:  # Columna 5 para 'Folio' en Google Sheets
    google_data.iloc[:, 5] = google_data.iloc[:, 5].astype(str)

    # Eliminar duplicados basados en las columnas 'Folio' (Excel en índice 6, Google Sheets en índice 5)
    unique_data = filtered_data[~filtered_data.iloc[:, folio_index_excel].isin(google_data.iloc[:, 5])]
else:
    unique_data = filtered_data

# Verificar cuántas filas se están procesando
print(f"Total de filas únicas a subir: {unique_data.shape[0]}")

# Encontrar la última fila libre en Google Sheets
last_row = len(worksheet.get_all_values()) + 1  # Sumar 1 para la siguiente fila libre

for _, row in unique_data.iterrows():
    print(f"Actualizando fila para el índice en Google Sheets: {last_row}")  # Mostrar el índice en Google Sheets
    update_row = row.astype(str).tolist()

    # Omitir la columna de índice 0
    update_row.pop(0)  # Quitar la columna de índice 0

    # Leer la fila existente
    existing_row = worksheet.row_values(last_row)

    # Mantener las columnas 7 a 11 en Google Sheets intactas
    for i in range(7, 12):
        if i < len(existing_row):
            update_row.insert(i, existing_row[i])
        else:
            update_row.insert(i, '')

    # Limitar update_row a las primeras 7 columnas (A-G)
    update_row = update_row[:7]

    # Definir el rango para escribir
    range_to_update = f'A{last_row}:G{last_row}'  # Ajustar el rango también

    # Actualizar los datos en Google Sheets
    worksheet.update(range_name=range_to_update, values=[update_row])  # Usar argumentos nombrados

    # Incrementar la fila libre
    last_row += 1

print("Datos subidos a Google Sheets sin duplicados desde el Excel filtrado.")
