import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# Configuración de Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r'C:\Users\lild\Desktop\Proyectos\SCM\credentials\credenciales.json',
    scope
)
client = gspread.authorize(creds)

# URL de la hoja de cálculo original
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1FMYazlh51vOWSt32Q5bOD_6WpjvhAwIx_aRLYNcZfUk/edit#gid=0'
sheet = client.open_by_url(spreadsheet_url).sheet1

# Obtener datos de la hoja
data = sheet.get_all_records()
df = pd.DataFrame(data)

# Convertir la columna 'Fecha de cierre' a tipo datetime
df['Fecha de cierre'] = pd.to_datetime(df['Fecha de cierre'], errors='coerce')

# Obtener la fecha actual y calcular dos meses atrás
current_date = datetime.now()
two_months_ago = current_date.replace(day=1) - pd.DateOffset(months=2)

# Filtrar los datos que cumplen la condición
filtered_df = df[df['Fecha de cierre'] <= two_months_ago]

# Crear el nombre del nuevo archivo de respaldo en Google Sheets
backup_month_year = current_date.strftime("%m-%Y")
backup_title = f"Respaldo {backup_month_year}"

# Crear una nueva hoja de cálculo para el respaldo
backup_sheet = client.create(backup_title)
backup_worksheet = backup_sheet.get_worksheet(0)  # Obtener la primera hoja

# Guardar el DataFrame filtrado en la nueva hoja de cálculo
if not filtered_df.empty:
    # Escribir los encabezados
    backup_worksheet.insert_row(filtered_df.columns.tolist(), 1)

    # Escribir los datos
    for index, row in filtered_df.iterrows():
        backup_worksheet.append_row(row.values.tolist())

    print(f"Respaldo creado en Google Sheets: {backup_title}")

    # Borrar datos del Google Sheet original
    rows_to_delete = filtered_df.index.tolist()  # Obtener los índices de las filas a eliminar
    rows_to_delete = [i + 2 for i in rows_to_delete]  # Ajustar para la numeración de Google Sheets (empezando desde 2)

    for row in sorted(rows_to_delete, reverse=True):  # Borrar de abajo hacia arriba
        sheet.delete_rows(row)

    print("Datos respaldados y eliminados del Google Sheet original.")
else:
    print("No hay datos que respaldar para el periodo especificado.")
