import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import warnings
import os

# Ignorar advertencias de openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Configuración de Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r'C:\Users\lild\Desktop\Proyectos\SCM\credentials\credenciales.json',
    scope
)
client = gspread.authorize(creds)

# Función para agregar validación de datos
def add_data_validation(sheet, start_row, end_row):
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = sheet.spreadsheet.id
    valid_values = ['Postuladas', 'No Postuladas']

    requests = [{
        "setDataValidation": {
            "range": {
                "sheetId": sheet.id,
                "startRowIndex": start_row - 1,
                "endRowIndex": end_row,
                "startColumnIndex": 4,
                "endColumnIndex": 5
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": value} for value in valid_values]
                },
                "inputMessage": "Selecciona 'Postuladas' o 'No Postuladas'",
                "strict": True,
                "showCustomUi": True
            }
        }
    }]

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
    except Exception as e:
        print(f"Error al agregar validación de datos: {e}")

# Abre la hoja de cálculo
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1hhW4m8J9uVLLhT4SqpUFtKWo0OOsqA-IrHwomLYN1l4/edit?gid=0#gid=0'
sheet = client.open_by_url(spreadsheet_url).sheet1

# Cargar datos de licitaciones
excel_file = r'C:\Users\lild\Desktop\Proyectos\SCM\Licitaciones\Licitaciones.xlsx'
if os.path.exists(excel_file):
    print(f"Procesando archivo de licitaciones: {excel_file}")
    data_licitaciones = pd.read_excel(excel_file, usecols='A,B,F').astype(str)
    
    # Añadir columna 'Tipo' y 'Región'
    data_licitaciones['Tipo'] = ''
    data_licitaciones['Región'] = ''  # Puedes llenar esto si tienes una lógica específica para las regiones

    # Leer datos existentes en Google Sheets
    existing_data = sheet.get_all_values()[1:]  # Ignorar encabezados
    existing_codes = {row[0] for row in existing_data}  # Columna A

    # Filtrar duplicados
    filtered_df_licitaciones = data_licitaciones[~data_licitaciones.iloc[:, 0].isin(existing_codes)]

    # Escribir en Google Sheets de licitaciones
    if not filtered_df_licitaciones.empty:
        # Crear lista de valores para el update
        values_to_write = filtered_df_licitaciones.values.tolist()
        rows_to_append = [[row[0], row[1], row[2], "", "", "", row[3], row[4]] for row in values_to_write]

        # Agrupar en una sola solicitud
        sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')

        print("Datos de licitaciones escritos en Google Sheet:")
        print(filtered_df_licitaciones)

        # Añadir lista desplegable en la columna "Estado" para licitaciones
        start_row_licitaciones = len(existing_data) + 2  # Empieza justo después de los datos existentes
        end_row_licitaciones = start_row_licitaciones + len(values_to_write)

        # Verificar si las filas existen antes de agregar la validación
        if start_row_licitaciones <= end_row_licitaciones:
            add_data_validation(sheet, start_row_licitaciones, end_row_licitaciones)
            print("Lista desplegable añadida para licitaciones.")
        else:
            print("No hay filas válidas para agregar validación de datos.")
    else:
        print("No se encontraron datos nuevos de licitaciones para procesar.")
else:
    print(f"El archivo de licitaciones {excel_file} no existe.")
