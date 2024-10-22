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
                "startColumnIndex": 4,  # Columna E (índice 4)
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

    # Ejecutar la solicitud para añadir la validación de datos
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

# Abre la hoja de cálculo
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1FMYazlh51vOWSt32Q5bOD_6WpjvhAwIx_aRLYNcZfUk/edit#gid=0'
sheet = client.open_by_url(spreadsheet_url).sheet1

# Carpeta base para compras ágiles
base_folder = r'C:\Users\lild\Desktop\Proyectos\SCM\Compras agiles'

# Cargar datos de compras desde todas las subcarpetas
dataframes_compras = []
for root, dirs, files in os.walk(base_folder):
    for file in files:
        if file.endswith(('.xls', '.xlsx')):  # Filtrar archivos de Excel
            file_path = os.path.join(root, file)
            print(f"Procesando archivo: {file_path}")
            try:
                df = pd.read_excel(file_path, usecols="A,B,E", engine='openpyxl')
                if not df.empty:
                    region_name = os.path.splitext(file)[0]  # Usar el nombre del archivo como nombre de región
                    # Añadir la columna de tipo y región al DataFrame
                    df['Tipo'] = 'Compra ágil'  # Asignar el tipo
                    df['Región'] = region_name  # Asignar el nombre de región
                    dataframes_compras.append(df)
                else:
                    print(f"El archivo {file_path} está vacío.")
            except ValueError:
                try:
                    df = pd.read_excel(file_path, usecols="A,B,E", engine='xlrd')
                    if not df.empty:
                        region_name = os.path.splitext(file)[0]
                        df['Tipo'] = 'Compra ágil'
                        df['Región'] = region_name
                        dataframes_compras.append(df)
                    else:
                        print(f"El archivo {file_path} está vacío.")
                except Exception as e:
                    print(f"Error al leer el archivo {file_path}: {e}")

# Concatenar y limpiar datos de compras
if dataframes_compras:
    combined_df_compras = pd.concat(dataframes_compras, ignore_index=True).drop_duplicates()

    # Leer datos existentes en Google Sheets
    existing_data = sheet.get_all_values()[1:]  # Ignorar encabezados
    existing_codes = {row[0] for row in existing_data}  # Columna A

    # Filtrar duplicados
    filtered_df_compras = combined_df_compras[~combined_df_compras.iloc[:, 0].isin(existing_codes)]

    # Escribir en Google Sheets de compras
    if not filtered_df_compras.empty:
        # Crear lista de valores para el update
        values_to_write = filtered_df_compras.values.tolist()
        rows_to_append = [[row[0], row[1], row[2], "", "", "", row[3], row[4]] for row in values_to_write]  # Cambiar el orden para que Tipo esté en G y Región en H

        # Agrupar en una sola solicitud
        sheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')

        print("Datos de compras ágiles escritos en Google Sheet:")
        print(filtered_df_compras)

        # Verificar si la columna E ya tiene datos
        existing_values_col_e = sheet.col_values(5)[1:]  # Ignorar encabezados
        if not any(existing_values_col_e):  # Si la columna está vacía
            # Añadir lista desplegable en la columna "Estado" para compras
            add_data_validation(sheet, 2, len(filtered_df_compras) + 1)
            print("Lista desplegable añadida para compras ágiles.")
        else:
            print("La lista desplegable ya existe en la columna E.")
    else:
        print("No se encontraron datos nuevos de compras ágiles para procesar.")
else:
    print("No se encontraron datos de compras ágiles para procesar.")
