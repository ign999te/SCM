import pandas as pd

# Cargar el archivo Excel
file_path = r"C:\Users\lild\Desktop\Proyectos\SCM\Facturas Rechazadas\Facturas_excel.xlsx"
df = pd.read_excel(file_path)

# Borrar las columnas de O a X (índices 14 a 23) y de AA a AQ (índices 26 a 43)
df.drop(df.columns[14:24], axis=1, inplace=True)  # O a X
df.drop(df.columns[25:43], axis=1, inplace=True)  # AA a AQ (ajustado)

# Eliminar las columnas de Q a Z (índices 16 a 25)
df.drop(df.columns[16:26], axis=1, inplace=True)  # Q a Z (ajustado por los drops anteriores)

# Eliminar las columnas A, B, C, G, H, I, K, M, N
columns_to_drop = [0, 2, 6, 7, 8, 10, 12, 13]
df.drop(df.columns[columns_to_drop], axis=1, inplace=True)

# Mover la columna C (índice 2) a la posición F (índice 5)
if df.shape[1] > 3:  # Asegurarse de que hay al menos 3 columnas
    col_c = df.iloc[:, 3]  # Guardamos la columna C (índice 2)
    df.drop(df.columns[3], axis=1, inplace=True)  # Eliminamos la columna C
    
    # Asegurarse de que el DataFrame tiene suficientes columnas para insertar en la posición 5
    insert_index = 6 if df.shape[1] > 6 else df.shape[1]
    df.insert(insert_index, col_c.name, col_c)  # Insertamos en la posición F o al final

# Guardar el archivo modificado
df.to_excel(file_path, index=False)
