import pymongo
import pandas as pd
from pandas import json_normalize
from datetime import datetime, date
from pandas import DataFrame, Series, NA, NaT
import cx_Oracle
import gspread
from google.oauth2 import service_account
from google.auth.transport.requests import Request
import os.path
import pickle

# Conexión a la base de datos Oracle
db_conn = cx_Oracle.connect("ruta")
cursor1 = db_conn.cursor()

# Configurar el formato de fecha y timestamp
cursor1.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS' NLS_TIMESTAMP_FORMAT = 'DD-MM-YYYY HH24:MI:SS.FF'")

# Ejecutar la consulta SQL
cursor1.execute("""select TRIM(codigo_postulante),TRIM(numero_documento),
                TRIM(nombre),TRIM(apellido_paterno), TRIM(apellido_materno), fecha_nacimiento, email from uadryan.postulante""")

# Obtener los resultados de la consulta
listado = cursor1.fetchall()
print(listado)
lista_mysql = [['codigo','dni','nombre','apellido_paterno','apellido_materno','fecha_nacimiento','email']]

for r in listado:
    lista_mysql.append(r)
encabezado = lista_mysql[0]

cursor1.close()
db_conn.close()

df_main = pd.DataFrame(lista_mysql[1:], columns=encabezado)

df_main['fecha_nacimiento'] = df_main['fecha_nacimiento'].astype(str)

# Exportar el DataFrame a Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

creds = None

# Cargar las credenciales desde el archivo credentials.json
credentials_file = 'credentials.json'
pickle_file = 'token.pickle'

if os.path.exists(pickle_file):
    with open(pickle_file, 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = service_account.Credentials.from_service_account_file(credentials_file, scopes=scope)
        with open(pickle_file, 'wb') as token:
            pickle.dump(creds, token)

client = gspread.authorize(creds)

# ID de la hoja de cálculo y nombre de la hoja
spreadsheet_id = '1hzs6_PASl9STUAwsZRaWXAE_PBk62C6VEI09fkwpk7k'
worksheet_name = 'postulantes_v'

# Abre la hoja de cálculo y la hoja usando el ID
spreadsheet = client.open_by_key(spreadsheet_id)
worksheet = spreadsheet.worksheet(worksheet_name)

# Limpia la hoja antes de escribir los datos
worksheet.clear()

# Escribe los datos del DataFrame en la hoja de cálculo
worksheet.update([df_main.columns.values.tolist()] + df_main.values.tolist())

print("Los datos se han exportado correctamente a Google Sheets.")
