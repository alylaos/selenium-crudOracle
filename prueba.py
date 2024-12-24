import pymongo
import pandas as pd
from pandas import json_normalize
from datetime import datetime,date

import cx_Oracle
import MySQLdb

# Conexión a la base de datos Oracle
db_conn = cx_Oracle.connect("ruta")
cursor1 = db_conn.cursor()

# Configurar el formato de fecha y timestamp
cursor1.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS' NLS_TIMESTAMP_FORMAT = 'DD-MM-YYYY HH24:MI:SS.FF'")

# Ejecutar la consulta SQL
cursor1.execute("""SELECT 
                T.NUMERO_DOCUMENTO
                --COUNT(*)
                FROM UADRYAN.V_TRABAJADOR T
                LEFT JOIN UADRYAN.TRABAJADOR_DETALLES TD ON T.CODIGO_UNICO = TD.CODIGO_UNICO
                WHERE T.compania = '01'
                AND T.situacion_trabajador IN ('A','P')
                AND T.tipo_trabajador IN ('E', 'O')
                AND T.unidad_funcional_organica IN ('00000082', '00000109', '00000084', '00000106', '00000107')
                AND TD.IDNUMERO IS NULL 
                """)

# Obtener los resultados de la consulta

results = cursor1.fetchall()
documento = results[0][0]
documento_limpio = documento.strip()
print(documento_limpio)
# Cerrar el cursor y la conexión a la base de datos
cursor1.close()
db_conn.close()

# Convertir los resultados a un DataFrame de Pandas
df = pd.DataFrame(results)

# Obtener los nombres de las columnas
#column_names = [col[0] for col in cursor1.description]

# Asignar los nombres de las columnas al DataFrame
#df.columns = column_names

# Exportar el DataFrame a un archivo Excel
nombre_archivo = 'trabajadores.xlsx'
df.to_excel(nombre_archivo, index=False, header=False)
