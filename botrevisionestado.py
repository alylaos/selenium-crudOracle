import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

# Cargar el archivo Excel
file_path = 'C:\\bot\\prueba.xlsx'
df = pd.read_excel(file_path)
browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

def obtener_estado(ruc, max_reintentos=150):
    intentos = 0
    while intentos < max_reintentos:
        try:
            url = 'https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp'
            browser.get(url)

            busqueda = browser.find_element(By.XPATH, '//*[@id="txtRuc"]')
            time.sleep(15)
            busqueda.send_keys(ruc)
            time.sleep(15)

            boton_busqueda = browser.find_element(By.XPATH, '//*[@id="btnAceptar"]')
            time.sleep(10)
            boton_busqueda.click()
            time.sleep(10)

            estado = browser.find_element(By.XPATH, '/html/body/div/div[2]/div/div[3]/div[2]/div[5]/div/div[2]/p').text
            return estado
        except:
            intentos += 1
            print(f"Intento {intentos} fallido para el RUC: {ruc}. Reintentando...")
            time.sleep(10)
            time.sleep(5)
            browser.refresh()
    return 'Estado no encontrado'

for index, row in df.iterrows():
    if index < 2707:  # Saltar a la fila que quieres, en este caso fila 4 (index 3)
        continue
    
    ruc = row.iloc[0]
    estado_actual = df.iat[index, 20]
    
    if estado_actual == 'ACTIVO':
        continue  

    if isinstance(ruc, float):
        ruc = str(int(ruc))
    else:
        ruc = str(ruc)
    
    print(f"Procesando RUC: {ruc}")
    estado = obtener_estado(ruc)
    
    # Actualizar el estado en la columna 21 de la misma fila
    df.iat[index, 20] = estado
    

    temp_file_path = file_path.replace('.xlsx', '_temp.xlsx')
    df.to_excel(temp_file_path, index=False)

 
    os.replace(temp_file_path, file_path)

    print(f"Estado para RUC {ruc} guardado en la columna 21")

print("Estados obtenidos y guardados en el archivo Excel según cada RUC.")
browser.quit()


