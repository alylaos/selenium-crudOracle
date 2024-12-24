from collections import namedtuple
from tokenize import String
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from setting import settings as sett
from PIL import Image
from time import sleep
from twocaptcha import TwoCaptcha
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
import cx_Oracle
import openpyxl
import os
import xlrd
from datetime import datetime

# Configuración de Chrome
opts = Options()
opts.add_argument(
    "user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Ubuntu Chromium/71.0.3578.80 Chrome/71.0.3578.80 Safari/537.36")
opts.add_argument("--start-maximized")

download_directory = "C:\\sucamecpdf"
opts.add_experimental_option('prefs', {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
})

s = Service('./utils/chromedriver.exe')
driver = webdriver.Chrome(service=s, options=opts)
# Leyendo los DNIs del excel
band=0
step=0

# Conexión a la base de datos Oracle
db_conn = cx_Oracle.connect("ruta")
cursor1 = db_conn.cursor()

# Configurar el formato de fecha y timestamp
cursor1.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS' NLS_TIMESTAMP_FORMAT = 'DD-MM-YYYY HH24:MI:SS.FF'")

# Ejecuta la consulta
cursor1.execute("""
SELECT
    T.NUMERO_DOCUMENTO, 
    T.MATRICULA
FROM 
    UADRYAN.V_TRABAJADOR T
WHERE
    T.compania = '01'
    AND T.situacion_trabajador IN ('A','P')
    AND T.tipo_trabajador IN ('E', 'O')
    AND T.unidad_funcional_organica IN ('00000082', '00000109', '00000084', '00000106', '00000107')
    AND 
    T.CODIGO_UNICO 
    NOT IN (
        SELECT TD.CODIGO_UNICO 
        FROM UADRYAN.TRABAJADOR_DETALLES TD 
        WHERE TD.fecha_creacion >= '2/07/2024'
    )
""")
print(cursor1)

oracle_results = cursor1.fetchall()

print("INICIO")

for row in oracle_results:
    documento = row[0]
    print(documento)
    matricula = row[1]
    print(matricula)
    if not documento:
        continue  # Si el número de documento está vacío, se salta esta iteración

    documento = documento.strip()
    print(documento)
    
    num_licencia = ""
    fechaEmision_lic = ""
    fechaVencimiento_lic = ""
    licencia = ""
        
    num_carnet = ""
    fec_emis_carnet = ""
    fec_cadu_carnet = ""
    carnet = ""

    fec_emis_curso = ""
    fechaVencimiento_curso= ""
    curso = ""

    nombre_cliente = ""

    # logeo
    if(step==0):
        driver.get(sett.urlSucamec())
        usuario_Sucamec = sett.userSucamec()
        clave_Sucamec =  sett.pass_sucamec()
        ruc_Sucamec = sett.rucSucamec()
        #user = "06170598" #Cambiar DNI
        solver = TwoCaptcha(sett.captchaKey())
        time.sleep(7)
        img_element = driver.find_element(By.ID,"loginForm:imgCaptcha")
        location = img_element.location
        size = img_element.size
        time.sleep(7)
        driver.save_screenshot('./imgCaptcha/captchaIMG.png')
        #720,390,1030,480
        Im = Image.open('./imgCaptcha/captchaIMG.png')
        ##Im = Im.crop((int(x),int(y),int(width),int(height)))
        ##Im = Im.crop((871,258,1075,322))
        Im = Im.crop((871, 258, 1075, 322))
        time.sleep(4)
        Im.save('./imgCaptcha/imagen_catcha.png')
        try:
            step=1
            resultado = solver.normal('./imgCaptcha/imagen_catcha.png')
            time.sleep(5)
            print('Resultado: ', str(resultado['code']))
            input_ruc = driver.find_element(By.ID,'loginForm:documento')
            input_ruc.send_keys(ruc_Sucamec)
            input_captcha = driver.find_element(By.ID,'loginForm:textoCaptcha')
            input_captcha.send_keys(resultado['code'])       
            input_usuario = driver.find_element(By.ID,'loginForm:usuario')
            input_usuario.send_keys(usuario_Sucamec)
            input_clave = driver.find_element(By.ID,'loginForm:clave')
            input_clave.send_keys(clave_Sucamec)
            time.sleep(2)
            consultar_boton = driver.find_element(By.ID,'loginForm:ingresar')
            consultar_boton.click()
            time.sleep(4)
            driver.get("https://www.sucamec.gob.pe/sel/faces/aplicacion/consultas/DcGsspTotVigilantes/List.xhtml")
        except NoSuchElementException:
            try:
                time.sleep(7)
                print("TRY 1 ")
                driver.save_screenshot('./imgCaptcha/captchaIMG.png')
                Im = Image.open('./imgCaptcha/captchaIMG.png')
                Im = Im.crop((871, 258, 1075, 322))
                time.sleep(4)
                Im.save('./imgCaptcha/imagen_catcha.png')
                resultado = solver.normal('./imgCaptcha/imagen_catcha.png')
                time.sleep(4)
                print('Resultado: ', str(resultado['code']))
                input_captcha = driver.find_element(By.ID,'loginForm:textoCaptcha')
                input_captcha.send_keys(resultado['code'])
                input_clave = driver.find_element(By.ID,'loginForm:clave')
                input_clave.send_keys(clave_Sucamec)
                time.sleep(2)
                consultar_boton = driver.find_element(By.ID,'loginForm:ingresar')
                consultar_boton.click()
                time.sleep(4)
                #Navegando en la intranet
                driver.get("https://www.sucamec.gob.pe/sel/faces/aplicacion/consultas/DcGsspTotVigilantes/List.xhtml")
            except NoSuchElementException:
                print("TRY 2")
                time.sleep(7)
                driver.save_screenshot('./imgCaptcha/captchaIMG.png')
                Im = Image.open('./imgCaptcha/captchaIMG.png')
                Im = Im.crop((871, 258, 1075, 322))
                time.sleep(4)
                Im.save('./imgCaptcha/imagen_catcha.png')
                resultado = solver.normal('./imgCaptcha/imagen_catcha.png')
                time.sleep(4)
                print('Resultado: ', str(resultado['code']))
                input_captcha = driver.find_element(By.ID,'loginForm:textoCaptcha')
                input_captcha.send_keys(resultado['code'])
                input_clave = driver.find_element(By.ID,'loginForm:clave')
                input_clave.send_keys(clave_Sucamec)
                time.sleep(2)
                consultar_boton = driver.find_element(By.ID,'loginForm:ingresar')
                consultar_boton.click()
                time.sleep(4)
                driver.get("https://www.sucamec.gob.pe/sel/faces/aplicacion/consultas/DcGsspTotVigilantes/List.xhtml")

                
    # busqueda x DNI   
    if(step==1):
        time.sleep(7)
        filtro_intranet = driver.find_element(By.XPATH,'//*[@id="buscarForm:tipoId"]')
        filtro_intranet.click()  

        # verifica el tipo de doocumento
        if len(documento) == 8:
            time.sleep(7)
            filtro_dni_intranet = driver.find_element(By.XPATH,'//*[@id="buscarForm:tipoId_1"]')
            time.sleep(5)

            filtro_dni_intranet.click()    
            time.sleep(7)
        elif len(documento) == 9:
            time.sleep(7)
            filtro_cext_intranet = driver.find_element(By.XPATH,'//*[@id="buscarForm:tipoId_2"]')
            filtro_cext_intranet.click()    
            time.sleep(7)
        
        input_intranet = driver.find_element(By.XPATH,'//*[@id="buscarForm:j_idt32"]')
        input_intranet.clear()
        input_intranet.send_keys(documento)
        time.sleep(2)
        time.sleep(2)

        boton_buscar = driver.find_element(By.XPATH,'//*[@id="buscarForm:botonBuscar"]/span[2]')
        boton_buscar.click()
        time.sleep(5)
        time.sleep(5)
        time.sleep(5)

        
        txt_ver = driver.find_element(By.XPATH,'//*[@id="buscarForm:buscarDatatable_data"]/tr/td').text
        time.sleep(5)
        time.sleep(5)
        time.sleep(5)

        # Verificar si hay resultados
        try:
            # Empleados que tiene más de un carné 
            try:
                link_ver = driver.find_element(By.XPATH,'//*[@id="buscarForm:buscarDatatable:3:j_idt65"]')
                link_ver.click()
                time.sleep(8)
            except NoSuchElementException:
                try:
                    link_ver = driver.find_element(By.XPATH,'//*[@id="buscarForm:buscarDatatable:2:j_idt65"]')
                    link_ver.click()
                    time.sleep(8)
                except NoSuchElementException:
                    try:
                        link_ver = driver.find_element(By.XPATH,'//*[@id="buscarForm:buscarDatatable:1:j_idt65"]')
                        link_ver.click()
                        time.sleep(8)
                    except NoSuchElementException:
                        link_ver = driver.find_element(By.XPATH,'//*[@id="buscarForm:buscarDatatable:0:j_idt65"]')
                        link_ver.click()
                        time.sleep(10)
            
            try:
                num_carnet= driver.find_element(By.XPATH,'//*[@id="verForm:j_idt72"]').text
                print("PRIMER INTENTO")
            except NoSuchElementException:
                try:
                    time.sleep(10)
                    num_carnet= driver.find_element(By.XPATH,'//*[@id="verForm:j_idt72"]').text
                    print("2DO INTENTO")
                except NoSuchElementException:
                    print("3ER INTENTO")
                    num_carnet = ""
            
            
            try:
                fechaVencimiento_curso= driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[6]').text
                print("PRIMER INTENTO")
            except NoSuchElementException:
                try:
                    time.sleep(10)
                    fechaVencimiento_curso= driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[6]').text
                    print("2DO INTENTO")
                except NoSuchElementException:
                    print("3ER INTENTO")
                    fechaVencimiento_curso = ""
            try:
                fec_emis_curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[5]').text
            except NoSuchElementException:
                try:
                    fec_emis_curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[5]').text
                except NoSuchElementException:
                    print("3ER INTENTO")
                    fec_emis_curso = ""

            try:
                curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[7]').text
            except NoSuchElementException:
                try:
                    curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr[1]/td[7]').text
                except NoSuchElementException:
                    curso = ""

            time.sleep(5)

            print(len(fechaVencimiento_curso))
            print(len(fec_emis_curso))
            print(len(curso))
            
            # Recorre la tabla del curso sucamec
            aux = 2
            if len(fechaVencimiento_curso)==0 and len(fec_emis_curso)==0 and len(curso)==0:
                    print("cuadro vacio")
            elif len(fechaVencimiento_curso)==0 and len(fec_emis_curso)==0:
                print("x")
                while len(fechaVencimiento_curso)==0 and len(fec_emis_curso)==0:
                    print("ENTRO")
                    fechaVencimiento_curso= driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr['+str(aux)+']/td[6]').text
                    fec_emis_curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr['+str(aux)+']/td[5]').text
                    curso = driver.find_element(By.XPATH,'//*[@id="verForm:buscarCurDatatable_data"]/tr['+str(aux)+']/td[7]').text
                    aux=aux+1
        
            txt_licencia = driver.find_element(By.XPATH,'//*[@id="verForm:licDatatable_data"]/tr/td').text
            if txt_licencia == "No se encontraron resultados.":
                licencia = "NO"
            else:
                licencia = "SI"
                fechaVencimiento_lic= driver.find_element(By.XPATH,'//*[@id="verForm:licDatatable_data"]/tr/td[3]').text
                time.sleep(1)
                fechaEmision_lic= driver.find_element(By.XPATH,'//*[@id="verForm:licDatatable_data"]/tr/td[2]').text
                time.sleep(1)
                num_licencia = driver.find_element(By.XPATH,'//*[@id="verForm:licDatatable_data"]/tr/td[1]').text
                
            try:
                carnet = driver.find_element(By.XPATH,'//*[@id="verForm:j_idt96"]').text
            except NoSuchElementException:
                carnet=""


            try:
                fec_emis_carnet = driver.find_element(By.XPATH,'//*[@id="verForm:j_idt88"]').text
            except NoSuchElementException:
                fec_emis_carnet=""

            try:
                fec_cadu_carnet = driver.find_element(By.XPATH,'//*[@id="verForm:j_idt90"]').text
            except NoSuchElementException:
                fec_cadu_carnet=""

            try:
                nombre_cliente = driver.find_element(By.XPATH,'//*[@id="verForm:j_idt78"]').text
            except NoSuchElementException:
                nombre_cliente=""
            

            boton_buscar = driver.find_element(By.XPATH,'//*[@id="verForm:j_idt206"]/span[2]')
            boton_buscar.click()
            time.sleep(7)
            input_intranet = driver.find_element(By.XPATH,'//*[@id="buscarForm:j_idt32"]').clear()

            print("------------------------------------")
            nombre_cliente_1 = nombre_cliente[0:50]
            print("NOMBRE CLIENTE:",nombre_cliente_1)
           

            print("NUM LICENCIA: ",num_licencia)
            print("FECHA EMISION LICENCIA: ",fechaEmision_lic)
            print("FECHA VENCIMIENTO LICENCIA: " ,fechaVencimiento_lic)
            
            print("NUM CARNÉ: ",num_carnet)
            print("FECHA EMISION CARNÉ: ",fec_emis_carnet)
            print("FECHA VENCIMIENTO CARNÉ: ",fec_cadu_carnet)

            print("FECHA EMISION CURSO SUCAMEC:",fec_emis_curso)
            print("FECHA VENCIMIENTO CURSO SUCAMEC:" ,fechaVencimiento_curso)
        
            print("------------------------------------")

            time.sleep(7)
            boton_gssp = driver.find_element(By.XPATH, '//*[@id="j_idt11:menuPrincipal"]/div[3]/h3/a')
            boton_gssp.click()
            time.sleep(5)

            boton_carne = driver.find_element(By.XPATH, '//*[@id="2_1"]/a/span[2]')
            boton_carne.click()
            time.sleep(5)

            boton_impresion = driver.find_element(By.XPATH, '//*[@id="2_1"]/ul/li[4]/a/span')
            boton_impresion.click()
            time.sleep(5)


            documento
            documento = str(documento).split(".")[0]
            print(f"Procesando fila , Documento: {documento}")

            boton_carne = driver.find_element(By.XPATH, '//*[@id="2_1"]/a/span[2]')
            boton_carne.click()
            time.sleep(5)
            time.sleep(7)
            boton_gssp = driver.find_element(By.XPATH, '//*[@id="j_idt11:menuPrincipal"]/div[3]/h3/a')
            boton_gssp.click()
            time.sleep(5)
            time.sleep(7)

            buscarpor = driver.find_element(By.XPATH, '//*[@id="impresionForm:buscarPor"]/div[3]/span')
            buscarpor.click()
            time.sleep(5)
            time.sleep(5)

            buscarpordni = driver.find_element(By.XPATH, '//*[@id="impresionForm:buscarPor_3"]')
            buscarpordni.click()
            time.sleep(5)
            time.sleep(5)

            input_intranet = driver.find_element(By.XPATH, '//*[@id="impresionForm:filtroBusqueda"]')
            input_intranet.clear()
            input_intranet.send_keys(documento)
            time.sleep(5)
            time.sleep(5)

            buscar_dni = driver.find_element(By.XPATH, '//*[@id="impresionForm:j_idt40"]/span[2]')
            buscar_dni.click()
            time.sleep(5)
            time.sleep(5)

            try:
                botonpdf = driver.find_element(By.XPATH, '//*[@id="impresionForm:dtResultados:0:btnVerPdf"]/span[1]')
                botonpdf.click()
                time.sleep(5)
                time.sleep(5)

                main_window = driver.current_window_handle
                all_windows = driver.window_handles

                for window in all_windows:
                    if window != main_window:
                        driver.switch_to.window(window)
                        break

                # Esperar a que el archivo PDF se descargue completamente
                download_wait_time = 30  # Ajusta este tiempo según el tamaño del PDF y la velocidad de descarga
                end_time = time.time() + download_wait_time
                while time.time() < end_time:
                    time.sleep(5)
                    files = os.listdir(download_directory)
                    if any(file.endswith('.crdownload') for file in files):
                        continue
                    else:
                        break

                # Renombrar el archivo descargado
                files = os.listdir(download_directory)
                files.sort(key=lambda x: os.path.getctime(os.path.join(download_directory, x)))
                latest_file = os.path.join(download_directory, files[-1])

                new_file_name = os.path.join(download_directory, f"{documento}_sucamec.pdf")
                os.rename(latest_file, new_file_name)

                print(f"Archivo guardado como: {new_file_name}")


            except NoSuchElementException:
                print(f"No se encontró el botón PDF para el documento {documento}. Continuando con el siguiente documento.")

            
            codigo_usuario = "43813803"
            fecha_actual = datetime.now()

            print(fecha_actual)

            if num_licencia and fechaEmision_lic:
                # Insertar datos de Licencia (sin PDF)

                # Obtener el máximo IDNUMCARGA y sumarle 1
                select_max_id_query = "SELECT NVL(MAX(IDNUMCARGA), 0) + 1 FROM uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC"
                cursor1.execute(select_max_id_query)
                new_idnumcarga = cursor1.fetchone()[0]

                insert_licencia = """
                INSERT INTO uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC (MATRICULA, IDTIPO, IDNUMERO, COMENTARIOS, FECHA_EMISION, FECHA_CADUCIDAD, IDNUMCARGA)
                VALUES (:matricula, :tipo, :numero, :comentarios, TO_DATE(:emision, 'DD/MM/YYYY'), TO_DATE(:caducidad, 'DD/MM/YYYY'), :idnumcarga)
                """
                cursor1.execute(insert_licencia, {'matricula': matricula, 'tipo': 'LICENCIA DE SUCAMEC', 'numero': num_licencia,
                                                  'comentarios': nombre_cliente_1, 'emision': fechaEmision_lic,
                                                  'caducidad': fechaVencimiento_lic, 'idnumcarga': new_idnumcarga})
                print("Insertar datos de Licencia")
                


                # Ejecutar procedimiento almacenado para 'LICENCIA DE SUCAMEC' (tipo 1)
                #call_proc = """
                #BEGIN
                #    UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC(
                #        p_codcia => :codcia,
                #        p_idnumcarga => :idnumcarga,
                #        p_tipo => :tipo,
                #        p_codigousuario_crea => :codigousuario_crea,
                #        p_fecha_crea => :fecha_crea,
                #        p_codigousuario_modifica => :codigousuario_modifica,
                #        p_fecha_modifica => :fecha_modifica
                #    );
                #END;
                #"""
                sql = "UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC"
                cursor1.callproc(sql, ['01',new_idnumcarga,'1',codigo_usuario,fecha_actual,codigo_usuario,fecha_actual])
                db_conn.commit()
                print("Sp ejecutado para tipo 1")
                time.sleep(5)

            if num_carnet and new_file_name:
                # Insertar datos de Carné (con PDF)
                # Obtener el máximo IDNUMCARGA y sumarle 1
                select_max_id_query = "SELECT NVL(MAX(IDNUMCARGA), 0) + 1 FROM uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC"
                cursor1.execute(select_max_id_query)
                new_idnumcarga = cursor1.fetchone()[0]

                with open(new_file_name, 'rb') as f:
                    pdf_read = f.read()

                insert_carnet = """
                INSERT INTO uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC (MATRICULA, IDTIPO, IDNUMERO, COMENTARIOS, FECHA_EMISION, FECHA_CADUCIDAD,IDNUMCARGA, PDF_CARNE)
                VALUES (:matricula, :tipo, :numero, :comentarios, TO_DATE(:emision, 'DD/MM/YYYY'), TO_DATE(:caducidad, 'DD/MM/YYYY'), :idnumcarga, :pdf)
                """
                cursor1.execute(insert_carnet, {'matricula': matricula, 'tipo': 'CARNET DE SUCAMEC', 'numero': num_carnet,
                                                'comentarios': nombre_cliente_1, 'emision': fec_emis_carnet,
                                                'caducidad': fec_cadu_carnet, 'idnumcarga': new_idnumcarga,
                                                'pdf': pdf_read})
                print("Insertar datos de Carné")
                time.sleep(5)

                # Ejecutar procedimiento almacenado para 'CARNET DE SUCAMEC' (tipo 2)
                #call_proc = """
                #BEGIN
                #    UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC(
                #        p_codcia => :codcia,
                #        p_idnumcarga => :idnumcarga,
                #        p_tipo => :tipo,
                #        p_codigousuario_crea => :codigousuario_crea,
                #        p_fecha_crea => :fecha_crea,
                #        p_codigousuario_modifica => :codigousuario_modifica,
                #        p_fecha_modifica => :fecha_modifica
                #    );
                #END;
                #"""
                
                sql = "UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC"
                cursor1.callproc(sql, ['01',new_idnumcarga,'2',codigo_usuario,fecha_actual,codigo_usuario,fecha_actual])
                db_conn.commit()
                print("Sp ejecutado para  tipo 2")
            print("Transaction committed")

            if fec_emis_curso and fechaVencimiento_curso:
                # Insertar datos de Curso (sin PDF)
                # Obtener el máximo IDNUMCARGA y sumarle 1
                select_max_id_query = "SELECT NVL(MAX(IDNUMCARGA), 0) + 1 FROM uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC"
                cursor1.execute(select_max_id_query)
                new_idnumcarga = cursor1.fetchone()[0]

                insert_curso = """
                INSERT INTO uadryan.CARGA_TRABAJADOR_DETA_SUCAMEC (MATRICULA, IDTIPO, COMENTARIOS, FECHA_EMISION, FECHA_CADUCIDAD,IDNUMCARGA )
                VALUES (:matricula, :tipo, :comentarios, TO_DATE(:emision, 'DD/MM/YYYY'), TO_DATE(:caducidad, 'DD/MM/YYYY'), :idnumcarga)
                """
                cursor1.execute(insert_curso, {'matricula': matricula, 'tipo': 'CURSO SUCAMEC', 'comentarios': nombre_cliente_1,
                                               'emision': fec_emis_curso, 'caducidad': fechaVencimiento_curso,'idnumcarga': new_idnumcarga })
                print("Insertar datos de Curso")
                db_conn.commit()

                # Ejecutar procedimiento almacenado para 'CURSO SUCAMEC' (tipo 3)
                #call_proc = """
                #BEGIN
                #    UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC(
                #        p_codcia => :codcia,
                #        p_idnumcarga => :idnumcarga,
                #        p_tipo => :tipo,
                #        p_codigousuario_crea => :codigousuario_crea,
                #        p_fecha_crea => :fecha_crea,
                #        p_codigousuario_modifica => :codigousuario_modifica,
                #        p_fecha_modifica => :fecha_modifica
                #    );
                #END;
                #"""
                sql = "UADRYAN.PKG_G4S_DATA_ADIC.SP_REGI_DETA_SUCAMEC"
                cursor1.callproc(sql, ['01',new_idnumcarga,'3',codigo_usuario,fecha_actual,codigo_usuario,fecha_actual])
                db_conn.commit()
                print("Sp ejecutado para  tipo 3")
            print("Transaction committed")


            # Acceder a la página web
            driver.get("https://www.sucamec.gob.pe/sel/faces/aplicacion/consultas/DcGsspTotVigilantes/List.xhtml")
            print("pagina")
            time.sleep(5)
            time.sleep(5)
                
        except NoSuchElementException:
            print(f"No data found for documento: {documento}")
            continue



# Cerrar cursor y conexión a la base de datos
cursor1.close()
db_conn.close() 
print("Database cerrado")
print("----------------------------------------------------------------------------------")
print('PROCESO TERMINADO')