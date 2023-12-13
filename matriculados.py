from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time

# Ruta
chrome_path = 'C:/Users/nando/Downloads/INSTALADORES/chorme driver/version 120.0.6099.71/chromedriver-win64/chromedriver.120.0.6099.71.exe'
# Configuración de Selenium
service = Service(chrome_path)
driver = webdriver.Chrome(service=service)

# URL de pagina
url_pagina_matriculacion = 'https://posgrado.unmsm.edu.pe/unidad_idiomas/registro_idiomas/index.php'
# Navega a la página de matriculación
driver.get(url_pagina_matriculacion)
# tiempo de espera
driver.implicitly_wait(10)

# Leer datos de Excel
excel_path = 'C:/Users/nando/PycharmProjects/BaseDatosMatriculados.xlsx'  # Reemplaza con la ruta correcta
alumnos_df = pd.read_excel(excel_path)
# iteracion del excel
for index, alumno in alumnos_df.iterrows():
    dni = str(alumno['DNI'])
    fechanacimiento = str(alumno['FechaNacimiento'])
    ap_paterno = str(alumno['ApellidoPaterno'])
    ap_materno = str(alumno['ApellidoMaterno'])
    nombres = str(alumno['Nombres'])
    sexo = str(alumno['Sexo'])
    tipo_documento = str(alumno['TipoDocumento'])
    fecha_nacimiento = str(alumno['FechaNacimiento'])
    correo_electronico = str(alumno['CorreoElectronico'])
    celular = str(alumno['Celular'])
    pais_origen = str(alumno['PaisOrigen'])
    nombre_curso = str(alumno['NombreCurso'])
    nombre_banco = str(alumno['NombreBanco'])
    numero_operacion = str(alumno['NumeroOperacion'])
    monto = str(alumno['Monto'])
    fecha_pago = str(alumno['FechaPago'])

    #####################PRIMERA SECCION########################################
    # Llenar los campos de la primera sección

    dni_input = driver.find_element(By.NAME, 'dni')
    dni_input.send_keys(dni)

    fechanacimiento_input = driver.find_element(By.NAME, 'fechanacimiento')
    fechanacimiento_input.send_keys(fechanacimiento)
    # boton de dar click automatico
    driver.find_element(By.XPATH, '//button[text()="Continuar"]').click()
    # tiempo de espera
    driver.implicitly_wait(5)

    #####################SEGUNDA SECCION########################################

    if not driver.find_elements(By.XPATH, '//span[contains(text(), "El estudiante ya se encuentra matriculado")]'):
        print('Hola')
        ap_paterno_input = driver.find_element(By.NAME, 'paterno')
        ap_paterno_input.send_keys(ap_paterno)

        ap_materno_input = driver.find_element(By.NAME, 'materno')
        ap_materno_input.send_keys(ap_materno)

        nombres_input = driver.find_element(By.NAME, 'nombres')
        nombres_input.send_keys(nombres)

        sexo_select = driver.find_element(By.NAME, 'sexo')
        sexo_select.send_keys(sexo)

        tipo_documento_select = driver.find_element(By.NAME, 'tipodocumento')
        tipo_documento_select.send_keys(tipo_documento)

        driver.find_element(By.XPATH, '//button[text()="Registrar"]').click()
        driver.implicitly_wait(5)

        #####################TERCER SECCION########################################
        # Llenar los campos de la tercera sección
        correo_input = driver.find_element(By.NAME, 'correoelectronico')
        correo_input.send_keys(correo_electronico)

        celular_input = driver.find_element(By.NAME, 'celular')
        celular_input.send_keys(celular)

        pais_origen_select = driver.find_element(By.NAME, 'pais')
        pais_origen_select.send_keys(pais_origen)

        # botón "Registrar"
        driver.find_element(By.XPATH, '//button[text()="Registrar"]').click()

        # Esperar un momento para que la página se actualice
        driver.implicitly_wait(50)

    #####################CUARTE SECCION########################################
    # llenar el curso deseado
    curso_select = driver.find_element(By.NAME, 'codCurso')
    curso_select.send_keys(nombre_curso)

    # Llenar detalles de pago
    banco_select = driver.find_element(By.NAME, 'banco')
    banco_select.send_keys(nombre_banco)

    operacion_input = driver.find_element(By.NAME, 'operacion')
    operacion_input.send_keys(numero_operacion)

    monto_select = driver.find_element(By.NAME, 'monto')
    monto_select.send_keys(monto)  # Puedes cambiar el monto según tus necesidades

    fecha_pago_input = driver.find_element(By.NAME, 'fechapago')
    fecha_pago_input.send_keys(fecha_pago)

    # botón "Registrar" para finalizar registro
    # driver.find_element(By.XPATH, '//button[text()="Realizar mi matrícula"]').click()

    # tiempo de espera entre cada iteración
    time.sleep(5)

# Esperar un momento para asegurarse de que la página se actualice
driver.implicitly_wait(100)

# Esperar un tiempo antes de cerrar la pestaña
time.sleep(100)
