import argparse
import sys
import pandas as pd
import re
import time
import random
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

# Configurar argparse para tomar la ruta del archivo desde la línea de comandos
parser = argparse.ArgumentParser(description='Script para validar modelos y números de serie en la página de HP.')
parser.add_argument('file_path', type=str, help='Ruta del archivo Excel con modelos y números de serie.')
parser.add_argument('sheet_name', type=str, help='Nombre de la hoja en el archivo Excel.')
parser.add_argument('--start_serie', type=str, help='Número de serie desde el cual empezar la validación', default=None)
args = parser.parse_args()

# Verificar que el archivo existe
file_path = args.file_path
if not os.path.isfile(file_path):
    sys.stderr.write(f"Error: El archivo '{file_path}' no existe en esa ubicacion.\n")
    sys.exit(1)

# Leer el archivo Excel y la hoja especificada ej. 'ECATEPEC'
sheet_name = args.sheet_name
df = pd.read_excel(file_path, sheet_name=sheet_name)  # DataFrame

# Verificar que las columnas 'MODELO', 'SERIE' existen en el DataFrame y columna en la posicion R
required_columns = ['MODELO', 'SERIE']
if not all(column in df.columns for column in required_columns) or len(df.columns) <= 17:
    sys.stderr.write(f"Error: Las columnas {required_columns} deben estar presentes en el archivo Excel y tener MARCA en la columna R.\n")
    sys.exit(1)

# Configurar opciones de Chrome
chrome_options = Options()
# Cambiar el user agent
chrome_options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")

# Configurar el WebDriver con las opciones (asegúrate de tener chromedriver en tu PATH)
driver = webdriver.Chrome(options=chrome_options)

# Función para verificar si todos los términos del modelo están en el texto del <h1>
def validar_modelo(modelo, texto_h1):
    # Dividir el modelo en términos individuales
    palabras_modelo = modelo.split()
    # Comprobar que cada palabra del modelo está presente en el texto del <h1>
    for palabra in palabras_modelo:
        # Usar re.IGNORECASE para no distinguir entre mayúsculas y minúsculas
        if not re.search(re.escape(palabra), texto_h1, re.IGNORECASE):
            return False
    return True

# Función para agregar retrasos aleatorios (para que el servidor no bloquee nuestras peticiones)
def espera_aleatoria(min_s=1, max_s=3):
    tiempo_espera = random.uniform(min_s, max_s)
    time.sleep(tiempo_espera)

# Variables de conteo
total_correctos = 0
total_incorrectos = 0
total_no_validados = 0
total_numeros_serie_invalidos = 0
series_no_validadas = []
series_incorrectas = []
valores_adecuados = []
numeros_serie_invalidos = []
series_no_hp = []

# Manejar cookies
def manejar_cookies():
    # Guardar cookies en una variable
    cookies = driver.get_cookies()
    return cookies

def cargar_cookies(cookies):
    for cookie in cookies:
        driver.add_cookie(cookie)

# Encontrar el índice de inicio basado en el número de serie opcional
start_index = 0

if args.start_serie:
    try:
        start_index = df.index[df['SERIE'] == args.start_serie].tolist()[0]
    except IndexError:
        print(f"No se encontró el número de serie '{args.start_serie}', comenzando desde el inicio.")
        start_index = 0

row_index = start_index + 2

try:
    # Manejar cookies para evitar bloqueos por sesión
    cookies_guardadas = None
    
    for index, row in df.iloc[start_index:].iterrows():
        modelo = row['MODELO']
        serie = row['SERIE']
        marca = row.iloc[17]  # Columna en la posición R (índice 17)

        # Verificar la marca antes de proceder
        if marca != 'HP':
            print(f"Esta serie: {serie} no es de la marca HP")
            series_no_hp.append(serie)
            continue  # Saltar a la siguiente fila

        # Paso 1: Abrir la página de búsqueda
        driver.get("https://support.hp.com/mx-es/drivers/laptops")
        if cookies_guardadas:
            cargar_cookies(cookies_guardadas)

        print(f"Validando numero de serie: {serie} fila: {row_index}")

        # Esperar tiempo aleatorio
        espera_aleatoria()

        # Paso 2: Ingresar el número de serie, el programa espera hasta que el elemento del HTML con ID "searchQueryField" se carga.
        search_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "searchQueryField"))
        )
        search_field.clear() # limpiamos lo que haya en el input 
        search_field.send_keys(serie) # ingresamos el numero de serie 

        # Paso 3: Hacer clic en el botón de búsqueda
        espera_aleatoria()
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "FindMyProduct"))
        )
        search_button.click()

         # Paso 4: Esperar y verificar si hay un mensaje de número de serie inválido
        try:
            invalid_message_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "not-valid-search-text-common"))
            )
            if invalid_message_element.is_displayed():
                print(f"Número de serie inválido: {serie} en la fila {row_index}. Saltando a la siguiente fila.")
                total_numeros_serie_invalidos += 1
                numeros_serie_invalidos.append(serie)
                continue  # Saltar a la siguiente fila
        except Exception as e:
            pass  # No se encontró el mensaje de error, continuar con el flujo normal

        # Paso 5: Esperar a que se cargue la nueva página y manejar el stale element
        wait = WebDriverWait(driver, 60)  # Incrementamos el tiempo de espera a 60 segundos
        wait.until(EC.url_changes("https://support.hp.com/mx-es/drivers/laptops"))

        # Intentar obtener el texto del <h1> con manejo de `StaleElementReferenceException`
        actual_text = ""
        retries = 15  # Número de intentos
        for i in range(retries):
            try:
                h1_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "product-name-text"))
                )
                actual_text = h1_element.text.strip()
                if actual_text:  # Si obtenemos texto, rompemos el ciclo
                    break
            except StaleElementReferenceException:
                print("Elemento obsoleto, reintentando...")

            time.sleep(2)  # Espera 2 segundos antes de intentar nuevamente
            print(f"Esperando el texto del elemento <h1>... Intento {i+1} de {retries}")

        # Si se agotaron los intentos y no obtuvimos texto, saltamos la fila actual
        if not actual_text:
            print(f"No se pudo validar la serie: {serie} de la fila: {row_index} reintentos alcanzados: {retries}")
            series_no_validadas.append(serie)
            total_no_validados += 1

            continue # Ir a la siguiente fila
        
        # Verificar si el modelo está en el texto del <h1>
        if validar_modelo(modelo, actual_text):
            total_correctos += 1
            print(f"Numero de serie: {serie}. Validación exitosa: El modelo '{modelo}' concuerda con '{actual_text}'.")
        else:
            total_incorrectos += 1
            series_incorrectas.append(serie)
            valores_adecuados.append(actual_text)
            print(f"Numero de serie: {serie}. Validación fallida: El modelo '{modelo}' no concuerda con '{actual_text}'.")
        
        cookies_guardadas = manejar_cookies()
        row_index += 1

except Exception as e:
    print(f"Ocurrió un error: {e}")

finally:
    # Cerrar el WebDriver
    driver.quit()
    print(f"Total correctos: {total_correctos}")
    print(f"Total incorrectos: {total_incorrectos}")
    
    print(f"Total no validados: {total_no_validados}")
    print(f"Series no validadas: {series_no_validadas}")

    print(f"Total de numeros de serie invalidos: {total_numeros_serie_invalidos}")
    print(f"Numeros de serie invalidos: {numeros_serie_invalidos}")

    print(f"Total series que no son HP {len(series_no_hp)}")
    print(f"Imprimiendo series que no son HP: {series_no_hp}")
    
    print("Imprimiendo los numeros de serie que no corresponden al modelo adecuado...")
    for index, serie in enumerate(series_incorrectas):
        print(f"Numero de serie: {serie} Modelo en pagina HP: {valores_adecuados[index]}")