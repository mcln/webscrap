import os
import time
import requests
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import platform

# Función para reproducir un sonido (funciona en Windows)
def reproducir_sonido():
    if platform.system() == "Windows":
        import winsound
        duration = 1000  # milisegundos (1 segundo)
        freq = 440  # frecuencia en Hertz (Hz)
        winsound.Beep(freq, duration)
    else:
        # Reproduce un sonido en otros sistemas operativos (si es posible)
        os.system('say "Tarea completada"')  # Mac OS
        os.system('beep')  # Linux con soporte

# Función para hacer que la pantalla parpadee
def parpadeo_pantalla(segundos):
    for _ in range(segundos * 2):  # Dos cambios por segundo
        if platform.system() == "Windows":
            os.system("cls")
        else:
            os.system("clear")
        time.sleep(0.25)  # Tiempo de espera para el parpadeo
        print("Proceso completado. Los datos se han guardado.")

# Cargar los códigos desde un archivo Excel (ubicado en C:\xampp\htdocs\python\webscrap\input)
input_path = r'C:\xampp\htdocs\python\webscrap\input\codigos567890123456.xlsx'
codigos = []
wb = openpyxl.load_workbook(input_path)
ws = wb.active

# Asumiendo que los códigos están en la primera columna (A)
for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):  # Ajustar min_row si hay encabezado
    codigo = row[0]
    # Verificar si el código es válido (no None y no una cadena vacía)
    if codigo is not None and str(codigo).strip():
        codigos.append(codigo)

# Configuración de Selenium
driver = webdriver.Chrome()  # Asegúrate de tener el driver configurado
wait = WebDriverWait(driver, 10)  # Espera máxima de 10 segundos
url_base = "https://dosestrellas.cl/busquedas?buscar="

# Crear el archivo Excel de salida (ubicado en C:\xampp\htdocs\python\webscrap\output)
output_path = r'C:\xampp\htdocs\python\webscrap\output\productos567890123456.xlsx'

# Lista para acumular los resultados
rows = []

# Procesar cada código
for codigo in codigos:
    driver.get(url_base + str(codigo))

    # Obtener el link del producto
    try:
        producto = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".productos-grilla .item a.headerProduct")))
        link_producto = producto.get_attribute("href")
        driver.get(link_producto)
    except:
        print(f"No se encontró el producto para el código {codigo}")
        rows.append({
            "Código": codigo,
            "Link Producto": "No disponible",
            "Título": "No disponible",
            "Descripción": "No disponible",
            "Imagen": "Sin imagen"
        })
        continue

    # Captura de información
    try:
        # Esperar y capturar el título
        titulo = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".name"))).text

        # Esperar y capturar la descripción
        descripcion = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".tabs .body"))).text

        # Verificar si hay imagen
        try:
            imagen = driver.find_element(By.CSS_SELECTOR, "img.zoomImg").get_attribute("src")
            img_filename = f"C:\\xampp\\htdocs\\python\\webscrap\\output\\img\\{codigo}.jpg"

            # Verificar si la imagen ya existe
            if not os.path.exists(img_filename):
                # Descargar imagen solo si no existe
                response = requests.get(imagen)
                if response.status_code == 200:
                    with open(img_filename, 'wb') as img_file:
                        img_file.write(response.content)
            else:
                print(f"La imagen para el código {codigo} ya existe. Se omite la descarga.")
        except:
            img_filename = "Sin imagen"

        # Agregar datos a la lista
        rows.append({
            "Código": codigo,
            "Link Producto": link_producto,
            "Título": titulo,
            "Descripción": descripcion,
            "Imagen": img_filename
        })

    except Exception as e:
        print(f"Error al procesar el código {codigo}: {e}")
        rows.append({
            "Código": codigo,
            "Link Producto": "No disponible",
            "Título": "No disponible",
            "Descripción": "No disponible",
            "Imagen": "Sin imagen"
        })
        continue

# Convertir la lista de resultados en un DataFrame
df = pd.DataFrame(rows)

# Guardar los resultados en un archivo Excel en la carpeta de salida
df.to_excel(output_path, index=False)

# Cerrar el navegador
driver.quit()

# Mensaje de finalización
print(f"Proceso completado. Los datos se han guardado en {output_path}")

# Llamar a las funciones de alerta
reproducir_sonido()
parpadeo_pantalla(10)


