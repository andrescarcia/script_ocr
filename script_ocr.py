import os
import pytesseract
from PIL import Image
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# Función para leer texto de una imagen
def ocr_image(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang='spa')
        return text
    except Exception as e:
        print(f"Error procesando {image_path}: {e}")
        return None

# Solicita al usuario la ruta de la carpeta de imágenes y la ruta de salida del archivo Excel
image_folder = input("Por favor, introduce la ruta de la carpeta que contiene las imágenes: ")
excel_path = input("Por favor, introduce la ruta donde deseas guardar el archivo de Excel: ")

# Lista para guardar los datos
data = []

# Usar ThreadPoolExecutor para procesar las imágenes en paralelo
with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
    # Prepara todas las tareas que quieres ejecutar en paralelo
    futures = []
    for filename in os.listdir(image_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
            futures.append(executor.submit(ocr_image, os.path.join(image_folder, filename)))

    # Procesa los resultados a medida que cada tarea se completa
    for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures), desc="Procesando imágenes"):
        result = future.result()
        if result:
            data.append(result)

# Crea un DataFrame de pandas con los datos
df = pd.DataFrame(data, columns=['Nombre de archivo', 'Texto'])

# Guarda el DataFrame en un archivo Excel
df.to_excel(excel_path, index=False)

print(f"El texto de las imágenes ha sido guardado en {excel_path}")
