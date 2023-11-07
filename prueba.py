import os
import pytesseract
from PIL import Image
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# Función para leer texto de una imagen
def ocr_image(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang='spa')
        print(f"Imagen {image_path} procesada.")  # Muestra un mensaje cuando la imagen se ha procesado
        return text
    except Exception as e:
        print(f"Error procesando {image_path}: {e}")
        return None

# Función para procesar una imagen y almacenar los resultados
def process_image(filename):
    if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
        image_path = os.path.join(image_folder, filename)
        text = ocr_image(image_path)
        return (filename, text)

# Directorio donde están las imágenes
image_folder = '/home/andre/Escritorio/imagenes'

# Lista para guardar los datos
data = []

# Usar ThreadPoolExecutor para procesar las imágenes en paralelo
with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
    # Prepara todas las tareas que quieres ejecutar en paralelo
    futures = [executor.submit(process_image, filename) for filename in os.listdir(image_folder)]
    # A medida que cada tarea se completa, procesa los resultados
    for future in futures:
        result = future.result()
        if result:
            data.append(result)

# Crea un DataFrame de pandas con los datos
df = pd.DataFrame(data, columns=['Nombre de archivo', 'Texto'])

# Guarda el DataFrame en un archivo Excel
excel_path = '/home/andre/Escritorio/output.xlsx'
df.to_excel(excel_path, index=False)

print(f"El texto de las imágenes ha sido guardado en {excel_path}")
