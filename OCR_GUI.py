import os
import threading
from PIL import Image
import pandas as pd
from tkinter import Tk, Label, Button, Entry, StringVar, Text, Scrollbar, filedialog
import pytesseract
from concurrent.futures import ThreadPoolExecutor

# Función para leer texto de una imagen
def ocr_image(image_path, text_widget):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang='spa')
        text_widget.insert('end', f"Procesado: {image_path}\n")
        text_widget.see('end')  # Auto-scroll
        return os.path.basename(image_path), text
    except Exception as e:
        text_widget.insert('end', f"Error procesando {image_path}: {e}\n")
        text_widget.see('end')  # Auto-scroll
        return os.path.basename(image_path), None

# Función para iniciar el proceso de OCR en un hilo separado
def start_ocr_thread(image_folder, excel_path, text_widget):
    ocr_thread = threading.Thread(target=start_ocr_process, args=(image_folder, excel_path, text_widget))
    ocr_thread.start()

# Función para iniciar el proceso de OCR
def start_ocr_process(image_folder, excel_path, text_widget):
    if not image_folder or not excel_path:
        text_widget.insert('end', "Por favor, selecciona la carpeta de imágenes y la ruta del archivo de Excel.\n")
        return

    data = []

    with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
        futures = [executor.submit(ocr_image, os.path.join(image_folder, filename), text_widget)
                   for filename in os.listdir(image_folder)
                   if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif'))]

        for future in futures:
            image_path, result = future.result()
            if result:
               # result = str(result)
                data.append((image_path, result))

    # Crea un DataFrame de pandas con los datos
    df = pd.DataFrame(data, columns=['Nombre de archivo', 'Texto'])
    # Intenta guardar el DataFrame en un archivo Excel
    try:
        df.to_excel(excel_path,  engine='xlsxwriter')
        text_widget.insert('end', f"El texto de las imágenes ha sido guardado en {excel_path}\n")
    except Exception as e:
        text_widget.insert('end', f"Error al guardar el archivo de Excel: {e}\n")


# Función para seleccionar la carpeta de imágenes
def select_image_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.set(folder_selected)

# Función para seleccionar la ruta del archivo Excel
def select_excel_path(entry):
    file_selected = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry.set(file_selected)

# Crear la ventana principal de la aplicación
root = Tk()
root.title("OCR App")

# Variables para almacenar las rutas de entrada y salida
image_folder_path = StringVar()
excel_file_path = StringVar()

# Crear y colocar los widgets
Label(root, text="Carpeta de Imágenes:").grid(row=0, column=0, sticky='e')
Entry(root, textvariable=image_folder_path, width=50).grid(row=0, column=1)
Button(root, text="Seleccionar", command=lambda: select_image_folder(image_folder_path)).grid(row=0, column=2)

Label(root, text="Guardar Excel en:").grid(row=1, column=0, sticky='e')
Entry(root, textvariable=excel_file_path, width=50).grid(row=1, column=1)
Button(root, text="Seleccionar", command=lambda: select_excel_path(excel_file_path)).grid(row=1, column=2)

Button(root, text="Iniciar OCR", command=lambda: start_ocr_thread(image_folder_path.get(), excel_file_path.get(), text_widget)).grid(row=2, column=0, columnspan=3)

# Área de texto para mostrar el progreso
text_widget = Text(root, height=10, width=80)
scrollbar = Scrollbar(root, command=text_widget.yview)
text_widget.configure(yscrollcommand=scrollbar.set)
text_widget.grid(row=3, column=0, columnspan=2)
scrollbar.grid(row=3, column=2, sticky='nsew')

# Ejecutar la aplicación
root.mainloop()