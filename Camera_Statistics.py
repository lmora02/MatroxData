"""
Camera Statistics Script
========================

This script provides a utility to manage and process files within a specified directory. The main features include:
- Classifying files based on their extensions.
- Generating statistical summaries.
- Handling image files.

Modules Imported:
-----------------
- os: Provides functions to interact with the operating system.
- pandas: Used for data manipulation and analysis.
- tkinter: Provides classes for creating graphical user interfaces.
- openpyxl.styles: Used for styling Excel files.
- shutil: Offers a number of high-level operations on files and collections of files.
- subprocess: Allows spawning new processes, connecting to their input/output/error pipes, and obtaining their return codes.
- PIL (Pillow): Adds image processing capabilities to Python.

Global Variables:
-----------------
- directorio_actual: Stores the path of the current directory where the script is located.
- carpeta_seleccionada: Global variable to store the selected folder.
- ruta_archivo_estadisticos: Global variable to store the path of the statistical file.

Functions:
----------
- clasificar_archivos(carpeta_principal)
    Classifies files into subfolders based on their extensions.

    Parameters:
    - carpeta_principal (str): The main directory to classify files.

- seleccionar_carpeta_principal()
    Opens a file dialog to select the main folder and calls the classification function.

- buscar_subcarpetas_txt(carpeta)
    Searches for all subfolders named 'TXT' within the specified directory.

    Parameters:
    - carpeta (str): The directory to search for 'TXT' subfolders.

- generar_estadisticos()
    Generates statistics from text files in the selected folder using file dialogs.

- abrir_estadisticos()
    Opens the generated statistics file if available.

- actualizar_etiqueta_ruta()
    Updates the GUI label with the path of the statistics file.

- generar_estadisticos_datos_especificos()
    Opens a new window to specify parameters and generate specific statistics.

- obtener_datos_camara()
    Opens a window to enter IP address for camera connection and file selection.

"""
import os
import threading
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl.styles import PatternFill
from datetime import datetime
import shutil
import subprocess
from PIL import Image, ImageTk

# Obtener la ruta del directorio actual donde se encuentra el script
directorio_actual = os.path.dirname(os.path.abspath(__file__))

# Variable global para almacenar la carpeta seleccionada
carpeta_seleccionada = None
ruta_archivo_estadisticos = None  # Variable global para almacenar la ruta del archivo de estadísticos
idioma = "ES"

# Función para clasificar archivos en subcarpetas según su extensión
def clasificar_archivos(carpeta_principal):
    carpetas = ['png', 'jpg', 'txt']
    # Recorrer la estructura de directorios de la carpeta principal
    for carpeta, _, archivos in os.walk(carpeta_principal):
        for archivo in archivos:
            archivo_path = os.path.join(carpeta, archivo)
            # Obtener la extensión del archivo
            extension = archivo.split('.')[-1].lower()
            # Verificar si la extensión está en la lista de extensiones a clasificar
            if extension in carpetas:
                # Crear la subcarpeta si no existe
                subcarpeta = os.path.join(carpeta, extension.lower())
                if not os.path.exists(subcarpeta):
                    os.makedirs(subcarpeta)
                # Mover el archivo a la subcarpeta correspondiente
                shutil.move(archivo_path, os.path.join(subcarpeta, archivo))

# Función para seleccionar una carpeta principal y clasificar sus archivos
def seleccionar_carpeta_principal():
    carpeta_principal = filedialog.askdirectory()
    if carpeta_principal:
        clasificar_archivos(carpeta_principal)
        mensaje_label.config(text="Archivos clasificados correctamente.")

# Función para buscar todas las subcarpetas llamadas 'TXT'
def buscar_subcarpetas_txt(carpeta):
    subcarpetas_txt = []
    for root, dirs, _ in os.walk(carpeta):
        for dir_name in dirs:
            if dir_name.lower() == 'txt':
                subcarpetas_txt.append(os.path.join(root, dir_name))
    return subcarpetas_txt

# Función para generar estadísticas de archivos de texto en la carpeta seleccionada
def generar_estadisticos():
    global carpeta_seleccionada, ruta_archivo_estadisticos
    # Seleccionar una carpeta
    carpeta_seleccionada = filedialog.askdirectory(title="Seleccionar Carpeta")
    if carpeta_seleccionada:
        # Buscar todas las subcarpetas llamadas 'TXT'
        subcarpetas_txt = buscar_subcarpetas_txt(carpeta_seleccionada)
        if subcarpetas_txt:
            # Mostrar mensaje informativo y procesar archivos .txt
            messagebox.showinfo("Generar Estadísticos", f"Se generarán estadísticos de la carpeta: {carpeta_seleccionada}")
            valores = []
            for subcarpeta_txt in subcarpetas_txt:
                for root, _, archivos in os.walk(subcarpeta_txt):
                    for archivo in archivos:
                        if archivo.endswith('.txt'):
                            datos = {}
                            ruta_archivo = os.path.join(root, archivo)
                            with open(ruta_archivo, 'r') as file:
                                contenido = file.readlines()
 
                                recipe_id = None
                                exposure_time = None
                                image_time_stamp = None  # Agregar para almacenar el Image Time Stamp
                                blob_data = {}
                                for i in range(1, 7):
                                    blob_key = f'Blob {i}'
                                    blob_data[blob_key] = {
                                        'Enabled': False,
                                        'Threshold': None,
                                        'Min': None,
                                        'Max': None,
                                        'Area': None
                                    }
 
                                for line in contenido:
                                    if "Recipe ID" in line:
                                        recipe_id = int(line.split(': ')[1])
                                    elif "Exposure Time" in line:
                                        exposure_time = int(line.split(': ')[1])
                                    elif "Image Time Stamp" in line:  # Buscar el Image Time Stamp
                                        image_time_stamp = line.split(': ')[1].strip()
                                    elif "Blob" in line and "Enabled: True" in line:
                                        blob_number = int(line.split()[1])
                                        blob_data[f'Blob {blob_number}']['Enabled'] = True
                                    elif "Blob" in line and "Threshold" in line:
                                        blob_number = int(line.split()[1])
                                        blob_data[f'Blob {blob_number}']['Threshold'] = int(line.split(': ')[1])
                                    elif "Blob" in line and "Min" in line:
                                        blob_number = int(line.split()[1])
                                        blob_data[f'Blob {blob_number}']['Min'] = int(line.split(': ')[1])
                                    elif "Blob" in line and "Max" in line:
                                        blob_number = int(line.split()[1])
                                        blob_data[f'Blob {blob_number}']['Max'] = int(line.split(': ')[1])
                                    elif "Blob" in line and "Area" in line:
                                        blob_number = int(line.split()[1])
                                        blob_data[f'Blob {blob_number}']['Area'] = int(line.split(': ')[1])
 
                                if recipe_id is not None and exposure_time is not None and image_time_stamp is not None:
                                    camara = archivo[:19]
                                    datos = {'Camara': camara, 'Archivo': archivo, 'Recipe ID': recipe_id, 'Exposure Time': exposure_time, 'Image Time Stamp': image_time_stamp}
                                    for blob_key, blob_info in blob_data.items():
                                        if blob_info['Enabled']:
                                            datos.update({
                                                f'{blob_key} Threshold': blob_info['Threshold'],
                                                f'{blob_key} Min': blob_info['Min'],
                                                f'{blob_key} Max': blob_info['Max'],
                                                f'{blob_key} Area': blob_info['Area']
                                            })
                                    valores.append(datos)
 
            # Crear un DataFrame de Pandas con los valores de Recipe ID, Exposure Time, Image Time Stamp, Blob 1 Threshold, Blob 1 Min, Blob 1 Max y Blob 1 Area
            df = pd.DataFrame(valores)
            # Obtener el nombre de la carpeta seleccionada
            nombre_carpeta = os.path.basename(carpeta_seleccionada)
 
            # Guardar el DataFrame en un archivo de Excel con el nombre de la carpeta en la raíz de la carpeta seleccionada
            ruta_excel = os.path.join(carpeta_seleccionada, f'{nombre_carpeta}.xlsx')
 
            # Agregar la columna con los nombres de los archivos .txt
            df['Archivo'] = [os.path.basename(path) for path in df['Archivo']]
 
            # Escribir el DataFrame en el archivo Excel con filtro en la primera fila
            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=True)
                # Ajustar automáticamente el tamaño de las columnas al contenido
                for column in writer.sheets['Sheet1'].columns:
                    max_length = 0
                    column = list(column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    writer.sheets['Sheet1'].column_dimensions[column[0].column_letter].width = adjusted_width
 
                # Resaltar en rojo las columnas que tienen campos vacíos
                for col in writer.sheets['Sheet1'].iter_cols():
                    for cell in col:
                        if cell.value is None or cell.value == '':
                            cell.fill = PatternFill(start_color="F08080", end_color="F08080", fill_type="solid")
 
            # Almacenar la ruta del archivo generado
            ruta_archivo_estadisticos = ruta_excel
            actualizar_etiqueta_ruta()
 
            messagebox.showinfo("Generar Estadísticos", f"Estadísticos generados y guardados en '{ruta_excel}'")
        else:
            # Mostrar mensaje de error si no se encuentra ninguna subcarpeta 'TXT'
            messagebox.showerror("Error", "No se encontraron subcarpetas 'TXT' en la carpeta seleccionada.")
 
 
# Función para abrir el archivo de estadísticos
def abrir_estadisticos():
    global ruta_archivo_estadisticos
    if ruta_archivo_estadisticos:
        os.startfile(ruta_archivo_estadisticos)
    else:
        messagebox.showinfo("Abrir Estadísticos", "Primero genera los estadísticos para abrir el archivo.")
 
# Función para actualizar la etiqueta con la ruta del archivo de estadísticos
def actualizar_etiqueta_ruta():
    if ruta_archivo_estadisticos:
        etiqueta_ruta.config(text=f"Ruta del archivo de estadísticos: {ruta_archivo_estadisticos}")
    else:
        etiqueta_ruta.config(text="")
 
# Función para generar estadísticos de datos específicos
def generar_estadisticos_datos_especificos():
    # Función para procesar los parámetros ingresados y buscar el texto en los archivos .txt
    def procesar_parametros():
        parametros = ["Recipe ID", "Exposure Time", "Image Time Stamp"] + entrada_parametros.get().split(';')
        carpeta_seleccionada = filedialog.askdirectory(title="Seleccionar Carpeta")
 
        if not parametros:
            messagebox.showerror("Error", "Por favor ingresa al menos un parámetro.")
            return
 
        if not carpeta_seleccionada:
            messagebox.showerror("Error", "Debes seleccionar una carpeta.")
            return
 
        archivos_encontrados = []
        nombres_archivos = []  # Lista para almacenar los nombres de los archivos .txt encontrados
        camaras_encontradas =[]
        for root, _, archivos in os.walk(carpeta_seleccionada):
            for archivo in archivos:
                if archivo.endswith('.txt'):
                    ruta_archivo = os.path.join(root, archivo)
                    with open(ruta_archivo, 'r') as file:
                        contenido = file.read()
                        if all(parametro in contenido for parametro in parametros):
                            archivos_encontrados.append(ruta_archivo)
                            nombres_archivos.append(archivo)
                            camaras_encontradas.append(archivo[:19])
 
        if not archivos_encontrados:
            messagebox.showinfo("Resultado", "No se encontraron coincidencias.")
            return
 
        generar_excel(archivos_encontrados, nombres_archivos, parametros, carpeta_seleccionada, camaras_encontradas)
 
    # Función para generar un archivo Excel con los datos de los archivos encontrados
    def generar_excel(archivos, nombres_archivos, parametros, carpeta_seleccionada, camaras_encontradas):
                             
        datos = {'Camaras': camaras_encontradas, 'Archivo': nombres_archivos}  # Inicializar el diccionario con los nombres de los archivos .txt
        for parametro in parametros:
            datos[parametro] = []
 
        for archivo in archivos:
            with open(archivo, 'r') as file:
                contenido = file.readlines()
                for linea in contenido:
                    for parametro in parametros:
                        if parametro in linea:
                            valor = linea.split(': ')[1].strip()
                            datos[parametro].append(valor)

        nombre_carpeta = os.path.basename(carpeta_seleccionada)
        df = pd.DataFrame(datos)
        ruta_excel = os.path.join(carpeta_seleccionada, f'{nombre_carpeta}_datos_especificos.xlsx')
        df.to_excel(ruta_excel, index=False, header=True)

        for column in df.columns:
            max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            df[column] = df[column].apply(lambda x: x.ljust(max_length))

        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=True)
            for column in writer.sheets['Sheet1'].columns:
                max_length = 0
                column = list(column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                writer.sheets['Sheet1'].column_dimensions[column[0].column_letter].width = adjusted_width

        messagebox.showinfo("Resultado", f"Estadísticos generados y guardados en '{ruta_excel}'")

    ventana_parametros = tk.Toplevel()
    ventana_parametros.title("Generar Estadísticos de Datos Específicos")
    ventana_parametros.geometry("400x200")

    etiqueta_parametros = ttk.Label(ventana_parametros, text="Parámetros (separados por ';'):")
    etiqueta_parametros.pack(pady=10)

    entrada_parametros = ttk.Entry(ventana_parametros, width=50)
    entrada_parametros.pack(pady=10)

    boton_procesar = ttk.Button(ventana_parametros, text="Procesar", command=procesar_parametros)
    boton_procesar.pack(pady=10)

# Función para cambiar el idioma de la aplicación
def cambiar_idioma():
    global idioma
    if idioma == 'EN':
        idioma = 'ES'
    else:
        idioma = 'EN'
    actualizar_texto_elementos()

# Función para actualizar el texto de los elementos de la interfaz según el idioma seleccionado
def actualizar_texto_elementos():
    if idioma == 'EN':
        root.title("Data Statistics Generator")
        seleccionar_button.config(text="Classify Files")
        estadisticos_button.config(text="Generate Statistics")
        abrir_estadisticos_button.config(text="Open Statistics")
        estadisticos_especificos_button.config(text="Specific Data Statistics")
        boton_cambiar_idioma.config(text="Spanish")
        mensaje_label.config(text="Select a main folder to classify files.")
        encabezado_label.config(text="—————————————| FUNCTIONS |——————————————")
        encabezadoSPECIAL_label.config(text="—————————| SPECIAL FUNCTIONS |————————————")
        instrucciones_label.config(text="———————————| INSTRUCTIONS |—————————————\n\n1. Select the folder to select the files to be classified.\n2. Depending on the data of interest, click on 'Generate\n    Stadistics' or if you need any other information click on the button\n    'Specific Data Statistics'.")
    else:
        root.title("Generador de Estadísticas de Datos")
        seleccionar_button.config(text="Clasificar Archivos")
        estadisticos_button.config(text="Generar Estadísticos")
        abrir_estadisticos_button.config(text="Abrir Estadísticos")
        estadisticos_especificos_button.config(text="Estadísticos Datos Específicos")
        boton_cambiar_idioma.config(text="Inglés")
        mensaje_label.config(text="Selecciona una carpeta principal para clasificar archivos.")
        encabezado_label.config(text="—————————————| FUNCIONES |——————————————")
        encabezadoSPECIAL_label.config(text="—————————| FUNCIONES ESPECIALES |————————————")
        instrucciones_label.config(text="———————————| INSTRUCCIONES |—————————————\n\n1. Seleccionar la carpeta para seleccionar los archivos a clasificar.\n2. Dependiendo de los datos de interes, dar click en el boton de 'Generar\n    Estadistico' o si se necesita algún otro dato dar click en el botón\n    'Estadisticos Datos Especificos'.")

# Configuración de la ventana principal
root = tk.Tk()
root.title("Inspection Tools Statistics V1.01")
root.geometry("450x500")  # Establecer el tamaño inicial de la ventana
root.resizable(False, False)  # Evitar que la ventana se redimensione

# Estilo de Material Design
style = ttk.Style()
style.theme_use("clam")  # Elige el estilo Material Design
style.configure("TLabel", foreground="black", background="#f0f0f0", font=('Roboto', 10))
style.configure("TButton", padding=10, relief="flat", background="#3f51b5", foreground="white", font=('Roboto', 8, 'bold'), width=30)
style.map("TButton", background=[('active', '#283593')])

# Nuevo estilo para los botones en la ventana principal con letra más pequeña
style.configure("BotonesVentanaPrincipal.TButton", font=('Roboto', 7, 'bold'))

# Etiqueta para mostrar las instrucciones
instrucciones_label = ttk.Label(root, text="———————————| INSTRUCCIONES |—————————————\n\n1. Seleccionar la carpeta para seleccionar los archivos a clasificar.\n2. Dependiendo de los datos de interes, dar click en el boton de 'Generar\n    Estadistico' o si se necesita algún otro dato dar click en el botón\n    'Estadisticos Datos Especificos'.")
instrucciones_label.place(x=10, y=380)

# Combinar la ruta del directorio actual con el nombre de la imagen
ruta_imagen_matrox = os.path.join(directorio_actual, "matrox.png")
# Cargar imagen Matrox en la ventana principal
imagen_matrox = tk.PhotoImage(file=ruta_imagen_matrox)

# Widget para mostrar la imagen Matrox en la ventana principal
matrox_label = tk.Label(root, image=imagen_matrox)
matrox_label.place(x=390, y=5)

# Combinar la ruta del directorio actual con el nombre de la imagen
ruta_imagen_vista = os.path.join(directorio_actual, "vista.png")
# Cargar imagen Vista en la ventana principal
imagen_vista = tk.PhotoImage(file=ruta_imagen_vista)

# Widget para mostrar la imagen Vista en la ventana principal
imagen_label = tk.Label(root, image=imagen_vista)
imagen_label.pack(side="top", padx=20, pady=15, anchor="nw")

# Etiqueta para mostrar las encabezado de los botones
encabezado_label = ttk.Label(root, text="—————————————| FUNCIONES |——————————————")
encabezado_label.pack(pady=10, anchor="w", padx=15)

# Botón para generar estadísticos
estadisticos_button = ttk.Button(root, text="Generar Estadísticos Blobs", command=generar_estadisticos, style="TButton")
estadisticos_button.place(x=230, y=107)

# Botón para clasificar
seleccionar_button = ttk.Button(root, text="Clasificar Extensiones", command=seleccionar_carpeta_principal, style="TButton")
seleccionar_button.pack(pady=10, anchor="w", padx=15)

# Botón para abrir el archivo de estadísticos
abrir_estadisticos_button = ttk.Button(root, text="Abrir Estadísticos", command=abrir_estadisticos, style="TButton")
abrir_estadisticos_button.pack(pady=5, anchor="w", padx=115)

# Etiqueta para mostrar las encabezado de los botones
encabezadoSPECIAL_label = ttk.Label(root, text="—————————| FUNCIONES ESPECIALES |————————————")
encabezadoSPECIAL_label.pack(pady=10, anchor="w", padx=15)

# Botón para generar estadísticos de datos específicos
estadisticos_especificos_button = ttk.Button(root, text="Estadísticos Datos Específicos", command=generar_estadisticos_datos_especificos, style="TButton")
estadisticos_especificos_button.pack(pady=5, anchor="w", padx=115)

# Etiqueta para mostrar mensajes
mensaje_label = ttk.Label(root, text="", style="TLabel")
mensaje_label.pack()

# Etiqueta para mostrar la ruta del archivo de estadísticos
etiqueta_ruta = ttk.Label(root, text="", style="TLabel")
etiqueta_ruta.place(x=1000, y=1000)

# Botón para cambiar el idioma de la aplicación
boton_cambiar_idioma = ttk.Button(root, text="Inglés", command=cambiar_idioma, style="TButton", width=10)
boton_cambiar_idioma.pack(pady=25, anchor="se", padx=10)

# Mostrar la ventana principal
root.mainloop()

#Test to commit#Test to branch tacos de canasta
