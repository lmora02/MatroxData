Español

Script de Estadísticas de Cámara

Este script proporciona una utilidad para gestionar y procesar archivos dentro de un directorio especificado. Las principales características incluyen la clasificación de archivos por extensiones, la generación de resúmenes estadísticos y el manejo de archivos de imagen.

Módulos Importados

  os: Proporciona funciones para interactuar con el sistema operativo.

  pandas: Utilizado para la manipulación y análisis de datos.

  tkinter: Proporciona clases para la creación de interfaces gráficas de usuario.

  openpyxl.styles: Utilizado para el estilo de archivos Excel.

  shutil: Ofrece varias operaciones de alto nivel en archivos y colecciones de archivos.

  subprocess: Permite iniciar nuevos procesos, conectar con sus tuberías de entrada/salida/error, y obtener sus códigos de retorno.

  PIL (Pillow): Añade capacidades de procesamiento de imágenes a Python.

Variables Globales

  directorio_actual: Almacena la ruta del directorio actual donde se encuentra el script.

  carpeta_seleccionada: Variable global para almacenar la carpeta seleccionada por el usuario.

  ruta_archivo_estadisticos: Variable global para almacenar la ruta del archivo de estadísticas generado.

  idioma: Idioma actual de la interfaz (por defecto: español).


Funciones Principales

  clasificar_archivos(carpeta_principal): Clasifica los archivos en subcarpetas según su extensión.

  seleccionar_carpeta_principal(): Abre un cuadro de diálogo para seleccionar la carpeta principal y clasificar los archivos.

  generar_estadisticos(): Genera estadísticas a partir de archivos de texto en la carpeta seleccionada.

  abrir_estadisticos(): Abre el archivo de estadísticas generado, si está disponible.

  generar_estadisticos_datos_especificos(): Abre una nueva ventana para especificar parámetros y generar estadísticas específicas.


Interfaz de Usuario (GUI)

El script utiliza tkinter para construir una interfaz gráfica de usuario que incluye botones para ejecutar las funciones principales y especiales, etiquetas para mostrar mensajes e información, y opciones para cambiar el idioma entre español e inglés.


Uso

1. Ejecuta el script.
2. Selecciona una carpeta para clasificar archivos o generar estadísticas específicas.
3. Utiliza los botones proporcionados para realizar las acciones deseadas.
4. Se mostrarán mensajes informativos y la ruta del archivo de estadísticas generado, si corresponde.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

English

Camera Statistics Script
This script provides a utility to manage and process files within a specified directory. The main features include classifying files by their extensions, generating statistical summaries, and handling image files.

Imported Modules

  os: Provides functions for interacting with the operating system.

  pandas: Used for data manipulation and analysis.

  tkinter: Provides classes for creating graphical user interfaces.

  openpyxl.styles: Used for styling Excel files.

  shutil: Offers a number of high-level operations on files and collections of files.

  subprocess: Allows spawning new processes, connecting to their input/output/error pipes, and obtaining their return codes.

  PIL (Pillow): Adds image processing capabilities to Python.

Global Variables

  directorio_actual: Stores the path of the current directory where the script is located.

  carpeta_seleccionada: Global variable to store the selected folder by the user.

  ruta_archivo_estadisticos: Global variable to store the path of the generated statistics file.

  idioma: Current language of the interface (default: Spanish).

Main Functions

  clasificar_archivos(carpeta_principal): Classifies files into subfolders based on their extensions.

  seleccionar_carpeta_principal(): Opens a file dialog to select the main folder and calls the classification function.

  generar_estadisticos(): Generates statistics from text files in the selected folder.

  abrir_estadisticos(): Opens the generated statistics file if available.

  generar_estadisticos_datos_especificos(): Opens a new window to specify parameters and generate specific statistics.


User Interface (GUI)

The script uses tkinter to build a graphical user interface that includes buttons to execute main and special functions, labels to display messages and information, and options to switch between Spanish and English languages.

Usage
1. Run the script.
2. Select a folder to classify files or generate specific statistics.
3. Use the provided buttons to perform desired actions.
4. Informative messages and the path to the generated statistics file, if applicable, will be displayed.
