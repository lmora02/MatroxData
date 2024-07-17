import os
import threading
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl.styles import PatternFill
import shutil
import subprocess
from PIL import Image, ImageTk  # Importar Image y ImageTk de Pillow
 
 
 # Obtener la ruta del directorio actual donde se encuentra el script
directorio_actual = os.path.dirname(os.path.abspath(__file__))
 
 
# Variable global para almacenar la carpeta seleccionada
carpeta_seleccionada = None
ruta_archivo_estadisticos = None  # Variable global para almacenar la ruta del archivo de estadísticos
 
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
            valores = []  # Lista para almacenar los valores de Recipe ID, Exposure Time, Image Time Stamp, Blob 1 Threshold, Blob 1 Min, Blob 1 Max y Blob 1 Area
            for subcarpeta_txt in subcarpetas_txt:
                for root, _, archivos in os.walk(subcarpeta_txt):
                    for archivo in archivos:
                        if archivo.endswith('.txt'):
                            # Inicializar el diccionario de datos para cada archivo
                            datos = {}                            
                            # Aquí se puede implementar la lógica para procesar cada archivo .txt
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
                                    camara=archivo[:19]
                                    datos = {'Camara':camara,'Archivo': archivo, 'Recipe ID': recipe_id, 'Exposure Time': exposure_time, 'Image Time Stamp': image_time_stamp}  # No se agrega 'Camara' aquí
                                    for blob_key, blob_info in blob_data.items():
                                        if blob_info['Enabled']:
                                            datos.update({f'{blob_key} Threshold': blob_info['Threshold'],
                                                          f'{blob_key} Min': blob_info['Min'],
                                                          f'{blob_key} Max': blob_info['Max'],
                                                          f'{blob_key} Area': blob_info['Area']})
                                    #datos['Camara'] = archivo[:19]  # Agregar 'Camara' con los primeros 19 caracteres de 'Archivo' aquí
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
                            datos[parametro].append(linea.strip().split(':')[1])
 
        df = pd.DataFrame(datos)
        ruta_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
 
        if ruta_excel:
            df.to_excel(ruta_excel, index=False)
            mensaje = f"Se ha generado el archivo Excel en:\n{ruta_excel}"
            messagebox.showinfo("Excel Generado", mensaje)
            abrir_excel_button = ttk.Button(ventana_estadisticos_especificos, text="Abrir archivo de Excel", command=lambda: abrir_archivo_excel(ruta_excel))
            abrir_excel_button.pack(pady=5)
 
            # Guardar la ruta del archivo generado para abrir la carpeta contenedora
            ventana_estadisticos_especificos.ruta_excel_generado = ruta_excel
 
    # Función para abrir el archivo de Excel generado
    def abrir_archivo_excel(ruta_excel):
        os.system(f'start excel "{ruta_excel}"')
 
    # Crear una nueva ventana para los estadísticos de datos específicos
    ventana_estadisticos_especificos = tk.Toplevel(root)
    ventana_estadisticos_especificos.title("Estadísticos Datos Específicos")
    ventana_estadisticos_especificos.geometry("450x500")  # Establecer el tamaño inicial de la ventana
    ventana_estadisticos_especificos.resizable(False, False)  # Evitar que la ventana se redimensione
 
    # Estilo de Material Design
    style = ttk.Style()
    style.theme_use("clam")  # Elige el estilo Material Design
    style.configure("TLabel", foreground="black", background="#f0f0f0", font=('Roboto', 10))
    style.configure("TButton", padding=10, relief="flat", background="#3f51b5", foreground="white", font=('Roboto', 8, 'bold'), width=30)
    style.map("TButton", background=[('active', '#283593')])
 
    # # Combinar la ruta del directorio actual con el nombre de la imagen
    ruta_imagen_matrox = os.path.join(directorio_actual, "matrox.png")
    # Cargar imagen Matrox en la ventana secundaria
    #imagen_matrox = tk.PhotoImage(file=ruta_imagen_matrox)
 
    # Widget para mostrar la imagen Matrox en la ventana secundaria
    #matrox_label = tk.Label(ventana_estadisticos_especificos, image=matrox)
    #matrox_label.pack()
 
    # Combinar la ruta del directorio actual con el nombre de la imagen
    ruta_imagen_vista = os.path.join(directorio_actual, "vista.png")
    # Cargar imagen Vista en la ventana secundaria
    imagen_vista = tk.PhotoImage(file=ruta_imagen_vista)
   
 
    # Widget para mostrar la imagen Vista en la ventana secundaria
    #imagen_label = tk.Label(ventana_estadisticos_especificos, image=imagen_vista)
    #imagen_label.pack()
 
    # Etiqueta para el texto de instrucción
    etiqueta_instruccion = ttk.Label(ventana_estadisticos_especificos, text="Escribe los parametros de inspeccion que deseas obtener\n Si ingresas más de un parametro separalo por un ; sin espacios\n Ejemplo IP Adress;PVI;PUN\n\n Al dar clic en el boton Procesar Seleccionar la carpeta\n donde se desean buscar los archivos puede ser\n alguna en especifico o de toda la camara", style="TLabel")
    etiqueta_instruccion.pack()
 
    # Campo de texto para ingresar los parámetros
    entrada_parametros = ttk.Entry(ventana_estadisticos_especificos, width=50)
    entrada_parametros.pack()
 
    # Botón para procesar los parámetros
    boton_procesar = ttk.Button(ventana_estadisticos_especificos, text="Procesar", command=procesar_parametros, style="TButton")
    boton_procesar.pack()
 
# Función para abrir la ventana de obtener datos de la cámara
def obtener_datos_camara():
    global monitor_running, ventana_archivos, conjunto_ip, ventana_ip
    monitor_running = True
    conjunto_ip = None  # Variable para almacenar las IPs seleccionadas desde el Excel
    ventana_ip = None  # Inicializar ventana_ip como None

    def procesar_direccion_ip(direccion_ip=None):
        global direccion_ip_global, monitor_conexion, monitor_running, ventana_ip

        if not direccion_ip:
            direccion_ip = entrada_ip.get()  # Obtener la dirección IP ingresada desde el campo de entrada

        if direccion_ip:
            # Realizar prueba de ping para verificar la conexión
            ping_exit_code = subprocess.call(['ping', '-n', '1', direccion_ip], stdout=subprocess.DEVNULL)
            if ping_exit_code == 0:  # Ping exitoso
                try:
                    # Intentar establecer conexión SMB a través de la ventana de comandos de Windows
                    comando = f"net use \\\\{direccion_ip}\\IPC$ /user:NAM\\mtxuser Matrox"
                    subprocess.run(comando, shell=True, check=True)
                    print(f"Conexión SMB establecida con {direccion_ip}")

                    # Mostrar mensaje de autenticación exitosa
                    messagebox.showinfo("Autenticación Exitosa", "Autenticación exitosa con la dirección IP especificada.")

                    # Ocultar la ventana para ingresar la IP si está definida
                    if ventana_ip:
                        ventana_ip.destroy()
                    # Abrir ventana para seleccionar archivos a copiar (JPG, PNG, TXT)
                    abrir_ventana_seleccion_archivos()
                    # Mostrar ventana con la dirección IP y estado de la conexión
                    mostrar_ventana_estado(direccion_ip)

                    # Iniciar monitoreo de conexión
                    monitor_running = True
                    monitor_conexion = threading.Thread(target=monitorizar_conexion, args=(direccion_ip,))
                    monitor_conexion.start()

                except subprocess.CalledProcessError as e:
                    # Mostrar mensaje de error si no se pudo establecer la conexión SMB
                    messagebox.showerror("Error", f"No se pudo establecer la conexión SMB con {direccion_ip}: {str(e)}")
                    # Abrir ventana para ingresar nuevas credenciales
                    abrir_ventana_credenciales(direccion_ip)
            else:
                # Si no se puede hacer ping, mostrar mensaje de error
                messagebox.showerror("Error", f"No se pudo hacer ping a la dirección IP {direccion_ip}.")
        else:
            # Si no se ingresó ninguna dirección IP, muestra un mensaje de error
            messagebox.showerror("Error", "Debes ingresar una dirección IP válida.")

    def mostrar_ventana_estado(direccion_ip):
        global direccion_ip_global, ventana_estado, estado_label
        direccion_ip_global = direccion_ip

        def cerrar_conexion():
            global monitor_running
            monitor_running = False
            comando_desconectar = f"net use \\\\{direccion_ip}\\IPC$ /delete"
            subprocess.run(comando_desconectar, shell=True, check=True)
            print(f"Conexión SMB cerrada con {direccion_ip}")
            ventana_estado.destroy()
            if ventana_archivos:
                ventana_archivos.destroy()
            messagebox.showinfo("Conexión Cerrada", f"Conexión cerrada con {direccion_ip}.")

        def on_closing():
            cerrar_conexion()
            messagebox.showinfo("Conexion Perdida", f"Se ha perdido la conexión con {direccion_ip}.")

        def actualizar_estado(estado):
            estado_label.config(text=f"Dirección IP: {direccion_ip}\nEstado: {estado}")

        ventana_estado = tk.Toplevel(root)
        ventana_estado.title("Estado de Conexión")
        ventana_estado.geometry("300x150")
        ventana_estado.resizable(False, False)

        estado_label = ttk.Label(ventana_estado, text=f"Dirección IP: {direccion_ip}\nEstado: Conectado")
        estado_label.pack(pady=20)

        boton_cerrar = ttk.Button(ventana_estado, text="Cerrar Conexión", command=cerrar_conexion)
        boton_cerrar.pack(pady=10)

        ventana_estado.protocol("WM_DELETE_WINDOW", on_closing)

    def monitorizar_conexion(direccion_ip):
        global direccion_ip_global, estado_label

        while monitor_running:
            ping_exit_code = subprocess.call(['ping', '-n', '1', direccion_ip], stdout=subprocess.DEVNULL)
            if ping_exit_code != 0:
                if direccion_ip == direccion_ip_global:
                    estado_label.config(text=f"Dirección IP: {direccion_ip}\nEstado: Desconectado")
                    messagebox.showwarning("Conexión Perdida", f"Se ha perdido la conexión con {direccion_ip}.")
                    ventana_estado.destroy()
                    if ventana_archivos:
                        ventana_archivos.destroy()
                    break

            time.sleep(5)

    def abrir_ventana_seleccion_archivos():
        global ventana_archivos
        ventana_archivos = tk.Toplevel(root)
        ventana_archivos.title("Seleccionar Archivos a Copiar")
        ventana_archivos.geometry("300x300")
        ventana_archivos.resizable(False, False)

        etiqueta_instrucciones = ttk.Label(ventana_archivos, text="Selecciona los archivos a copiar:")
        etiqueta_instrucciones.pack(pady=10)

        var_jpg = tk.IntVar()
        check_jpg = ttk.Checkbutton(ventana_archivos, text="JPG", variable=var_jpg)
        check_jpg.pack()

        var_png = tk.IntVar()
        check_png = ttk.Checkbutton(ventana_archivos, text="PNG", variable=var_png)
        check_png.pack()

        var_txt = tk.IntVar()
        check_txt = ttk.Checkbutton(ventana_archivos, text="TXT", variable=var_txt)
        check_txt.pack()

        boton_extraer = ttk.Button(ventana_archivos, text="Extraer Archivos",
                                   command=lambda: extraer_archivos(var_jpg, var_png, var_txt))
        boton_extraer.pack(pady=10)

        boton_cerrar = ttk.Button(ventana_archivos, text="Cerrar", command=ventana_archivos.destroy)
        boton_cerrar.pack(pady=10)

    def extraer_archivos(var_jpg, var_png, var_txt):
        carpeta_destino = filedialog.askdirectory(title="Selecciona la carpeta de destino")
        if not carpeta_destino:
            return

        extensiones_seleccionadas = []
        if var_jpg.get():
            extensiones_seleccionadas.append(".jpg")
        if var_png.get():
            extensiones_seleccionadas.append(".png")
        if var_txt.get():
            extensiones_seleccionadas.append(".txt")

        if not extensiones_seleccionadas:
            messagebox.showwarning("Advertencia", "No se ha seleccionado ningún tipo de archivo para copiar.")
            return

        ventana_progreso = tk.Toplevel(root)
        ventana_progreso.title("Progreso de Copia de Archivos")
        ventana_progreso.geometry("300x100")
        ventana_progreso.resizable(False, False)

        etiqueta_progreso = ttk.Label(ventana_progreso, text="Copiando archivos...")
        etiqueta_progreso.pack(pady=10)

        barra_progreso = ttk.Progressbar(ventana_progreso, orient="horizontal", length=200, mode="determinate")
        barra_progreso.pack(pady=10)

        total_archivos = 0
        if conjunto_ip:
            for ip in conjunto_ip:
                ruta_origen = f"\\\\{ip}\\mtxuser"
                total_archivos += sum(1 for _, _, archivos in os.walk(ruta_origen) if any(
                    archivo.lower().endswith(ext) for ext in extensiones_seleccionadas for archivo in archivos))
        else:
            ruta_origen = f"\\\\{direccion_ip_global}\\mtxuser"
            total_archivos = sum(1 for _, _, archivos in os.walk(ruta_origen) if any(
                archivo.lower().endswith(ext) for ext in extensiones_seleccionadas for archivo in archivos))

        barra_progreso["maximum"] = total_archivos

        def copiar_archivos(ip, carpeta_destino, extensiones_seleccionadas):
            ruta_origen = f"\\\\{ip}\\mtxuser"
            for raiz, dirs, archivos in os.walk(ruta_origen):
                for archivo in archivos:
                    if any(archivo.lower().endswith(ext) for ext in extensiones_seleccionadas):
                        ruta_completa_origen = os.path.join(raiz, archivo)
                        ruta_completa_destino = os.path.join(carpeta_destino, archivo)
                        try:
                            shutil.copy2(ruta_completa_origen, ruta_completa_destino)
                            print(f"Archivo copiado: {ruta_completa_origen} -> {ruta_completa_destino}")
                            barra_progreso["value"] += 1
                            root.update_idletasks()
                        except Exception as e:
                            messagebox.showerror("Error", f"No se pudo copiar el archivo {archivo}: {str(e)}")

        if conjunto_ip:
            for ip in conjunto_ip:
                copiar_archivos(ip, carpeta_destino, extensiones_seleccionadas)
        else:
            copiar_archivos(direccion_ip_global, carpeta_destino, extensiones_seleccionadas)

        messagebox.showinfo("Proceso Completo", "Se han copiado todos los archivos seleccionados.")
        ventana_progreso.destroy()

    def abrir_ventana_credenciales(direccion_ip):
        global usuario_global, contrasena_global
        ventana_credenciales = tk.Toplevel(root)
        ventana_credenciales.title("Ingresar Credenciales")
        ventana_credenciales.geometry("300x200")
        ventana_credenciales.resizable(False, False)

        etiqueta_usuario = ttk.Label(ventana_credenciales, text="Usuario:")
        etiqueta_usuario.pack(pady=10)
        entrada_usuario = ttk.Entry(ventana_credenciales)
        entrada_usuario.pack()

        etiqueta_contrasena = ttk.Label(ventana_credenciales, text="Contraseña:")
        etiqueta_contrasena.pack(pady=10)
        entrada_contrasena = ttk.Entry(ventana_credenciales, show="*")
        entrada_contrasena.pack()

        def guardar_credenciales():
            global usuario_global, contrasena_global
            usuario_global = entrada_usuario.get()
            contrasena_global = entrada_contrasena.get()
            ventana_credenciales.destroy()
            autenticar_conexion(direccion_ip, usuario_global, contrasena_global)

        boton_guardar = ttk.Button(ventana_credenciales, text="Guardar", command=guardar_credenciales)
        boton_guardar.pack(pady=20)

    def autenticar_conexion(direccion_ip, usuario, contrasena):
        global direccion_ip_global, monitor_conexion, monitor_running
        try:
            comando = f"net use \\\\{direccion_ip}\\IPC$ /user:{usuario} {contrasena}"
            subprocess.run(comando, shell=True, check=True)
            print(f"Conexión SMB establecida con {direccion_ip}")

            # Mostrar mensaje de autenticación exitosa
            messagebox.showinfo("Autenticación Exitosa", "Autenticación exitosa con la dirección IP especificada.")

            # Abrir ventana para seleccionar archivos a copiar (JPG, PNG, TXT)
            abrir_ventana_seleccion_archivos()

            # Mostrar ventana con la dirección IP y estado de la conexión
            mostrar_ventana_estado(direccion_ip)

            # Iniciar monitoreo de conexión
            monitor_running = True
            monitor_conexion = threading.Thread(target=monitorizar_conexion, args=(direccion_ip,))
            monitor_conexion.start()

        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"No se pudo establecer la conexión SMB con {direccion_ip}: {str(e)}")
            abrir_ventana_credenciales(direccion_ip)

    def mostrar_ventana_ip():
        global entrada_ip, ventana_ip
        ventana_ip = tk.Toplevel(root)
        ventana_ip.title("Ingrese la Dirección IP")
        ventana_ip.geometry("300x150")
        ventana_ip.resizable(False, False)

        etiqueta_ip = ttk.Label(ventana_ip, text="Dirección IP:")
        etiqueta_ip.pack(pady=10)

        entrada_ip = ttk.Entry(ventana_ip)
        entrada_ip.pack()

        boton_procesar = ttk.Button(ventana_ip, text="Procesar", command=procesar_direccion_ip)
        boton_procesar.pack(pady=10)

    def extraer_ips_desde_excel():
        global conjunto_ip
        archivo_excel = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
        if not archivo_excel:
            return

        df = pd.read_excel(archivo_excel)
        if 'IP' not in df.columns:
            messagebox.showerror("Error", "El archivo de Excel no contiene una columna 'IP'.")
            return

        conjunto_ip = df['IP'].dropna().tolist()

        for ip in conjunto_ip:
            procesar_direccion_ip(ip)

    def mostrar_menu():
        menu = tk.Menu(root)
        root.config(menu=menu)

        menu_archivo = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Archivo", menu=menu_archivo)
        menu_archivo.add_command(label="Ingresar IP", command=mostrar_ventana_ip)
        menu_archivo.add_command(label="Extraer direcciones IP desde un excel", command=extraer_ips_desde_excel)
        menu_archivo.add_separator()
        menu_archivo.add_command(label="Salir", command=root.quit)

    mostrar_menu()





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
encabezado_label = ttk.Label(root, text="—————————| FUNCIONES ESPECIALES |————————————")
encabezado_label.pack(pady=10, anchor="w", padx=15)

# Botón para generar estadísticos de datos específicos
estadisticos_especificos_button = ttk.Button(root, text="Estadísticos Datos Específicos", command=generar_estadisticos_datos_especificos, style="TButton")
estadisticos_especificos_button.pack(pady=5, anchor="w", padx=115)

# Etiqueta para mostrar mensajes
mensaje_label = ttk.Label(root, text="", style="TLabel")
mensaje_label.pack()

# Etiqueta para mostrar la ruta del archivo de estadísticos
etiqueta_ruta = ttk.Label(root, text="", style="TLabel")
etiqueta_ruta.place(x=1000, y=1000)

# Botón para obtener datos de la cámara con estilo personalizado
boton_obtener_datos = ttk.Button(root, text="Obtener datos de la cámara", command=obtener_datos_camara, style="TButton")
boton_obtener_datos.pack(pady=5, anchor="w", padx=115)

# Iniciar el bucle de eventos de la interfaz gráfica
root.mainloop()

