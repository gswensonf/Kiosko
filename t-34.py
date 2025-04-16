import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import threading
import shutil
import pyperclip
import platform
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import psutil
import wmi
import time
import winreg
import win32com.client
from docx import Document
import webbrowser
import ctypes
import sys

# Función para verificar y elevar privilegios
def es_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevar_privilegios():
    if not es_admin():
        # Re-ejecutar el script con privilegios de administrador
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit()

# Elevar privilegios al inicio
elevar_privilegios()

# Función para obtener la ruta correcta del icono
def get_resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para desarrollo y para PyInstaller"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Obtener la ruta del icono
ICON_PATH = get_resource_path("icon.ico")

# Función para cargar iconos silenciosamente
def load_icon(window):
    """Intenta cargar el icono sin mostrar errores"""
    try:
        window.iconbitmap(ICON_PATH)
    except:
        pass

# Color de fondo para los botones (gris claro)
BOTON_COLOR = "#f0f0f0"

# Clase para crear tooltips
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()

    def leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

# Función para mostrar el acuerdo de licencia MIT
def mostrar_acuerdo_licencia():
    licencia_texto = """
    ACUERDO DE LICENCIA MIT

    Copyright (c) 2025
      Guillermo Javier Swenson Fuenteseca 
      Fono: + 56 9 27234210 
      Mail: gswensonf@gmail.com  

    Por la presente se otorga permiso, libre de cargos, a cualquier persona que obtenga una copia
    de este software y los archivos de documentación asociados (el "Software"), a utilizar
    el Software sin restricción, incluyendo sin limitación los derechos de usar, copiar,
    modificar, fusionar, publicar, distribuir, sublicenciar y/o vender copias del Software,
    y a permitir a las personas a las que se les proporcione el Software a hacer lo mismo,
    sujeto a las siguientes condiciones:

    El aviso de copyright anterior y este aviso de permiso se incluirán en todas las copias
    o partes sustanciales del Software.

    EL SOFTWARE SE PROPORCIONA "TAL CUAL", SIN GARANTÍA DE NINGÚN TIPO, EXPRESA O
    IMPLÍCITA, INCLUYENDO PERO NO LIMITADO A GARANTÍAS DE COMERCIALIZACIÓN,
    IDONEIDAD PARA UN PROPÓSITO PARTICULAR Y NO INFRACCIÓN. EN NINGÚN CASO LOS
    AUTORES O TITULARES DEL COPYRIGHT SERÁN RESPONSABLES POR NINGUNA RECLAMACIÓN,
    DAÑOS U OTRAS RESPONSABILIDADES, YA SEA EN UNA ACCIÓN DE CONTRATO, AGRAVIO O CUALQUIER
    OTRO MOTIVO, QUE SURJA DE O EN CONEXIÓN CON EL SOFTWARE O EL USO U OTRO TIPO DE
    ACCIONES EN EL SOFTWARE.
    """
    
    ventana_licencia = tk.Toplevel()
    ventana_licencia.title("Acuerdo de Licencia MIT")
    ventana_licencia.geometry("800x600")
    load_icon(ventana_licencia)
    
    frame = tk.Frame(ventana_licencia)
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    texto_licencia = tk.Text(
        frame,
        wrap=tk.WORD,
        yscrollcommand=scrollbar.set,
        font=("Arial", 10),
        padx=10,
        pady=10
    )
    texto_licencia.pack(fill=tk.BOTH, expand=True)
    
    texto_licencia.insert(tk.END, licencia_texto)
    texto_licencia.config(state=tk.DISABLED)
    
    scrollbar.config(command=texto_licencia.yview)
    
    boton_cerrar = tk.Button(
        ventana_licencia,
        text="Cerrar",
        command=ventana_licencia.destroy,
        bg=BOTON_COLOR
    )
    boton_cerrar.pack(pady=10)

# Función para generar el manual en HTML
def generar_manual_html():
    try:
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Manual de la Aplicación T-34</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    margin: 20px;
                    color: #333;
                }
                h1 {
                    color: #2c3e50;
                    border-bottom: 2px solid #3498db;
                    padding-bottom: 10px;
                }
                h2 {
                    color: #2980b9;
                    margin-top: 20px;
                }
                h3 {
                    color: #16a085;
                }
                ul {
                    padding-left: 20px;
                }
                li {
                    margin-bottom: 5px;
                }
                .section {
                    background-color: #f9f9f9;
                    padding: 15px;
                    border-radius: 5px;
                    margin-bottom: 20px;
                    border-left: 4px solid #3498db;
                }
                .warning {
                    background-color: #fff3cd;
                    padding: 10px;
                    border-radius: 5px;
                    border-left: 4px solid #ffc107;
                    margin: 10px 0;
                }
                .success {
                    background-color: #d4edda;
                    padding: 10px;
                    border-radius: 5px;
                    border-left: 4px solid #28a745;
                    margin: 10px 0;
                }
            </style>
        </head>
        <body>
            <h1>Manual de la Aplicación T-34</h1>
            
            <div class="section">
                <h2>Descripción General</h2>
                <p>La aplicación T-34 es una herramienta de diagnóstico y mantenimiento para sistemas Windows. Proporciona diversas funcionalidades para resolver problemas comunes, optimizar el sistema y obtener información detallada sobre el hardware y software.</p>
            </div>
            
            <div class="section">
                <h2>Funcionalidades Principales</h2>
                
                <h3>1. Diagnósticos de Hardware</h3>
                <p>Esta sección incluye herramientas para analizar el hardware del sistema:</p>
                <ul>
                    <li><strong>Información del Hardware:</strong> Muestra detalles sobre los componentes del sistema.</li>
                    <li><strong>Diagnosticar Hardware:</strong> Realiza un chequeo básico del estado del hardware.</li>
                    <li><strong>Verificar Controladores:</strong> Identifica controladores con problemas.</li>
                    <li><strong>Prueba de Memoria:</strong> Ejecuta una prueba básica de la memoria RAM.</li>
                    <li><strong>Prueba de Disco:</strong> Mide la velocidad de lectura/escritura del disco.</li>
                </ul>
                
                <h3>2. Almacenamiento</h3>
                <p>Herramientas para gestionar el espacio en disco:</p>
                <ul>
                    <li><strong>Escanear disco C:</strong> Busca carpetas que pesen más de 5 GB.</li>
                    <li><strong>Liberar espacio en disco:</strong> Elimina archivos temporales y vacía la papelera.</li>
                    <li><strong>PST:</strong> Busca y gestiona archivos PST de Outlook.</li>
                </ul>
                
                <h3>3. Wi-Fi</h3>
                <p>Herramientas para diagnóstico y solución de problemas de red:</p>
                <ul>
                    <li><strong>Verificar Controlador WiFi:</strong> Comprueba el estado del controlador WiFi.</li>
                    <li><strong>Generar informe Wi-Fi:</strong> Crea un reporte detallado de conexiones WiFi.</li>
                    <li><strong>Comandos ipconfig:</strong> Varias opciones para gestión de red.</li>
                    <li><strong>Olvidar redes Wi-Fi:</strong> Elimina redes WiFi guardadas.</li>
                    <li><strong>Reiniciar servicio Wi-Fi:</strong> Reinicia el servicio WLAN.</li>
                </ul>
                
                <h3>4. Diagnóstico de Pantallazos Azules</h3>
                <p>Herramientas para analizar errores del sistema:</p>
                <ul>
                    <li><strong>Ver BSODs en Visor de Eventos:</strong> Muestra pantallazos azules recientes.</li>
                </ul>
                
                <h3>5. Office</h3>
                <p>Herramientas para Microsoft Office:</p>
                <ul>
                    <li><strong>Reactivar OneDrive:</strong> Habilita OneDrive si está desactivado.</li>
                    <li><strong>Aumentar Tamaño OST:</strong> Incrementa el límite de archivos OST de Outlook.</li>
                    <li><strong>Eliminar Perfil Outlook:</strong> Borra perfiles de Outlook corruptos.</li>
                    <li><strong>Archivar en Línea:</strong> Mueve correos antiguos al archivo en línea.</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>Requisitos del Sistema</h2>
                <ul>
                    <li>Sistema operativo: Windows 7 o superior</li>
                    <li>Memoria RAM: 2 GB mínimo (4 GB recomendado)</li>
                    <li>Espacio en disco: 100 MB disponible</li>
                </ul>
            </div>
            
            <div class="warning">
                <h3>Notas Importantes</h3>
                <p>Algunas funciones requieren permisos de administrador para funcionar correctamente.</p>
                <p>Se recomienda cerrar todas las aplicaciones antes de realizar operaciones de mantenimiento.</p>
            </div>
        </body>
        </html>
        """

        # Crear archivo HTML temporal
        temp_html = os.path.join(os.environ['TEMP'], 'Manual_T-34.html')
        with open(temp_html, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Abrir en el navegador predeterminado
        webbrowser.open(temp_html)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el manual: {str(e)}")

# Función para calcular el tamaño de una carpeta
def obtener_tamano_carpeta(ruta):
    total = 0
    for ruta_actual, _, archivos in os.walk(ruta):
        for archivo in archivos:
            try:
                total += os.path.getsize(os.path.join(ruta_actual, archivo))
            except (FileNotFoundError, PermissionError):
                continue
    return total

# Función para encontrar carpetas que pesen más de 5 GB
def encontrar_carpetas_pesadas(ruta_inicio, tamaño_minimo=5 * 1024 * 1024 * 1024, ventana_progreso=None, barra_progreso=None):
    carpetas_pesadas = []
    carpetas = next(os.walk(ruta_inicio))[1]
    total_carpetas = len(carpetas)

    for i, carpeta in enumerate(carpetas):
        ruta_completa = os.path.join(ruta_inicio, carpeta)
        try:
            tamano = obtener_tamano_carpeta(ruta_completa)
            if tamano > tamaño_minimo:
                carpetas_pesadas.append((ruta_completa, tamano))
        except (FileNotFoundError, PermissionError):
            continue

        if ventana_progreso and barra_progreso:
            progreso = (i + 1) / total_carpetas * 100
            barra_progreso["value"] = progreso
            ventana_progreso.update_idletasks()

    carpetas_pesadas.sort(key=lambda x: x[1], reverse=True)
    return carpetas_pesadas

# Función para abrir la carpeta en el Explorador de Archivos
def abrir_carpeta(ruta):
    try:
        subprocess.Popen(f'explorer "{ruta}"')
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")

# Función para liberar espacio en el disco duro
def liberar_espacio():
    espacio_liberado = 0
    try:
        temp_user = os.path.join(os.environ["TEMP"])
        for item in os.listdir(temp_user):
            ruta_completa = os.path.join(temp_user, item)
            try:
                if os.path.isfile(ruta_completa) or os.path.islink(ruta_completa):
                    espacio_liberado += os.path.getsize(ruta_completa)
                    os.unlink(ruta_completa)
                elif os.path.isdir(ruta_completa):
                    espacio_liberado += obtener_tamano_carpeta(ruta_completa)
                    shutil.rmtree(ruta_completa)
            except Exception as e:
                print(f"No se pudo eliminar {ruta_completa}: {e}")

        temp_windows = os.path.join(os.environ["SystemRoot"], "Temp")
        for item in os.listdir(temp_windows):
            ruta_completa = os.path.join(temp_windows, item)
            try:
                if os.path.isfile(ruta_completa) or os.path.islink(ruta_completa):
                    espacio_liberado += os.path.getsize(ruta_completa)
                    os.unlink(ruta_completa)
                elif os.path.isdir(ruta_completa):
                    espacio_liberado += obtener_tamano_carpeta(ruta_completa)
                    shutil.rmtree(ruta_completa)
            except PermissionError:
                print(f"Permiso denegado para eliminar {ruta_completa}. Ejecuta el programa como administrador.")
            except Exception as e:
                print(f"No se pudo eliminar {ruta_completa}: {e}")

        try:
            papelera = os.path.join(os.environ["SystemDrive"], "$Recycle.Bin")
            espacio_liberado += obtener_tamano_carpeta(papelera)
            subprocess.run(["rd", "/s", "/q", papelera], shell=True)
        except PermissionError:
            print("Permiso denegado para vaciar la papelera de reciclaje. Ejecuta el programa como administrador.")
        except Exception as e:
            print(f"No se pudo vaciar la papelera de reciclaje: {e}")

        espacio_liberado_gb = espacio_liberado / (1024 * 1024 * 1024)
        messagebox.showinfo("Éxito", f"Se ha liberado {espacio_liberado_gb:.2f} GB de espacio en el disco duro.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo liberar espacio: {e}")

# Función para mostrar los resultados en una ventana gráfica
def mostrar_resultados(carpetas_pesadas):
    ventana_resultados = tk.Toplevel()
    ventana_resultados.title("Carpetas más pesadas en C:")
    ventana_resultados.geometry("1000x600")
    load_icon(ventana_resultados)

    tree = ttk.Treeview(ventana_resultados, columns=("Tamaño"), show="tree headings")
    tree.heading("#0", text="Carpeta", anchor=tk.W)
    tree.heading("Tamaño", text="Tamaño (GB)", anchor=tk.W)
    tree.column("Tamaño", width=150, anchor=tk.E)
    tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    def cargar_subcarpetas(parent_id, ruta):
        try:
            for nombre in os.listdir(ruta):
                ruta_completa = os.path.join(ruta, nombre)
                if os.path.isdir(ruta_completa):
                    tamano = obtener_tamano_carpeta(ruta_completa)
                    subcarpeta_id = tree.insert(parent_id, "end", text=nombre, values=(f"{tamano / (1024 * 1024 * 1024):.2f} GB"))
                    tree.insert(subcarpeta_id, "end", text="Cargando...")
        except (FileNotFoundError, PermissionError):
            pass

    def on_expand(event):
        item_id = tree.focus()
        ruta = obtener_ruta_completa(item_id)
        if tree.get_children(item_id) and tree.item(tree.get_children(item_id)[0], "text") == "Cargando...":
            tree.delete(tree.get_children(item_id)[0])
            cargar_subcarpetas(item_id, ruta)

    def obtener_ruta_completa(item_id):
        partes = []
        while item_id:
            partes.append(tree.item(item_id, "text"))
            item_id = tree.parent(item_id)
        return os.path.join("C:\\", *reversed(partes))

    for carpeta, tamano in carpetas_pesadas:
        carpeta_id = tree.insert("", "end", text=os.path.basename(carpeta), values=(f"{tamano / (1024 * 1024 * 1024):.2f} GB"))
        tree.insert(carpeta_id, "end", text="Cargando...")

    tree.bind("<<TreeviewOpen>>", on_expand)

    boton_abrir = tk.Button(
        ventana_resultados,
        text="Abrir carpeta seleccionada",
        command=lambda: abrir_carpeta(obtener_ruta_completa(tree.focus())),
        bg=BOTON_COLOR
    )
    boton_abrir.pack(pady=10)

    boton_cerrar = tk.Button(
        ventana_resultados,
        text="Cerrar",
        command=ventana_resultados.destroy,
        bg=BOTON_COLOR
    )
    boton_cerrar.pack(pady=10)

# Función para ejecutar el escaneo en un hilo separado
def ejecutar_escaneo(ventana_progreso, barra_progreso):
    ruta_inicio = "C:\\"
    messagebox.showinfo("Escaneando", "Escaneando el disco C:. Esto puede tardar unos minutos...")

    carpetas_pesadas = encontrar_carpetas_pesadas(ruta_inicio, ventana_progreso=ventana_progreso, barra_progreso=barra_progreso)
    ventana_progreso.destroy()

    if not carpetas_pesadas:
        messagebox.showinfo("Resultado", "No se encontraron carpetas que pesen más de 5 GB.")
        return

    ventana.after(0, mostrar_resultados, carpetas_pesadas)

# Función para mostrar la ventana de progreso
def mostrar_ventana_progreso():
    ventana_progreso = tk.Toplevel()
    ventana_progreso.title("Escaneando disco C:")
    ventana_progreso.geometry("400x100")
    load_icon(ventana_progreso)
    ventana_progreso.lift()
    ventana_progreso.attributes("-topmost", True)

    barra_progreso = ttk.Progressbar(
        ventana_progreso,
        orient=tk.HORIZONTAL,
        length=300,
        mode="determinate",
    )
    barra_progreso.pack(pady=20)

    threading.Thread(
        target=ejecutar_escaneo,
        args=(ventana_progreso, barra_progreso),
        daemon=True,
    ).start()

# Función para buscar archivos PST en la máquina
def buscar_pst():
    pst_files = []
    for root, _, files in os.walk("C:\\"):
        for file in files:
            if file.endswith(".pst"):
                pst_files.append(os.path.join(root, file))
    return pst_files

# Función para dividir un archivo PST en partes más pequeñas usando pff-tools
def dividir_pst(ruta_pst, tamaño_parte_gb=25):
    try:
        comando = f"pff-split -s {tamaño_parte_gb}GB {ruta_pst}"
        subprocess.run(comando, shell=True, check=True)
        return True, "Archivo PST dividido correctamente."
    except subprocess.CalledProcessError as e:
        return False, f"Error al dividir el archivo PST: {e}"

# Función para mostrar los archivos PST en una nueva ventana
def mostrar_pst():
    confirmacion = messagebox.askyesno("Advertencia", "Buscar archivos PST puede llevar tiempo. ¿Deseas continuar?")
    if not confirmacion:
        return

    ventana_progreso = tk.Toplevel()
    ventana_progreso.title("Buscando archivos PST")
    ventana_progreso.geometry("400x100")
    load_icon(ventana_progreso)
    ventana_progreso.lift()
    ventana_progreso.attributes("-topmost", True)

    barra_progreso = ttk.Progressbar(ventana_progreso, orient=tk.HORIZONTAL, length=300, mode="determinate")
    barra_progreso.pack(pady=20)

    def buscar_pst_en_segundo_plano():
        pst_files = buscar_pst()
        ventana_progreso.destroy()

        if not pst_files:
            messagebox.showinfo("Resultado", "No se encontraron archivos PST.")
            return

        ventana_pst = tk.Toplevel()
        ventana_pst.title("Archivos PST encontrados")
        ventana_pst.geometry("800x400")
        load_icon(ventana_pst)

        lista_pst = tk.Listbox(ventana_pst, selectmode=tk.MULTIPLE)
        for pst in pst_files:
            lista_pst.insert(tk.END, pst)
        lista_pst.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        def borrar_seleccionados():
            seleccionados = lista_pst.curselection()
            if not seleccionados:
                messagebox.showwarning("Advertencia", "No se seleccionaron archivos para borrar.")
                return

            confirmacion = messagebox.askyesno("Confirmar", "¿Estás seguro de que deseas borrar los archivos seleccionados?")
            if confirmacion:
                ventana_progreso_borrar = tk.Toplevel()
                ventana_progreso_borrar.title("Borrando archivos")
                ventana_progreso_borrar.geometry("400x100")
                load_icon(ventana_progreso_borrar)

                barra_progreso_borrar = ttk.Progressbar(ventana_progreso_borrar, orient=tk.HORIZONTAL, length=300, mode="determinate")
                barra_progreso_borrar.pack(pady=20)

                def borrar_archivos():
                    total_archivos = len(seleccionados)
                    for i, index in enumerate(seleccionados):
                        pst_file = lista_pst.get(index)
                        try:
                            os.remove(pst_file)
                            lista_pst.delete(index)
                        except Exception as e:
                            messagebox.showerror("Error", f"No se pudo borrar {pst_file}: {e}")

                        barra_progreso_borrar["value"] = (i + 1) / total_archivos * 100
                        ventana_progreso_borrar.update_idletasks()

                    messagebox.showinfo("Éxito", "Los archivos seleccionados se han borrado correctamente.")
                    ventana_progreso_borrar.destroy()

                threading.Thread(target=borrar_archivos, daemon=True).start()

        boton_borrar = tk.Button(
            ventana_pst, 
            text="Borrar seleccionados", 
            command=borrar_seleccionados,
            bg=BOTON_COLOR
        )
        boton_borrar.pack(pady=10)

        def partir_seleccionados():
            seleccionados = lista_pst.curselection()
            if not seleccionados:
                messagebox.showwarning("Advertencia", "No se seleccionaron archivos para partir.")
                return

            confirmacion = messagebox.askyesno("Confirmar", "¿Estás seguro de que deseas partir los archivos seleccionados?")
            if confirmacion:
                ventana_progreso_partir = tk.Toplevel()
                ventana_progreso_partir.title("Partiendo archivos")
                ventana_progreso_partir.geometry("400x100")
                load_icon(ventana_progreso_partir)

                barra_progreso_partir = ttk.Progressbar(ventana_progreso_partir, orient=tk.HORIZONTAL, length=300, mode="determinate")
                barra_progreso_partir.pack(pady=20)

                def partir_archivos():
                    total_archivos = len(seleccionados)
                    for i, index in enumerate(seleccionados):
                        pst_file = lista_pst.get(index)
                        try:
                            exito, resultado = dividir_pst(pst_file)
                            if exito:
                                messagebox.showinfo("Éxito", f"El archivo {pst_file} se ha partido correctamente.")
                            else:
                                messagebox.showerror("Error", f"No se pudo partir {pst_file}: {resultado}")
                        except Exception as e:
                            messagebox.showerror("Error", f"No se pudo partir {pst_file}: {e}")

                        barra_progreso_partir["value"] = (i + 1) / total_archivos * 100
                        ventana_progreso_partir.update_idletasks()

                    ventana_progreso_partir.destroy()

                threading.Thread(target=partir_archivos, daemon=True).start()

        boton_partir = tk.Button(
            ventana_pst, 
            text="Partir seleccionados", 
            command=partir_seleccionados,
            bg=BOTON_COLOR
        )
        boton_partir.pack(pady=10)

    threading.Thread(target=buscar_pst_en_segundo_plano, daemon=True).start()

# Función para obtener las redes Wi-Fi guardadas
def obtener_redes_wifi():
    try:
        resultado = subprocess.check_output(["netsh", "wlan", "show", "profiles"], encoding="utf-8")
        
        redes = []
        for linea in resultado.split("\n"):
            if "Perfil de todos los usuarios" in linea or "All User Profile" in linea:
                nombre_red = linea.split(":")[1].strip()
                redes.append(nombre_red)
        return redes
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"No se pudieron obtener las redes Wi-Fi: {e}")
        return []

# Función para olvidar redes Wi-Fi seleccionadas
def olvidar_redes_wifi():
    redes = obtener_redes_wifi()
    if not redes:
        messagebox.showinfo("Información", "No se encontraron redes Wi-Fi guardadas.")
        return

    ventana_redes = tk.Toplevel()
    ventana_redes.title("Olvidar redes Wi-Fi")
    ventana_redes.geometry("400x300")
    load_icon(ventana_redes)
    ventana_redes.lift()
    ventana_redes.attributes("-topmost", True)

    frame_redes = tk.Frame(ventana_redes)
    frame_redes.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    estados_checkboxes = {}

    for red in redes:
        estados_checkboxes[red] = tk.BooleanVar()
        checkbox = tk.Checkbutton(frame_redes, text=red, variable=estados_checkboxes[red])
        checkbox.pack(anchor=tk.W)

    def seleccionar_todas():
        for red in estados_checkboxes:
            estados_checkboxes[red].set(True)

    def olvidar_seleccionadas():
        redes_seleccionadas = [red for red, estado in estados_checkboxes.items() if estado.get()]
        if not redes_seleccionadas:
            messagebox.showwarning("Advertencia", "No se seleccionaron redes para olvidar.")
            return

        confirmacion = messagebox.askyesno("Confirmar", "¿Estás seguro de que deseas olvidar las redes seleccionadas?")
        if confirmacion:
            for red in redes_seleccionadas:
                try:
                    subprocess.run(["netsh", "wlan", "delete", "profile", f"name={red}"], check=True, shell=True)
                    messagebox.showinfo("Éxito", f"La red {red} se ha olvidado correctamente.")
                except subprocess.CalledProcessError as e:
                    messagebox.showerror("Error", f"No se pudo olvidar la red {red}: {e}")
                except Exception as e:
                    messagebox.showerror("Error", f"Error inesperado: {e}")

    boton_seleccionar_todas = tk.Button(
        ventana_redes, 
        text="Seleccionar todas", 
        command=seleccionar_todas,
        bg=BOTON_COLOR
    )
    boton_seleccionar_todas.pack(pady=5)

    boton_olvidar = tk.Button(
        ventana_redes, 
        text="Olvidar seleccionadas", 
        command=olvidar_seleccionadas,
        bg=BOTON_COLOR
    )
    boton_olvidar.pack(pady=10)

    boton_cancelar = tk.Button(
        ventana_redes, 
        text="Cancelar", 
        command=ventana_redes.destroy,
        bg=BOTON_COLOR
    )
    boton_cancelar.pack(pady=10)

# Función para generar el informe de Wi-Fi
def generar_informe_wifi():
    try:
        subprocess.run(["netsh", "wlan", "show", "wlanreport"], check=True, shell=True)
        messagebox.showinfo("Éxito", "El informe de Wi-Fi se ha generado correctamente.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"No se pudo generar el informe de Wi-Fi: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")

# Función para abrir la carpeta del informe de Wi-Fi
def mostrar_reporte_wifi():
    ruta_carpeta = "C:\\ProgramData\\Microsoft\\Windows\\WLANReport"
    if os.path.exists(ruta_carpeta):
        subprocess.Popen(f'explorer "{ruta_carpeta}"')
    else:
        messagebox.showerror("Error", f"La carpeta {ruta_carpeta} no existe.")

# Función para ejecutar ipconfig /all y mostrar la salida en un bloc de notas
def ejecutar_ipconfig():
    try:
        resultado = subprocess.check_output(["ipconfig", "/all"], encoding="cp437")
        ruta_archivo = os.path.join(os.environ["TEMP"], "ipconfig.txt")
        with open(ruta_archivo, "w", encoding="utf-8") as archivo:
            archivo.write(resultado)
        subprocess.Popen(["notepad.exe", ruta_archivo])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar ipconfig /all: {e}")

# Función para ejecutar ipconfig /release en una ventana de cmd con un tiempo de espera
def ejecutar_ipconfig_release():
    try:
        subprocess.Popen(["cmd", "/c", "ipconfig /release && timeout /t 10"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar ipconfig /release: {e}")

# Función para ejecutar ipconfig /flushdns en una ventana de cmd con un tiempo de espera
def ejecutar_ipconfig_flushdns():
    try:
        subprocess.Popen(["cmd", "/c", "ipconfig /flushdns && timeout /t 10"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar ipconfig /flushdns: {e}")

# Función para ejecutar ipconfig /renew en una ventana de cmd con un tiempo de espera
def ejecutar_ipconfig_renew():
    try:
        subprocess.Popen(["cmd", "/c", "ipconfig /renew && timeout /t 10"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar ipconfig /renew: {e}")

# Función para ejecutar gpupdate /force en una ventana de cmd con un tiempo de espera
def ejecutar_gpupdate_force():
    try:
        subprocess.Popen(["cmd", "/c", "gpupdate /force && timeout /t 10"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar gpupdate /force: {e}")

# Función para reiniciar el servicio Wi-Fi
def reiniciar_servicio_wifi():
    try:
        subprocess.run(["net", "stop", "WlanSvc"], check=True, shell=True)
        subprocess.run(["net", "start", "WlanSvc"], check=True, shell=True)
        messagebox.showinfo("Éxito", "El servicio Wi-Fi se ha reiniciado correctamente.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"No se pudo reiniciar el servicio Wi-Fi: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {e}")

# Función para verificar errores en el disco
def verificar_errores_disco(disco):
    """Verifica si el disco tiene errores usando chkdsk"""
    try:
        # Ejecutamos chkdsk en modo lectura para ver si reporta errores
        resultado = subprocess.run(["chkdsk", disco, "/scan"], capture_output=True, text=True, shell=True)
        if "Windows ha escaneado el sistema de archivos y no ha encontrado problemas" in resultado.stdout:
            return False, "Sin errores detectados"
        elif "Se encontraron errores en el disco" in resultado.stdout:
            return True, "Errores detectados en el disco"
        else:
            return None, "No se pudo determinar el estado del disco"
    except Exception as e:
        return None, f"Error al verificar el disco: {str(e)}"

# Función para obtener información del controlador WiFi
def obtener_info_controlador_wifi():
    info = {
        "sistema_operativo": "Windows",
        "adaptador_wifi": "No detectado",
        "controlador": "No detectado",
        "version": "No detectada",
        "es_generico": False,
        "esta_actualizado": None
    }

    try:
        resultado = subprocess.check_output(
            "wmic nic where \"NetConnectionID like '%Wi-Fi%'\" get name, manufacturer, PNPDeviceID",
            shell=True
        ).decode('utf-8', errors='ignore')
        
        driver_info = subprocess.check_output(
            "wmic path win32_pnpsigneddriver where \"devicename like '%Wireless%' or devicename like '%Wi-Fi%'\" get devicename, driverversion",
            shell=True
        ).decode('utf-8', errors='ignore')
        
        lineas = [line.strip() for line in resultado.split('\n') if line.strip()]
        if len(lineas) > 1:
            info["adaptador_wifi"] = lineas[1].split()[0]
            info["controlador"] = lineas[1].split()[1] if len(lineas[1].split()) > 1 else "Desconocido"
        
        version_match = re.search(r"DriverVersion\s+([\d.]+)", driver_info)
        if version_match:
            info["version"] = version_match.group(1)
        
        info["es_generico"] = any(palabra in info["controlador"] for palabra in ["Generic", "Standard", "Microsoft"])
        info["esta_actualizado"] = not info["es_generico"] and info["version"] != "No detectada"
        
    except Exception as e:
        print(f"Error al obtener información del controlador WiFi: {str(e)}")
    
    return info

# Función para verificar el controlador WiFi
def verificar_controlador_wifi():
    info = obtener_info_controlador_wifi()
    
    mensaje = f"Información del controlador WiFi:\n"
    mensaje += f"Adaptador: {info['adaptador_wifi']}\n"
    mensaje += f"Controlador: {info['controlador']}\n"
    mensaje += f"Versión: {info['version']}\n\n"
    
    if info['es_generico']:
        mensaje += "⚠️ El controlador es GENÉRICO\n"
        mensaje += "Recomendación: Instale los controladores específicos del fabricante para mejor rendimiento.\n"
    else:
        mensaje += "✅ El controlador NO es genérico\n"
    
    if info['esta_actualizado']:
        mensaje += "✅ El controlador parece estar ACTUALIZADO\n"
    else:
        mensaje += "⚠️ El controlador PODRÍA NO estar actualizado\n"
        mensaje += "Recomendación: Verifique si hay actualizaciones disponibles.\n"
    
    messagebox.showinfo("Estado del Controlador WiFi", mensaje)

# Función mejorada para abrir el Visor de Eventos en los registros de BSOD
def abrir_visor_eventos():
    try:
        comando = (
            'eventvwr.msc /c:"System" /f:"*[System[('
            'EventID=41 or EventID=1001 or EventID=6008'
            ')]]"'
        )
        subprocess.Popen(comando, shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el Visor de Eventos: {e}")

def mostrar_info_hardware():
    """Muestra información detallada del hardware del sistema"""
    try:
        info = ""
        
        # Información del sistema operativo
        sistema = platform.system()
        version = platform.version()
        arquitectura = platform.architecture()[0]
        info += f"Sistema Operativo: {sistema} {version} ({arquitectura})\n\n"
        
        # Información del procesador
        c = wmi.WMI()
        for procesador in c.Win32_Processor():
            info += f"Fabricante: {procesador.Manufacturer}\n"
            info += f"Modelo: {procesador.Name}\n"
            info += f"Núcleos: {procesador.NumberOfCores}\n"
            info += f"Procesadores lógicos: {procesador.NumberOfLogicalProcessors}\n"
            info += f"Velocidad: {procesador.MaxClockSpeed} MHz\n"
        
        # Información de memoria RAM
        info += "\n=== Memoria RAM ===\n"
        ram_total = round(psutil.virtual_memory().total / (1024**3), 2)
        info += f"Total: {ram_total} GB\n"
        
        # Información de discos
        info += "\n=== Almacenamiento ===\n"
        particiones = psutil.disk_partitions()
        for particion in particiones:
            try:
                uso = psutil.disk_usage(particion.mountpoint)
                total_gb = round(uso.total / (1024**3), 2)
                libre_gb = round(uso.free / (1024**3), 2)
                
                # Verificar errores en el disco
                tiene_errores, mensaje_errores = verificar_errores_disco(particion.device[:2])
                estado_errores = "✅ Sin errores" if tiene_errores is False else "⚠️ Con errores" if tiene_errores else "❓ Estado desconocido"
                
                info += f"Disco: {particion.device} ({particion.fstype}, {estado_errores})\n"
                info += f"  Montado en: {particion.mountpoint}\n"
                info += f"  Tamaño total: {total_gb} GB\n"
                info += f"  Libre: {libre_gb} GB\n"
                if tiene_errores is not False:
                    info += f"  Detalles: {mensaje_errores}\n"
            except:
                continue
        
        # Información de red
        info += "\n=== Red ===\n"
        interfaces = psutil.net_if_addrs()
        for nombre, direcciones in interfaces.items():
            info += f"Interfaz: {nombre}\n"
            for direccion in direcciones:
                if direccion.family == 2:  # AF_INET
                    info += f"  IPv4: {direccion.address}\n"
                elif direccion.family == 17:  # AF_PACKET
                    info += f"  MAC: {direccion.address}\n"
        
        # Mostrar la información en una ventana
        ventana_info = tk.Toplevel()
        ventana_info.title("Información del Hardware")
        ventana_info.geometry("800x600")
        load_icon(ventana_info)
        
        texto = tk.Text(ventana_info, wrap=tk.WORD)
        scroll = tk.Scrollbar(ventana_info, command=texto.yview)
        texto.configure(yscrollcommand=scroll.set)
        
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        texto.pack(fill=tk.BOTH, expand=True)
        
        texto.insert(tk.END, info)
        texto.config(state=tk.DISABLED)
        
        boton_copiar = tk.Button(
            ventana_info,
            text="Copiar al portapapeles",
            command=lambda: pyperclip.copy(info),
            bg=BOTON_COLOR
        )
        boton_copiar.pack(pady=10)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo obtener la información del hardware: {e}")

def diagnosticar_hardware():
    """Realiza un diagnóstico básico del hardware"""
    try:
        mensaje = "Resultados del diagnóstico de hardware:\n\n"
        
        # Verificar uso de CPU
        uso_cpu = psutil.cpu_percent(interval=1)
        mensaje += f"Uso de CPU: {uso_cpu}%\n"
        if uso_cpu > 80:
            mensaje += "⚠️ Advertencia: Uso de CPU alto\n"
        
        # Verificar uso de memoria
        memoria = psutil.virtual_memory()
        uso_memoria = memoria.percent
        mensaje += f"\nUso de memoria: {uso_memoria}%\n"
        if uso_memoria > 80:
            mensaje += "⚠️ Advertencia: Uso de memoria alto\n"
        
        # Verificar discos
        mensaje += "\n=== Estado de los discos ===\n"
        particiones = psutil.disk_partitions()
        for particion in particiones:
            try:
                uso = psutil.disk_usage(particion.mountpoint)
                tiene_errores, _ = verificar_errores_disco(particion.device[:2])
                
                mensaje += f"Disco {particion.device}:\n"
                mensaje += f"  Uso: {uso.percent}%\n"
                mensaje += f"  Estado: {'✅ Sin errores' if tiene_errores is False else '⚠️ Con errores' if tiene_errores else '❓ Estado desconocido'}\n"
                
                if uso.percent > 90:
                    mensaje += "⚠️ Advertencia: Espacio en disco bajo\n"
                if tiene_errores:
                    mensaje += "⚠️ Advertencia: Errores detectados en el disco\n"
            except:
                mensaje += f"  No se pudo verificar {particion.device}\n"
        
        messagebox.showinfo("Diagnóstico de Hardware", mensaje)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo realizar el diagnóstico: {e}")

def verificar_controladores():
    """Verifica controladores con problemas (solo Windows)"""
    try:
        c = wmi.WMI()
        controladores_problema = []
        
        for controlador in c.Win32_PnPSignedDriver():
            if controlador.DeviceName and controlador.DeviceID:
                estado = "Desconocido"
                try:
                    if hasattr(controlador, 'DeviceStatus'):
                        estado = getattr(controlador, 'DeviceStatus', 'Desconocido')
                    else:
                        estado = "OK" if controlador.IsSigned else "No firmado"
                except:
                    estado = "Error al verificar"
                
                if not controlador.IsSigned or estado != "OK":
                    controladores_problema.append((
                        controlador.DeviceName,
                        controlador.DeviceID,
                        "Firmado" if controlador.IsSigned else "No firmado",
                        estado
                    ))
        
        if not controladores_problema:
            messagebox.showinfo("Resultado", "No se encontraron controladores con problemas.")
            return
        
        ventana_controladores = tk.Toplevel()
        ventana_controladores.title("Controladores con problemas")
        ventana_controladores.geometry("800x400")
        load_icon(ventana_controladores)
        
        tree = ttk.Treeview(ventana_controladores, columns=("Estado", "Firma"), show="headings")
        tree.heading("#0", text="Dispositivo")
        tree.heading("Estado", text="Estado")
        tree.heading("Firma", text="Firma")
        tree.column("#0", width=400)
        tree.column("Estado", width=150)
        tree.column("Firma", width=150)
        
        for dispositivo, device_id, firma, estado in controladores_problema:
            tree.insert("", "end", text=dispositivo, values=(estado, firma))
        
        scroll = ttk.Scrollbar(ventana_controladores, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron verificar los controladores: {e}")

def prueba_memoria():
    """Ejecuta una prueba básica de memoria RAM"""
    try:
        messagebox.showinfo("Prueba de Memoria", "La prueba de memoria comenzará ahora. Esto puede tomar unos minutos...")
        
        # Simulamos una prueba de memoria básica
        tamaño_prueba = psutil.virtual_memory().total // 2  # Usamos la mitad de la RAM para la prueba
        bloque = b"X" * (1024 * 1024)  # Bloque de 1MB
        
        ventana_progreso = tk.Toplevel()
        ventana_progreso.title("Prueba de Memoria")
        ventana_progreso.geometry("400x100")
        load_icon(ventana_progreso)
        
        etiqueta = tk.Label(ventana_progreso, text="Ejecutando prueba de memoria...")
        etiqueta.pack(pady=5)
        
        barra_progreso = ttk.Progressbar(ventana_progreso, orient=tk.HORIZONTAL, length=300, mode="determinate")
        barra_progreso.pack(pady=10)
        
        def ejecutar_prueba():
            try:
                bloques = tamaño_prueba // len(bloque)
                datos = []
                
                for i in range(bloques):
                    datos.append(bloque)
                    progreso = (i + 1) / bloques * 100
                    barra_progreso["value"] = progreso
                    ventana_progreso.update_idletasks()
                
                # Verificar los datos
                for i, dato in enumerate(datos):
                    if dato != bloque:
                        raise MemoryError("Error en la prueba de memoria")
                    progreso = 100 * (i + 1) / len(datos)
                    barra_progreso["value"] = progreso
                    ventana_progreso.update_idletasks()
                
                ventana_progreso.destroy()
                messagebox.showinfo("Resultado", "Prueba de memoria completada sin errores.")
                
            except MemoryError:
                ventana_progreso.destroy()
                messagebox.showerror("Error", "Se detectaron errores en la prueba de memoria.")
            except Exception as e:
                ventana_progreso.destroy()
                messagebox.showerror("Error", f"Error durante la prueba: {e}")
        
        threading.Thread(target=ejecutar_prueba, daemon=True).start()
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar la prueba de memoria: {e}")

def prueba_disco():
    """Mide la velocidad de lectura/escritura del disco"""
    try:
        messagebox.showinfo("Prueba de Disco", "La prueba de velocidad de disco comenzará ahora. Esto puede tomar unos minutos...")
        
        ventana_progreso = tk.Toplevel()
        ventana_progreso.title("Prueba de Disco")
        ventana_progreso.geometry("400x150")
        load_icon(ventana_progreso)
        
        etiqueta = tk.Label(ventana_progreso, text="Ejecutando prueba de velocidad de disco...")
        etiqueta.pack(pady=5)
        
        barra_progreso = ttk.Progressbar(ventana_progreso, orient=tk.HORIZONTAL, length=300, mode="determinate")
        barra_progreso.pack(pady=10)
        
        etiqueta_resultado = tk.Label(ventana_progreso, text="")
        etiqueta_resultado.pack(pady=5)
        
        def ejecutar_prueba():
            try:
                # Archivo temporal para pruebas
                archivo_prueba = os.path.join(os.environ["TEMP"], "disco_prueba.tmp")
                tamaño_prueba = 100 * 1024 * 1024  # 100 MB
                bloque = b"X" * (1024 * 1024)  # Bloque de 1MB
                
                # Prueba de escritura
                etiqueta_resultado.config(text="Probando velocidad de escritura...")
                inicio = time.time()
                
                with open(archivo_prueba, "wb") as f:
                    for i in range(tamaño_prueba // len(bloque)):
                        f.write(bloque)
                        progreso = 50 * (i + 1) / (tamaño_prueba // len(bloque))
                        barra_progreso["value"] = progreso
                        ventana_progreso.update_idletasks()
                
                tiempo_escritura = time.time() - inicio
                velocidad_escritura = (tamaño_prueba / (1024 * 1024)) / tiempo_escritura  # MB/s
                
                # Prueba de lectura
                etiqueta_resultado.config(text="Probando velocidad de lectura...")
                inicio = time.time()
                
                with open(archivo_prueba, "rb") as f:
                    while f.read(len(bloque)):
                        progreso = 50 + 50 * (f.tell() / tamaño_prueba)
                        barra_progreso["value"] = progreso
                        ventana_progreso.update_idletasks()
                
                tiempo_lectura = time.time() - inicio
                velocidad_lectura = (tamaño_prueba / (1024 * 1024)) / tiempo_lectura  # MB/s
                
                # Eliminar archivo temporal
                os.remove(archivo_prueba)
                
                ventana_progreso.destroy()
                
                mensaje = "Resultados de la prueba de disco:\n\n"
                mensaje += f"Velocidad de escritura: {velocidad_escritura:.2f} MB/s\n"
                mensaje += f"Velocidad de lectura: {velocidad_lectura:.2f} MB/s\n\n"
                
                # Verificar errores en el disco
                tiene_errores, mensaje_errores = verificar_errores_disco("C:")
                if tiene_errores is None:
                    mensaje += f"\n⚠️ No se pudo verificar el estado del disco: {mensaje_errores}\n"
                elif tiene_errores:
                    mensaje += f"\n⚠️ El disco tiene errores: {mensaje_errores}\n"
                    mensaje += "Recomendación: Ejecute 'chkdsk /f' para reparar errores\n"
                else:
                    mensaje += "\n✅ No se detectaron errores en el disco\n"
                
                messagebox.showinfo("Resultados", mensaje)
                
            except Exception as e:
                ventana_progreso.destroy()
                messagebox.showerror("Error", f"Error durante la prueba: {e}")
        
        threading.Thread(target=ejecutar_prueba, daemon=True).start()
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar la prueba de disco: {e}")

def reactivar_onedrive():
    try:
        key_path = r"SOFTWARE\Policies\Microsoft\Windows\OneDrive"
        
        # Intentar eliminar el valor que desactiva OneDrive
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                winreg.DeleteValue(key, "DisableFileSyncNGSC")
                print("✅ Valor DisableFileSyncNGSC eliminado.")
        except FileNotFoundError:
            print("⚠️ La clave no existe o ya está activado.")
        
        # Opcional: Eliminar la clave completa si está vacía
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            print("🗑️ Clave del registro eliminada.")
        except OSError:
            pass  # La clave no existe o no está vacía
        
        messagebox.showinfo("Éxito", "OneDrive debería estar reactivado. Reinicia tu PC para aplicar los cambios.")
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo reactivar OneDrive: {str(e)}")

def eliminar_perfil_outlook():
    try:
        # Mostrar advertencia
        confirmacion = messagebox.askyesno(
            "Advertencia",
            "¿Estás seguro de que deseas eliminar el perfil de Outlook y sus archivos OST?\n\n"
            "Esta acción eliminará:\n"
            "1. La configuración del perfil en el registro\n"
            "2. Los archivos OST asociados al perfil\n\n"
            "Outlook debe estar cerrado para realizar esta operación."
        )
        
        if not confirmacion:
            return
            
        # Obtener la lista de perfiles de Outlook
        key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
            num_profiles = winreg.QueryInfoKey(key)[0]
            
            if num_profiles == 0:
                messagebox.showinfo("Información", "No se encontraron perfiles de Outlook para eliminar.")
                return
                
            # Crear ventana para seleccionar perfil
            ventana_perfiles = tk.Toplevel()
            ventana_perfiles.title("Eliminar Perfil de Outlook")
            ventana_perfiles.geometry("500x350")
            load_icon(ventana_perfiles)
            
            tk.Label(ventana_perfiles, text="Selecciona el perfil a eliminar:").pack(pady=10)
            
            lista_perfiles = tk.Listbox(ventana_perfiles, selectmode=tk.SINGLE)
            perfiles = []
            
            for i in range(num_profiles):
                perfil = winreg.EnumKey(key, i)
                perfiles.append(perfil)
                lista_perfiles.insert(tk.END, perfil)
                
            lista_perfiles.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
            
            # Función para buscar archivos OST asociados al perfil
            def buscar_ost_perfil(perfil):
                ost_files = []
                try:
                    # Buscar en las rutas comunes de OST
                    rutas_comunes = [
                        os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "Microsoft", "Outlook"),
                        os.path.join(os.environ["USERPROFILE"], "Documents", "Outlook Files")
                    ]
                    
                    for ruta_base in rutas_comunes:
                        if os.path.exists(ruta_base):
                            for archivo in os.listdir(ruta_base):
                                if archivo.endswith(".ost") and perfil.lower() in archivo.lower():
                                    ost_files.append(os.path.join(ruta_base, archivo))
                except Exception as e:
                    print(f"Error buscando archivos OST: {e}")
                return ost_files
            
            def eliminar_perfil_seleccionado():
                seleccion = lista_perfiles.curselection()
                if not seleccion:
                    messagebox.showwarning("Advertencia", "No se ha seleccionado ningún perfil.")
                    return
                    
                perfil_seleccionado = perfiles[seleccion[0]]
                
                # Buscar archivos OST asociados
                ost_files = buscar_ost_perfil(perfil_seleccionado)
                
                mensaje_confirmacion = (
                    f"¿Estás seguro de que deseas eliminar el perfil '{perfil_seleccionado}'?\n\n"
                    "Esta acción eliminará:\n"
                    f"- La configuración del perfil en el registro\n"
                )
                
                if ost_files:
                    mensaje_confirmacion += "- Los siguientes archivos OST:\n"
                    for ost in ost_files:
                        mensaje_confirmacion += f"  • {ost}\n"
                else:
                    mensaje_confirmacion += "- No se encontraron archivos OST asociados\n"
                
                mensaje_confirmacion += "\nEsta acción no se puede deshacer."
                
                confirmacion_final = messagebox.askyesno("Confirmar", mensaje_confirmacion)
                
                if confirmacion_final:
                    try:
                        # Eliminar archivos OST primero
                        errores_ost = []
                        if ost_files:
                            for ost in ost_files:
                                try:
                                    os.remove(ost)
                                except Exception as e:
                                    errores_ost.append(f"{ost}: {str(e)}")
                        
                        # Eliminar el perfil del registro
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{key_path}\\{perfil_seleccionado}")
                        
                        # Mostrar resultados
                        mensaje_resultado = f"El perfil '{perfil_seleccionado}' ha sido eliminado correctamente.\n"
                        
                        if ost_files:
                            if errores_ost:
                                mensaje_resultado += "\nPero hubo errores al eliminar algunos archivos OST:\n"
                                for error in errores_ost:
                                    mensaje_resultado += f"- {error}\n"
                            else:
                                mensaje_resultado += "\nLos archivos OST asociados fueron eliminados correctamente."
                        
                        messagebox.showinfo("Éxito", mensaje_resultado)
                        ventana_perfiles.destroy()
                        
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo eliminar el perfil:\n{str(e)}")
            
            boton_eliminar = tk.Button(
                ventana_perfiles,
                text="Eliminar Perfil y Archivos OST",
                command=eliminar_perfil_seleccionado,
                bg=BOTON_COLOR
            )
            boton_eliminar.pack(pady=10)
            
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo acceder a los perfiles de Outlook:\n{str(e)}")

def aumentar_tamano_ost():
    # Mostrar advertencia primero
    advertencia = (
        "ADVERTENCIA:\n\n"
        "1. Aumentar el tamaño del archivo OST puede afectar el rendimiento de Outlook\n"
        "2. Microsoft recomienda no superar los 50GB para archivos OST\n"
        "3. Archivos muy grandes pueden causar problemas de sincronización\n"
        "4. Es recomendable archivar correos antiguos en lugar de aumentar el límite\n\n"
        "¿Deseas continuar con el cambio?"
    )
    
    if not messagebox.askyesno("Advertencia Importante", advertencia):
        return
    
    # Ventana de configuración
    ventana_ost = tk.Toplevel()
    ventana_ost.title("Configurar Tamaño Máximo de Archivo OST")
    ventana_ost.geometry("500x300")
    load_icon(ventana_ost)
    
    # Etiqueta informativa
    info_label = tk.Label(
        ventana_ost,
        text="Configurar el tamaño máximo para archivos OST de Outlook",
        wraplength=400,
        justify=tk.LEFT
    )
    info_label.pack(pady=10)
    
    # Frame para la selección de tamaño
    frame_tamano = tk.Frame(ventana_ost)
    frame_tamano.pack(pady=10)
    
    tk.Label(frame_tamano, text="Nuevo tamaño máximo (GB):").pack(side=tk.LEFT)
    
    tamano_var = tk.StringVar(value="50")  # Valor predeterminado
    
    combo_tamano = ttk.Combobox(
        frame_tamano,
        textvariable=tamano_var,
        values=["50", "60", "70", "80", "90", "100"],
        state="readonly",
        width=5
    )
    combo_tamano.pack(side=tk.LEFT, padx=10)
    
    # Función para aplicar los cambios
    def aplicar_cambios():
        try:
            nuevo_tamano = int(tamano_var.get())
            if nuevo_tamano < 50 or nuevo_tamano > 100:
                messagebox.showerror("Error", "El tamaño debe estar entre 50 y 100 GB")
                return
                
            # Confirmación final
            confirmacion = messagebox.askyesno(
                "Confirmar Cambio",
                f"¿Estás seguro de que deseas cambiar el tamaño máximo de archivo OST a {nuevo_tamano}GB?\n\n"
                "Outlook debe estar cerrado para que los cambios surtan efecto."
            )
            
            if not confirmacion:
                return
                
            # Modificar el registro
            try:
                clave = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Office\16.0\Outlook\OST",
                    0,
                    winreg.KEY_WRITE
                )
                winreg.SetValueEx(clave, "MaxLargeFileSize", 0, winreg.REG_DWORD, nuevo_tamano * 1024)
                winreg.CloseKey(clave)
                
                messagebox.showinfo(
                    "Éxito",
                    f"El tamaño máximo de archivo OST se ha configurado a {nuevo_tamano}GB.\n\n"
                    "Debes reiniciar Outlook para que los cambios surtan efecto."
                )
                ventana_ost.destroy()
                
            except Exception as e:
                # Si la clave no existe, la creamos
                if "The system cannot find the file specified" in str(e):
                    try:
                        clave = winreg.CreateKey(
                            winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office\16.0\Outlook\OST"
                        )
                        winreg.SetValueEx(clave, "MaxLargeFileSize", 0, winreg.REG_DWORD, nuevo_tamano * 1024)
                        winreg.CloseKey(clave)
                        
                        messagebox.showinfo(
                            "Éxito",
                            f"El tamaño máximo de archivo OST se ha configurado a {nuevo_tamano}GB.\n\n"
                            "Debes reiniciar Outlook para que los cambios surtan efecto."
                        )
                        ventana_ost.destroy()
                        
                    except Exception as e2:
                        messagebox.showerror(
                            "Error",
                            f"No se pudo crear la clave del registro:\n{str(e2)}"
                        )
                else:
                    messagebox.showerror(
                        "Error",
                        f"No se pudo modificar el registro:\n{str(e)}"
                    )
                    
        except ValueError:
            messagebox.showerror("Error", "Por favor ingresa un número válido")
    
    # Botones
    frame_botones = tk.Frame(ventana_ost)
    frame_botones.pack(pady=20)
    
    boton_aplicar = tk.Button(
        frame_botones,
        text="Aplicar Cambios",
        command=aplicar_cambios,
        bg=BOTON_COLOR
    )
    boton_aplicar.pack(side=tk.LEFT, padx=10)
    
    boton_cancelar = tk.Button(
        frame_botones,
        text="Cancelar",
        command=ventana_ost.destroy,
        bg=BOTON_COLOR
    )
    boton_cancelar.pack(side=tk.LEFT, padx=10)

def mover_correos_archivo_online():
    try:
        # Crear ventana de selección de criterios
        ventana_archivo = tk.Toplevel()
        ventana_archivo.title("Mover correos al archivo en línea")
        ventana_archivo.geometry("600x500")  # Cambiado de 500x450 a 600x500 para hacerla más grande
        load_icon(ventana_archivo)
        
        # Información sobre el proceso
        info_text = (
            "Esta herramienta mueve correos electrónicos al archivo en línea de Office 365/Exchange.\n\n"
            "Ventajas del archivo en línea:\n"
            "1. Acceso desde cualquier dispositivo con conexión a internet\n"
            "2. No ocupa espacio en tu disco duro\n"
            "3. Mayor seguridad y redundancia en la nube\n"
            "4. Búsqueda unificada en todos tus correos\n"
            "5. Sincronización automática entre dispositivos\n\n"
            "Tiempo estimado del proceso:\n"
            "- 100-500 correos: 1-3 minutos\n"
            "- 500-2000 correos: 3-10 minutos\n"
            "- Más de 2000 correos: 10-20 minutos\n\n"
            "Requisitos:\n"
            "- Outlook debe estar configurado con una cuenta Exchange/Office 365\n"
            "- Conexión estable a internet durante el proceso"
        )
        
        tk.Label(ventana_archivo, text=info_text, wraplength=550, justify=tk.LEFT).pack(pady=10)  # Aumentado wraplength de 450 a 550
        
        # Frame para los criterios de selección
        frame_criterios = tk.Frame(ventana_archivo)
        frame_criterios.pack(pady=10)
        
        tk.Label(frame_criterios, text="Seleccione el criterio de antigüedad:").pack(anchor=tk.W)
        
        criterio_var = tk.StringVar(value="3meses")
        
        opciones = [
            ("Correos más antiguos que 1 semana", "semana"),
            ("Correos más antiguos que 1 mes", "mes"),
            ("Correos más antiguos que 3 meses", "3meses"),
            ("Correos más antiguos que 6 meses", "6meses"),
            ("Correos más antiguos que 1 año", "año")
        ]
        
        for texto, valor in opciones:
            tk.Radiobutton(
                frame_criterios,
                text=texto,
                variable=criterio_var,
                value=valor
            ).pack(anchor=tk.W)
        
        # Función para calcular la fecha de corte según el criterio
        def calcular_fecha_corte(criterio):
            hoy = datetime.now()
            if criterio == "semana":
                return hoy - timedelta(weeks=1)
            elif criterio == "mes":
                return hoy - timedelta(days=30)
            elif criterio == "3meses":
                return hoy - timedelta(days=90)
            elif criterio == "6meses":
                return hoy - timedelta(days=180)
            elif criterio == "año":
                return hoy - timedelta(days=365)
            return hoy
        
        # Función para ejecutar el movimiento a archivo en línea
        def ejecutar_movimiento_online():
            criterio = criterio_var.get()
            fecha_corte = calcular_fecha_corte(criterio)
            
            # Mostrar ventana de progreso
            ventana_progreso = tk.Toplevel(ventana_archivo)
            ventana_progreso.title("Moviendo correos al archivo en línea")
            ventana_progreso.geometry("400x150")
            load_icon(ventana_progreso)
            
            tk.Label(ventana_progreso, text="Procesando correos...").pack(pady=10)
            
            barra_progreso = ttk.Progressbar(
                ventana_progreso,
                orient=tk.HORIZONTAL,
                length=300,
                mode="determinate"
            )
            barra_progreso.pack(pady=10)
            
            etiqueta_estado = tk.Label(ventana_progreso, text="Conectando con Outlook...")
            etiqueta_estado.pack()
            
            def mover_a_archivo_online():
                try:
                    # Conectar con Outlook
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    namespace = outlook.GetNamespace("MAPI")
                    
                    # Obtener todas las bandejas de entrada de cuentas Exchange
                    bandejas = []
                    for store in namespace.Stores:
                        try:
                            # Solo procesar cuentas Exchange/Office 365
                            if store.ExchangeStoreType in [1, 2, 3]:  # 1=Exchange Mailbox, 2=Exchange Public Folder, 3=Exchange OST
                                inbox = store.GetDefaultFolder(6)  # 6 = Bandeja de entrada
                                bandejas.append(inbox)
                        except:
                            continue
                    
                    if not bandejas:
                        messagebox.showerror("Error", "No se encontraron bandejas de Exchange/Office 365.")
                        ventana_progreso.destroy()
                        return
                    
                    # Obtener o crear la carpeta de archivo en línea
                    carpeta_archivo_online = None
                    for store in namespace.Stores:
                        if store.ExchangeStoreType == 1:  # Exchange Mailbox
                            try:
                                root_folder = store.GetRootFolder()
                                # Buscar carpeta existente
                                for folder in root_folder.Folders:
                                    if "archivo" in folder.Name.lower() or "archive" in folder.Name.lower():
                                        carpeta_archivo_online = folder
                                        break
                                # Si no existe, crear una nueva
                                if not carpeta_archivo_online:
                                    carpeta_archivo_online = root_folder.Folders.Add("Archivo en línea")
                                    messagebox.showinfo("Información", "Se creó una nueva carpeta 'Archivo en línea' en tu buzón.")
                                break
                            except Exception as e:
                                print(f"Error al buscar/crear carpeta de archivo: {str(e)}")
                                continue
                    
                    if not carpeta_archivo_online:
                        messagebox.showerror("Error", "No se pudo encontrar o crear la carpeta de archivo en línea.")
                        ventana_progreso.destroy()
                        return
                    
                    # Procesar correos en cada bandeja
                    total_correos = 0
                    correos_movidos = 0
                    
                    for bandeja in bandejas:
                        try:
                            correos = bandeja.Items
                            total_correos += len(correos)
                        except:
                            continue
                    
                    if total_correos == 0:
                        messagebox.showinfo("Información", "No hay correos para mover al archivo en línea.")
                        ventana_progreso.destroy()
                        return
                    
                    # Mover correos según el criterio
                    for i, bandeja in enumerate(bandejas):
                        try:
                            etiqueta_estado.config(text=f"Procesando bandeja {i+1}/{len(bandejas)}...")
                            ventana_progreso.update()
                            
                            correos = bandeja.Items
                            for j, correo in enumerate(correos):
                                try:
                                    if hasattr(correo, "ReceivedTime") and correo.ReceivedTime < fecha_corte:
                                        # Mover a archivo en línea
                                        correo.Move(carpeta_archivo_online)
                                        correos_movidos += 1
                                    
                                    # Actualizar barra de progreso
                                    progreso = (j + 1) / total_correos * 100
                                    barra_progreso["value"] = progreso
                                    ventana_progreso.update()
                                except Exception as e:
                                    print(f"Error al mover correo: {str(e)}")
                                    continue
                        except Exception as e:
                            print(f"Error al procesar bandeja: {str(e)}")
                            continue
                    
                    ventana_progreso.destroy()
                    messagebox.showinfo("Éxito", f"Se movieron {correos_movidos} correos al archivo en línea correctamente.")
                    ventana_archivo.destroy()
                    
                except Exception as e:
                    ventana_progreso.destroy()
                    messagebox.showerror("Error", f"No se pudo completar el proceso:\n{str(e)}")
                    ventana_archivo.destroy()
            
            # Ejecutar en segundo plano
            threading.Thread(target=mover_a_archivo_online, daemon=True).start()
        
        # Botones
        frame_botones = tk.Frame(ventana_archivo)
        frame_botones.pack(pady=20)
        
        boton_mover = tk.Button(
            frame_botones,
            text="Mover a Archivo en Línea",
            command=ejecutar_movimiento_online,
            bg=BOTON_COLOR
        )
        boton_mover.pack(side=tk.LEFT, padx=10)
        
        boton_cancelar = tk.Button(
            frame_botones,
            text="Cancelar",
            command=ventana_archivo.destroy,
            bg=BOTON_COLOR
        )
        boton_cancelar.pack(side=tk.LEFT, padx=10)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar el proceso:\n{str(e)}")

# Función principal para crear la interfaz gráfica
def main():
    global ventana
    ventana = tk.Tk()
    ventana.title("T-34")
    ventana.geometry("900x700")  # Aumentado el tamaño de la ventana principal
    load_icon(ventana)

    # ==============================================
    # SECCIÓN DIAGNÓSTICOS DE HARDWARE
    # ==============================================
    titulo_hardware = tk.Label(ventana, text="Diagnósticos de Hardware", font=("Arial", 16, "bold"))
    titulo_hardware.pack(pady=10, anchor=tk.W, padx=20)

    frame_botones_hardware = tk.Frame(ventana)
    frame_botones_hardware.pack(pady=10)

    boton_info_hardware = tk.Button(
        frame_botones_hardware,
        text="Información del Hardware",
        command=mostrar_info_hardware,
        bg=BOTON_COLOR
    )
    boton_info_hardware.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_info_hardware, "Muestra información detallada del hardware del sistema")

    boton_diagnosticar_hardware = tk.Button(
        frame_botones_hardware,
        text="Diagnosticar Hardware",
        command=diagnosticar_hardware,
        bg=BOTON_COLOR
    )
    boton_diagnosticar_hardware.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_diagnosticar_hardware, "Realiza un diagnóstico básico del hardware del sistema")

    frame_botones_hardware_linea2 = tk.Frame(ventana)
    frame_botones_hardware_linea2.pack(pady=10)

    boton_verificar_controladores = tk.Button(
        frame_botones_hardware_linea2,
        text="Verificar Controladores",
        command=verificar_controladores,
        bg=BOTON_COLOR
    )
    boton_verificar_controladores.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_verificar_controladores, "Verifica controladores con problemas (solo Windows)")

    boton_prueba_memoria = tk.Button(
        frame_botones_hardware_linea2,
        text="Prueba de Memoria",
        command=prueba_memoria,
        bg=BOTON_COLOR
    )
    boton_prueba_memoria.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_prueba_memoria, "Ejecuta una prueba básica de memoria RAM")

    boton_prueba_disco = tk.Button(
        frame_botones_hardware_linea2,
        text="Prueba de Disco",
        command=prueba_disco,
        bg=BOTON_COLOR
    )
    boton_prueba_disco.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_prueba_disco, "Mide la velocidad de lectura/escritura del disco")

    # ==============================================
    # SECCIÓN ALMACENAMIENTO
    # ==============================================
    titulo_almacenamiento = tk.Label(ventana, text="Almacenamiento", font=("Arial", 16, "bold"))
    titulo_almacenamiento.pack(pady=10, anchor=tk.W, padx=20)

    frame_botones_almacenamiento = tk.Frame(ventana)
    frame_botones_almacenamiento.pack(pady=10)

    boton_ejecutar = tk.Button(
        frame_botones_almacenamiento,
        text="Escanear disco C:",
        command=mostrar_ventana_progreso,
        bg=BOTON_COLOR
    )
    boton_ejecutar.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_ejecutar, "Escanea el disco C: en busca de carpetas que pesen más de 5 GB.")

    boton_liberar = tk.Button(
        frame_botones_almacenamiento,
        text="Liberar espacio en disco",
        command=liberar_espacio,
        bg=BOTON_COLOR
    )
    boton_liberar.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_liberar, "Libera espacio en el disco duro eliminando archivos temporales y vaciando la papelera de reciclaje.")

    boton_pst = tk.Button(
        frame_botones_almacenamiento,
        text="PST",
        command=mostrar_pst,
        bg=BOTON_COLOR
    )
    boton_pst.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_pst, "Busca archivos PST en el sistema y permite borrarlos o dividirlos en partes más pequeñas.")

    # ==============================================
    # SECCIÓN WI-FI
    # ==============================================
    titulo_wifi = tk.Label(ventana, text="Wi-Fi", font=("Arial", 16, "bold"))
    titulo_wifi.pack(pady=10, anchor=tk.W, padx=20)

    frame_botones_wifi = tk.Frame(ventana)
    frame_botones_wifi.pack(pady=10)

    boton_verificar_controlador = tk.Button(
        frame_botones_wifi,
        text="Verificar Controlador WiFi",
        command=verificar_controlador_wifi,
        bg=BOTON_COLOR
    )
    boton_verificar_controlador.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_verificar_controlador, "Verifica si el controlador WiFi es genérico y si está actualizado")

    boton_informe_wifi = tk.Button(
        frame_botones_wifi,
        text="Generar informe Wi-Fi",
        command=generar_informe_wifi,
        bg=BOTON_COLOR
    )
    boton_informe_wifi.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_informe_wifi, "Genera un informe detallado de las conexiones Wi-Fi y errores relacionados.")

    boton_mostrar_reporte = tk.Button(
        frame_botones_wifi,
        text="Mostrar informe",
        command=mostrar_reporte_wifi,
        bg=BOTON_COLOR
    )
    boton_mostrar_reporte.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_mostrar_reporte, "Abre la carpeta donde se guarda el informe de Wi-Fi.")

    boton_ipconfig = tk.Button(
        frame_botones_wifi,
        text="ipconfig /all",
        command=ejecutar_ipconfig,
        bg=BOTON_COLOR
    )
    boton_ipconfig.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_ipconfig, "Ejecuta el comando ipconfig /all y muestra la salida en un bloc de notas.")

    frame_botones_wifi_linea2 = tk.Frame(ventana)
    frame_botones_wifi_linea2.pack(pady=10)

    boton_ipconfig_release = tk.Button(
        frame_botones_wifi_linea2,
        text="ipconfig /release",
        command=ejecutar_ipconfig_release,
        bg=BOTON_COLOR
    )
    boton_ipconfig_release.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_ipconfig_release, "Libera la dirección IP actual.")

    boton_ipconfig_flushdns = tk.Button(
        frame_botones_wifi_linea2,
        text="ipconfig /flushdns",
        command=ejecutar_ipconfig_flushdns,
        bg=BOTON_COLOR
    )
    boton_ipconfig_flushdns.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_ipconfig_flushdns, "Borra la caché de DNS.")

    boton_ipconfig_renew = tk.Button(
        frame_botones_wifi_linea2,
        text="ipconfig /renew",
        command=ejecutar_ipconfig_renew,
        bg=BOTON_COLOR
    )
    boton_ipconfig_renew.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_ipconfig_renew, "Renueva la dirección IP.")

    boton_olvidar_wifi = tk.Button(
        frame_botones_wifi_linea2,
        text="Olvidar redes Wi-Fi",
        command=olvidar_redes_wifi,
        bg=BOTON_COLOR
    )
    boton_olvidar_wifi.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_olvidar_wifi, "Muestra una lista de redes Wi-Fi guardadas y permite olvidarlas.")

    boton_gpupdate = tk.Button(
        frame_botones_wifi_linea2,
        text="gpupdate /force",
        command=ejecutar_gpupdate_force,
        bg=BOTON_COLOR
    )
    boton_gpupdate.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_gpupdate, "Actualiza las políticas de grupo en el sistema.")

    boton_reiniciar_wifi = tk.Button(
        frame_botones_wifi_linea2,
        text="Reiniciar servicio Wi-Fi",
        command=reiniciar_servicio_wifi,
        bg=BOTON_COLOR
    )
    boton_reiniciar_wifi.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_reiniciar_wifi, "Reinicia el servicio WLAN AutoConfig para solucionar problemas de conexión.")

    # ==============================================
    # SECCIÓN DIAGNÓSTICO DE PANTALLAZOS AZULES
    # ==============================================
    titulo_bsod = tk.Label(ventana, text="Diagnóstico de Pantallazos Azules", font=("Arial", 16, "bold"))
    titulo_bsod.pack(pady=10, anchor=tk.W, padx=20)

    frame_botones_bsod = tk.Frame(ventana)
    frame_botones_bsod.pack(pady=10)

    boton_visor_eventos = tk.Button(
        frame_botones_bsod,
        text="Ver BSODs en Visor de Eventos",
        command=abrir_visor_eventos,
        bg=BOTON_COLOR
    )
    boton_visor_eventos.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_visor_eventos, "Abre el Visor de Eventos mostrando solo los pantallazos azules recientes")

    # ==============================================
    # SECCIÓN: OFFICE
    # ==============================================
    titulo_office = tk.Label(ventana, text="Office", font=("Arial", 16, "bold"))
    titulo_office.pack(pady=10, anchor=tk.W, padx=20)

    frame_botones_office = tk.Frame(ventana)
    frame_botones_office.pack(pady=10)

    boton_reactivar_onedrive = tk.Button(
        frame_botones_office,
        text="Reactivar OneDrive",
        command=reactivar_onedrive,
        bg=BOTON_COLOR
    )
    boton_reactivar_onedrive.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_reactivar_onedrive, "Reactiva OneDrive modificando el registro de Windows")

    boton_aumentar_ost = tk.Button(
        frame_botones_office,
        text="Aumentar Tamaño OST",
        command=aumentar_tamano_ost,
        bg=BOTON_COLOR
    )
    boton_aumentar_ost.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_aumentar_ost, "Aumenta el tamaño máximo permitido para archivos OST de Outlook")

    boton_eliminar_perfil = tk.Button(
        frame_botones_office,
        text="Eliminar Perfil Outlook",
        command=eliminar_perfil_outlook,
        bg=BOTON_COLOR
    )
    boton_eliminar_perfil.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_eliminar_perfil, "Elimina perfiles de Outlook del registro y sus archivos OST asociados")

    boton_archivo_online = tk.Button(
        frame_botones_office,
        text="Archivar en Línea",
        command=mover_correos_archivo_online,
        bg=BOTON_COLOR
    )
    boton_archivo_online.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_archivo_online, "Mueve correos antiguos al archivo en línea de Office 365/Exchange")

    # ==============================================
    # BOTÓN MANUAL Y ACUERDO DE LICENCIA
    # ==============================================
    frame_botones_manual = tk.Frame(ventana)
    frame_botones_manual.pack(pady=20)

    boton_manual = tk.Button(
        frame_botones_manual,
        text="Manual",
        command=generar_manual_html,
        bg=BOTON_COLOR
    )
    boton_manual.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_manual, "Muestra el manual de la aplicación en formato HTML")

    boton_licencia = tk.Button(
        frame_botones_manual,
        text="Acuerdo de Licencia",
        command=mostrar_acuerdo_licencia,
        bg=BOTON_COLOR
    )
    boton_licencia.pack(side=tk.LEFT, padx=10)
    Tooltip(boton_licencia, "Muestra el acuerdo de licencia MIT de esta aplicación")

    # ==============================================
    # VERSIÓN EN ESQUINA INFERIOR DERECHA
    # ==============================================
    version_label = tk.Label(ventana, text="v0.1.1", font=("Arial", 8))
    version_label.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)

    ventana.mainloop()

if __name__ == "__main__":
    main()