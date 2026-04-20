import importlib
import json
import os
import pyperclip
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
# Instalar librerias
librerias = {"pandas": "pandas", "pyautogui": "pyautogui", "openpyxl": "openpyxl"}
for modulo, paquete in librerias.items():
    try:
        importlib.import_module(modulo)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", paquete])
import pandas as pd
import pyautogui
# Funcion que muestra las instrucciones
def mostrar_instrucciones(ventana_origen):
    ventana = tk.Toplevel(ventana_origen)
    ventana.title("Instrucciones")
    ventana.geometry("500x400")
    ventana.configure(bg="#D3D3D3")
    ventana.transient(ventana_origen)
    ventana.grid_columnconfigure(0, weight=1)
    ventana.grid_rowconfigure(0, weight=1)
    frame = tk.Frame(ventana, bg="#D3D3D3")
    frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    texto = tk.Text(frame, wrap="word", font=("Segoe UI", 10), bg="white", relief="flat")
    texto.grid(row=0, column=0, sticky="nsew")
    scroll = tk.Scrollbar(frame, command=texto.yview)
    scroll.grid(row=0, column=1, sticky="ns")
    texto.config(yscrollcommand=scroll.set)
    texto.tag_configure("titulo", font=("Segoe UI", 12, "bold"), foreground="#1E90FF")
    texto.tag_configure("negrita", font=("Segoe UI", 10, "bold"))
    texto.insert("end", "INSTRUCCIONES DE USO\n\n", "titulo")
    texto.insert("end", "1. Parametros", "negrita")
    texto.insert("end", "Hay que seleccionar tanto el archivo excel con las instrucciones como la carpeta con archivos con los que se desea trabajar\n")
    texto.insert("end", "2. Uso del programa", "negrita")
    texto.insert("end", "El programa contiene dos botones uno que muestra las coordenadas del raton y otro que ejecuta unas instrucciones con el fin de automatizar el movimiento del raton y realizar algunas operaciones con teclado y comandos especiales.\n")
    texto.insert("end", "3. Detalles:", "negrita")
    texto.insert("end", "La columna indice representa en que orden se realizaran las acciones\nX e Y son las coordenadas a las que se movera el raton\n")
    texto.insert("end", "Accion es la accion que se ejecutara\nTiempo de espera es el tiempo que espera el programa despues de hacer la accion\n")
    texto.insert("end", "Bucle es el paso al que se volvera cuando se usen las opciones Bucle archivos y Bucle infinito\nComando es un atajo de teclado que se ejcutara si son comandos concatenados se separa por , y dentro del mismo comando por +. Ej:Ctrl+V,Ctrl+K")
    texto.insert("end", "4. Botones", "negrita")
    texto.insert("end", "Pulsa Detectar para ver las coordenadas del ratón. Pulsa Ejecutar para realizar las operaciones indicadas en el excel\n")
    texto.config(state="disabled")
    tk.Button(ventana, text="Cerrar", command=ventana.destroy, bg="#1E90FF", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=15, pady=5, cursor="hand2").grid(row=1, column=0, pady=(0, 15))
# Funcion que devuelve el listado de archivos de una carpeta y subcarpetas (busqueda profunda)
def obtener_archivos(carpeta):
    return [os.path.join(raiz, nombre) for raiz, _, nombres_archivos in os.walk(carpeta) for nombre in nombres_archivos]
# Funcion que ejecuta las instrucciones correctas segun el tipo de comando solicitado
def comandos_automatizados(tipo, parametros):
    if tipo == 'Click Izquierdo':
        pyautogui.moveTo(parametros[0], parametros[1])
        pyautogui.click()
        time.sleep(parametros[2])
    elif tipo == 'Click Derecho':
        pyautogui.moveTo(parametros[0], parametros[1])
        pyautogui.rightClick()
        time.sleep(parametros[2])
    elif tipo == 'Doble Click':
        pyautogui.moveTo(parametros[0], parametros[1])
        pyautogui.doubleClick()
        time.sleep(parametros[2])
    elif tipo == 'Mover':
        pyautogui.moveTo(parametros[0], parametros[1])
        time.sleep(parametros[2])
    elif tipo == 'No hacer nada':
        time.sleep(parametros[2])
    elif tipo == 'Pulsar Enter':
        pyautogui.press('enter')
        time.sleep(parametros[2])
    elif tipo == 'Seleccionar archivo actual':
        pyperclip.copy(parametros[3])
        time.sleep(parametros[2])
    elif tipo == 'Comando':
        for combo in parametros[3].split(","):
            pyautogui.hotkey(*combo.strip().split("+"))
            print("Combo", *combo.strip().split("+"))
        time.sleep(parametros[2])
# Funcion que llama a la funcion de comandos y actualiza, si es necesario, la lista de archivos
def accion_fila(accion, ruta_archivos, parametros, archivos):
    if accion == 'Seleccionar archivo actual' and archivos == [-1]:
        archivos[:] = obtener_archivos(ruta_archivos)
    if accion == 'Seleccionar archivo actual':
        parametros[3] = archivos[0]
        archivos.pop(0)
    comandos_automatizados(accion, parametros)
# Funcion que ejecuta las acciones especificadas
def main_ejecutar(ventana):
    ruta_ordenes, ruta_archivos, archivos = ventana.ruta_ordenes.get(), ventana.ruta_archivos.get(), [-1]
    raton, comandos = pd.read_excel(ruta_ordenes, sheet_name = 'Raton', dtype=str, keep_default_na=False).replace("", "0"), pd.read_excel(ruta_ordenes, sheet_name = 'Comandos especiales', dtype=str, keep_default_na=False).replace("", "0")
    raton['Comando'] = ''
    comandos['X'], comandos['Y'], comandos['Accion'], comandos['Bucle'] = '0', '0', 'Comando', ''
    datos = pd.concat([raton, comandos], ignore_index=True)
    datos['Indice'] = datos['Indice'].astype(int)
    datos = datos.sort_values(by='Indice')
    for indice, fila in datos.iterrows():
        print(fila)
        accion = fila['Accion']
        parametros = [int(fila['X']), int(fila['Y']), float(fila['Tiempo espera']), fila['Comando']]
        if accion != 'Bucle por archivos' and accion != 'Bucle infinito':
            accion_fila(accion, ruta_archivos, parametros, archivos)
        elif accion == 'Bucle por archivos':
            indice_minimo, indice_maximo = int(fila['Bucle']), int(fila['Indice'])
            while archivos:
                for indice_bucle in range(indice_minimo, indice_maximo):
                    fila_bucle = datos[datos['Indice'] == indice_bucle]
                    print(fila_bucle)
                    if not fila_bucle.empty:
                        fila_bucle = fila_bucle.iloc[0]
                        accion_bucle = fila_bucle['Accion']
                        parametros = [int(fila_bucle['X']), int(fila_bucle['Y']), float(fila_bucle['Tiempo espera']), fila_bucle['Comando']]
                        accion_fila(accion_bucle, ruta_archivos, parametros, archivos)
        elif accion == 'Bucle infinito':
            indice_minimo, indice_maximo = int(fila['Bucle']), int(fila['Indice'])
            while True:
                for indice_bucle in range(indice_minimo, indice_maximo):
                    fila_bucle = datos[datos['Indice'] == indice_bucle]
                    print(fila_bucle)
                    if not fila_bucle.empty:
                        fila_bucle = fila_bucle.iloc[0]
                        accion_bucle = fila_bucle['Accion']
                        parametros = [int(fila_bucle['X']), int(fila_bucle['Y']), float(fila_bucle['Tiempo espera']), fila_bucle['Comando']]
                        accion_fila(accion_bucle, ruta_archivos, parametros, archivos)
    messagebox.showinfo("Completado", "El programa ha finalizado correctamente")
    return 1
# Funcion que muestra la posicion del raton
def main_detectar(ventana):
    while True and ventana.detener:
        x, y = pyautogui.position()
        ventana.texto_coordenadas.set(f"\rX: {x:<5} Y: {y:<5}")
        time.sleep(0.05)
# Clase que crea la App
class App(tk.Tk):
    # Funcion que inicializa la App
    def __init__(self):
        super().__init__()
        self.title("Automatizador")
        self.geometry("550x430")
        self.configure(bg = "#D3D3D3")
        # Variables
        self.ruta = os.path.join(os.environ["APPDATA"], "MiApp", "config.json")
        self.ruta_ordenes = tk.StringVar()
        self.ruta_archivos = tk.StringVar()
        self.texto_coordenadas = tk.StringVar(value = "Coordenadas")
        self.detener = False
        self._cargar_config()
        self.ruta_ordenes.trace_add("write", lambda *args: self._guardar_config())
        self.ruta_archivos.trace_add("write", lambda *args: self._guardar_config())
        # Partes de la App
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 0)
        self.grid_rowconfigure(1, weight = 0)
        self.grid_rowconfigure(2, weight = 0)
        self._crear_header()
        self._crear_cuerpo()
        self._crear_footer()
    # Funcion que crea el header
    def _crear_header(self):
        header = tk.Frame(self, bg = "#4A4A4A", height = 50)
        header.grid(row = 0, column = 0, sticky = "ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight = 1)
        titulo = tk.Label(header, text = "Automatizador", bg = "#4A4A4A" , fg = "white", font = ("Segoe UI", 14, "bold"))
        titulo.grid(row = 0, column = 0, pady = 10)
    # Funcion que crea el cuerpo
    def _crear_cuerpo(self):
        cuerpo = tk.Frame(self, bg = "#D3D3D3")
        cuerpo.grid(row = 1, column = 0, sticky = "nsew", padx = 20, pady = 20)
        cuerpo.grid_columnconfigure(0, weight = 1)
        cuerpo.grid_rowconfigure(0, weight = 0)
        cuerpo.grid_rowconfigure(1, weight = 0)
        cuerpo.grid_rowconfigure(2, weight = 0)
        # Entrada ruta ordenes
        label_ordenes = tk.Label(cuerpo, text = "Selecciona el archivo .xlsx:", bg = "#D3D3D3", font = ("Segoe UI", 10))
        label_ordenes.grid(row = 0, column = 0, sticky = "w", pady = (0, 5))
        campo_ordenes = tk.Entry(cuerpo, textvariable = self.ruta_ordenes, font = ("Segoe UI", 11), relief = "flat", highlightthickness = 1, highlightbackground = "#888", highlightcolor = "#1E90FF")
        campo_ordenes.grid(row = 1, column = 0, sticky = "ew", ipady = 5)
        boton = tk.Button(cuerpo,text = "Buscar archivo",command = self._seleccionar_xlsx)
        boton.grid(row = 1, column = 1, padx = 5)
        # Entrada ruta archivos
        label_archivo = tk.Label(cuerpo, text = "Selecciona la carpeta con los archivos:", bg = "#D3D3D3", font = ("Segoe UI", 10))
        label_archivo.grid(row = 2, column = 0, sticky = "w", pady = (0, 5))
        campo_archivos = tk.Entry(cuerpo, textvariable = self.ruta_archivos, font = ("Segoe UI", 11), relief = "flat", highlightthickness = 1, highlightbackground = "#888", highlightcolor = "#1E90FF")
        campo_archivos.grid(row = 3, column = 0, sticky = "ew", ipady = 5)
        boton = tk.Button(cuerpo,text = "Buscar carpeta",command = self._seleccionar_carpeta)
        boton.grid(row = 3, column = 1, padx = 5)
        # Etiqueta que muestra las coordenadas
        etiqueta_coordenadas = tk.Label(cuerpo, bg = "#D3D3D3", font = ("Segoe UI", 10), textvariable = self.texto_coordenadas)
        etiqueta_coordenadas.grid(row = 4, column = 0, sticky = "w", pady = (0, 5))
        # Botones
        contenedor_botones = tk.Frame(cuerpo, bg = "#D3D3D3")
        contenedor_botones.grid(row = 5, column = 0, pady = 20)
        contenedor_botones.grid_columnconfigure(0, weight = 1)
        contenedor_botones.grid_columnconfigure(1, weight = 1)
        self.boton_detectar = self._crear_boton_funcion(contenedor_botones, "Detectar", lambda: self._ejecutar_boton('Detectar'))
        self.boton_detectar.grid(row = 0, column = 0, padx = 10)
        self.boton_ejecutar = self._crear_boton_funcion(contenedor_botones, "Ejecutar", lambda: self._ejecutar_boton('Ejecutar'))
        self.boton_ejecutar.grid(row = 0, column = 1, padx = 10)
    # Funcion que crea el footer
    def _crear_footer(self):
        footer = tk.Frame(self, bg="#4A4A4A", height=120)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)
        footer.grid_columnconfigure(0, weight=1)
        boton_info = tk.Button(footer, text="ℹ", command=lambda: mostrar_instrucciones(self), bg="#4A4A4A", fg="white", font=("Segoe UI", 16, "bold"), relief="flat",  borderwidth=0, cursor="hand2", activebackground="#4A4A4A", activeforeground="#1E90FF")
        boton_info.grid(row=0, column=0, sticky="e", padx=10)
    # Funcion para crear los botones que ejecutan cosas
    def _crear_boton_funcion(self, parent, texto, comando):
        boton = tk.Button(parent, text = texto, command = comando, bg = "#1E90FF", fg = "white", activebackground = "#187BCD", activeforeground = "white", relief = "flat", font = ("Segoe UI", 10, "bold"), padx = 15, pady = 6, cursor = "hand2", borderwidth = 0)
        boton.bind("<Enter>", self._color_entrada)
        boton.bind("<Leave>", self._color_salida)
        return boton
    # Funcion para seleccionar el color del boton al estar encima de él
    def _color_entrada(self, event):
        if event.widget is getattr(self, "boton_detectar", None) and self.boton_detectar['state'] == "normal":
            self.boton_detectar['bg'] = 'red' if self.detener else "#187BCD"
        elif event.widget is getattr(self, "boton_ejecutar", None) and self.boton_ejecutar['state'] == "normal":
            self.boton_ejecutar['bg'] = 'red' if self.detener else "#187BCD"
    # Funcion para seleccionar el color del boton al dejar de estar encima de él
    def _color_salida(self, event):
        if event.widget is getattr(self, "boton_detectar", None) and self.boton_detectar['state'] == "normal":
            self.boton_detectar['bg'] = 'red' if self.detener else "#1E90FF"
        elif event.widget is getattr(self, "boton_ejecutar", None) and self.boton_ejecutar['state'] == "normal":
            self.boton_ejecutar['bg'] = 'red' if self.detener else "#1E90FF"
    # Funcion para cargar los ultimos valores usados en los parametros
    def _cargar_config(self):
        if os.path.exists(self.ruta):
            with open(self.ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
            self.ruta_ordenes.set(datos.get("ruta_ordenes", ""))
            self.ruta_archivos.set(datos.get("ruta_archivos", ""))
    # Funcion para guardar los nuevos valores de los parametros
    def _guardar_config(self):
        datos = {"ruta_ordenes": self.ruta_ordenes.get(), "ruta_archivos": self.ruta_archivos.get()}
        os.makedirs(os.path.dirname(self.ruta), exist_ok=True)
        with open(self.ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
    # Función para abrir el explorador y seleccionar un archivo
    def _seleccionar_xlsx(self):
        archivo = filedialog.askopenfilename(title = "Seleccionar archivo Excel", filetypes = [("Archivos Excel", "*.xlsx")])
        if archivo:
            self.ruta_ordenes.set(archivo)
    # Funcion para abrir el explorador y seleccionar un archivo
    def _seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar la carpeta con los archivos")
        if carpeta:
            self.ruta_archivos.set(carpeta)
    # Funcion que se ejecuta al pulsar un boton
    def _ejecutar_boton(self, tipo):
        if tipo == 'Detectar':
            hilo = threading.Thread(target = self._ejecutar_detectar, daemon = True)
            hilo.start()
        elif tipo == 'Ejecutar':
            hilo = threading.Thread(target = self._ejecutar_ejecutar, daemon = True)
            hilo.start()
    # Funcion que se ejecuta al pulsar el boton detectar
    def _ejecutar_detectar(self):
        self.detener = not self.detener
        self.boton_detectar['text'] = 'Detener' if self.detener else 'Detectar'
        self.boton_detectar['bg'] = 'red' if self.detener else "#187BCD"
        self.boton_ejecutar['state'] =  'disabled' if self.detener else "normal"
        if self.detener:
            main_detectar(self)
    # Funcion que se ejecuta al pulsar el boton ejecutar
    def _ejecutar_ejecutar(self):
        self.boton_detectar['state'] =  'disabled'
        self.boton_ejecutar['state'] =  'disabled'
        messagebox.showwarning("Aviso", "El programa se iniciara en 10 segundos")
        time.sleep(10)
        try:
            main_ejecutar(self)
        except Exception as error:
            messagebox.showerror("Error", f"Ha ocurrido un error:\n\n{error}")
            self.boton_detectar['state'] =  "normal"
            self.boton_ejecutar['state'] =  "normal"
if __name__ == "__main__":
    ventana = App()
    ventana.mainloop()