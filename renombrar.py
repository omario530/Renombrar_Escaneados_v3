from tkinter import *
from tkinter import ttk
import tkinter as tk
import os
from datetime import datetime
import webbrowser
import PyPDF2
import shutil
import subprocess


# Rutas de carpeta
entregados = r'escaneados'
copiados = r'renombrados'

# Listas
Archivo = []
Renombrado = []

# Variables
Actual = 0
content = ""
ini_pdf = ""
year = 0
totalPages = 0


# Extraer los nombres de los archivos
def escaneados():
    
    result_generator = os.walk(entregados)
    files_result = [x for x in result_generator]


    for folder, dir_list, file_list in files_result:
        # Obtener los nombres de cada archivo
        Archivo.extend([os.path.join(file) for file in file_list])

    global Actual
    Actual = 1

# Funcion para el valor seleccionado en el combobox
def on_select(event):
    global ini_pdf
    dg = selected_value.get()
    if dg == 'DGGC':
        ini_pdf = "ASEA_UGSIVC_DGGC_"

    elif dg == 'DGGOI':
        ini_pdf = "ASEA_UGI_DGGOI_"

    elif dg == 'DGGPI':
        ini_pdf = "ASEA_UGI_DGGPI_"

    elif dg == 'DGGEERC':
        ini_pdf = "ASEA_UGI_DGGEERC_"

    elif dg == 'DGGEERNCM':
        ini_pdf = "ASEA_UGI_DDGGEERNCM_"
    
    elif dg == 'DGGEERNCT':
        ini_pdf = "ASEA_UGI_DGGEERNCT_"

    elif dg == 'UGI':
        ini_pdf = "ASEA_UGI_"

    Inicio.set(ini_pdf)

def oficio_capturado(var):
    global content
    content= var.get()
    Inicio.set(ini_pdf + content + "_" + str(year) + ".pdf")

def ano_ingresado(ano_1):
    global year
    captura= ano_1.get()
    year = captura
    Inicio.set(ini_pdf + str(content) + "_" + str(captura) + ".pdf")

def renombrar():
    global Actual
    rep = 0

    # Extraer los nombres de los archivos renombrados
    result_generator = os.walk(copiados)
    files_result = [x for x in result_generator]

    for folder, dir_list, file_list in files_result:
        # Obtener los nombres de cada archivo
        Renombrado.extend([os.path.join(file) for file in file_list])


        leyenda = str(len(Archivo)) + " de " + str(len(Archivo))
        Conteo.config(text=leyenda)
        # llamar al pdf actual
        n = Actual
        anterior = os.path.join(entregados, Archivo[n-1])
        nuevo_1 = os.path.join(copiados, Nuevo_Nombre.get())

        # verificar si el nuevo nombre contiene .pdf
        if nuevo_1.find(".pdf") == -1:
            nuevo = nuevo_1 + ".pdf"
        else:
            nuevo = nuevo_1
        
        # verificar que no se repita el nombre
        for x in Renombrado:
            if os.path.join(copiados, x) == nuevo:
                rep = rep + 1

        if rep < 1 and var.get() != "Repetido":
            # copiar el archivo con un nuevo nombre a la carpeta de renombrados
            shutil.copyfile(anterior, nuevo)

            # moverse al siguiente pdf en la lista Archivo
            Actual = n + 1

            # llamar rutinas
            cuenta()
            abrir_pdf()

            # limpiar entrada
            var.set("")

        else:
            var.set("Repetido")

# Hipervinculo    
def callback(url):
    webbrowser.open_new(url)

# abrir carpeta escaneados
def carpeta_escaneados():
    subprocess.run(['xdg-open', entregados])

# abrir carpeta renombrados
def carpeta_renombrados():
    subprocess.run(['xdg-open', copiados])


# Parametros del formulario ----------------------------------------------------
root = tk.Tk()
root.geometry('655x370')
root.title('Renombrar Oficios Escaneados')

frm = ttk.Frame(root, padding=10)
frm.grid()

# Varible para almacenar la captura del usuario
Inicio = tk.StringVar()
ano_1 = StringVar()
ano_1.trace_add("write", lambda name, index,mode, var=ano_1: ano_ingresado(var))

var = StringVar()
var.trace_add("write", lambda name, index,mode, var=var: oficio_capturado(var))


# Conteo
Conteo = ttk.Label(frm, width=30, font=('Helvetica', 12))
Conteo.grid(column=0, row=0, sticky = tk.W)
# Separador
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=1, row=1, sticky=tk.W)


# Archivo actual
ttk.Label(frm, text="Archivo actual:", font=('Helvetica', 10, 'underline')).grid(column=0, row=2, sticky=tk.W)

paginas = ttk.Label(frm, width=30, font=('Helvetica', 12))
paginas.grid(column=1, row=2, sticky=tk.W)
paginas.config(foreground ='green')

Archivo_actual = ttk.Label(frm, width=50, font=('Helvetica', 12))
Archivo_actual.grid(column=0, row=3, sticky = tk.W+tk.E, columnspan=3)
Archivo_actual.config(foreground ='red')

# Separador
#ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=3, sticky=tk.W)
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=4, sticky=tk.W)

ttk.Label(frm, text="Generar nombre a partir de DG, Folio y Año", font=('Helvetica', 10, 'underline')).grid(column=0, row=5, sticky = tk.W+tk.E, columnspan=3)


# Inicio del oficio
ttk.Label(frm, text="Seleccione la DG", font=('Helvetica', 12)).grid(column=0, row=6)

selected_value = tk.StringVar()
combobox = ttk.Combobox(frm, textvariable=selected_value, width=15, font=('Helvetica', 12))
combobox.grid(column=0, row=7)
combobox['values'] = ('DGGC', 'DGGOI', 'DGGPI', 'DGGEERC', 'DGGEERNCM', 'DGGEERNCT', 'UGI')
combobox['state'] = 'readonly'
combobox.bind('<<ComboboxSelected>>', on_select)


# Numero de folio
ttk.Label(frm, text="Escriba el folio", font=('Helvetica', 12)).grid(column=1, row=6)
folio = ttk.Entry(frm, width=15, font=('Helvetica', 12, "bold"),textvariable=var).grid(column=1, row=7)


# Ano
ttk.Label(frm, text="Año", width=8, font=('Helvetica', 12)).grid(column=2, row=6)
ano = ttk.Entry(frm, width=8, font=('Helvetica', 12, "bold"),textvariable = ano_1)
ano.grid(column=2, row=7)
# Separador
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=8, sticky=tk.W)
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=9, sticky=tk.W)


# Nuevo nombre
ttk.Label(frm, text="Nuevo nombre:", font=('Helvetica', 10, 'underline')).grid(column=0, row=10, sticky=tk.W)
Nuevo_Nombre = ttk.Entry(frm, font=('Helvetica', 12), textvariable = Inicio)
Nuevo_Nombre.grid(column=0, row=11, sticky = tk.W+tk.E, columnspan=3)
# Separador
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=12, sticky=tk.W)


# Botones --------------------------------------
# Renombrar
ttk.Button(frm, text="Renombrar", command=renombrar).grid(column=0, row=13)

# Salir
ttk.Button(frm, text="Salir", command=root.destroy).grid(column=2, row=13)


# carpeta_escaneados
ttk.Button(frm, text="escaneados", command=carpeta_escaneados).grid(column=1, row=0)

# carpeta_renombrados
ttk.Button(frm, text="renombrados", command=carpeta_renombrados).grid(column=2, row=0)



# Credito
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=14, sticky=tk.W)
ttk.Label(frm, text="                                        ", font=('Helvetica', 12)).grid(column=0, row=15, sticky=tk.W)
credito = ttk.Label(frm, text="Por Omar Escamilla Gutiérrez para la ASEA", font=('Helvetica', 10, 'underline'))
credito.config(foreground ='blue')
credito.grid(column=0, row=16, sticky = tk.W+tk.E, columnspan=3)
credito.bind("<Button-1>", lambda e: callback("www.linkedin.com/in/omar-escamilla-gutierrez-b1457396"))

def ano_texto(value):
    ano.delete(0, tk.END)  # Clear any existing value
    ano.insert(0, value)   # Insert the new value
    global year
    year = value


def cuenta():
    leyenda = str(Actual) + " de " + str(len(Archivo))
    Conteo.config(text=leyenda)
    

def abrir_pdf():
    n = Actual

    #contenido de la carpeta
    contents = os.listdir(entregados)

    if len(contents) == 0:
        Archivo_actual.config(text="Sin archivos a renombrar")


    # verificar si se alcanzo el total de archivos
    elif Actual > len(Archivo):
        Archivo_actual.config(text="Acabamos")
        leyenda = str(len(Archivo)) + " de " + str(len(Archivo))
        Conteo.config(text=leyenda)

    elif n-1 < len(Archivo):
        archivo_abrir = os.path.join(entregados, Archivo[n-1])

        # Contar las paginas
        file = open(archivo_abrir, 'rb')
        pdfReader = PyPDF2.PdfReader(file)
        global totalPages
        totalPages = len(pdfReader.pages)

        leyenda = str(Archivo[n-1])
        Archivo_actual.config(text=leyenda)
        
        paginas.config(text="págs: " + str(totalPages))
        # Abrir el PDF
        webbrowser.open_new(archivo_abrir)


escaneados()
cuenta()
abrir_pdf()

# Ejecutar formulario
ano_texto(datetime.now().year)


root.mainloop()