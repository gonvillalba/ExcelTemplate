import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.constants import BOTTOM
import recibir

# import excelConPython

ventana = tk.Tk()
ventana.geometry("400x300")


def seleccion():
    filetypes = (("excel files", "*.xlsx"), ("All files", "*.*"))
    global filename
    filename = fd.askopenfilename(
        title="Seleccionar archivo", initialdir="/", filetypes=filetypes
    )
    print(filename, len(filename))


def template():
    # excelConPython.toPath(filename)
    recibir.recibir(filename)
    print(filename, len(filename))


etiquta = tk.Label(ventana, text="Template 4G")
etiquta.pack()

hecho = tk.Label(ventana, text="Realizado por extgjv")
hecho.config(font=("Courier", 8))
hecho.pack(side=BOTTOM)

abrirButton = tk.Button(ventana, text="Seleccionar archivo", command=seleccion)
abrirButton.pack()

generarTemplate = tk.Button(ventana, text="Generar Template 4G", command=template)

generarTemplate.pack()

ventana.mainloop()
