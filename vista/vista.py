import tkinter as tk
from tkinter import filedialog

class Vista:
    def __init__(self, modelo, controlador):
        self.modelo = modelo
        self.controlador = controlador
        self.iniciar_interfaz()

    def iniciar_interfaz(self):
        self.ventana = tk.Tk()
        self.ventana.title("Bienvenido")
        self.ventana.geometry("400x300")

        self.boton_subir = tk.Button(self.ventana, text="Subir archivo", command=self.subir_archivo)
        self.boton_subir.pack()

        self.boton_generar = tk.Button(self.ventana, text="Generar", command=self.controlador.generar)
        self.boton_generar.pack()

        self.ventana.mainloop()

    def subir_archivo(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta_archivo:
            self.controlador.leer_excel(ruta_archivo)
