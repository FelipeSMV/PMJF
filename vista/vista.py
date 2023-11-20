from tkinter import Tk, Label, Button, filedialog

class Vista:
    def __init__(self, controlador):
        self.controlador = controlador
        self.ventana = Tk()
        self.ventana.title("Proyecto MJF")

        self.etiqueta = Label(self.ventana, text="Bienvenido al Proyecto")
        self.etiqueta.pack()

        self.boton_subir_tipo1 = Button(self.ventana, text="Chilefilms", command=self.subir_archivo_tipo1)
        self.boton_subir_tipo1.pack()

        self.boton_subir_tipo2 = Button(self.ventana, text="CHF Interacional", command=self.subir_archivo_tipo2)
        self.boton_subir_tipo2.pack()

        self.boton_subir_tipo3 = Button(self.ventana, text="Cine color", command=self.subir_archivo_tipo3)
        self.boton_subir_tipo3.pack()

    def iniciar(self):
        self.ventana.mainloop()

    def subir_archivo_tipo1(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        self.controlador.procesar_archivo_tipo1(ruta_archivo)

    def subir_archivo_tipo2(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        self.controlador.procesar_archivo_tipo2(ruta_archivo)

    def subir_archivo_tipo3(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        self.controlador.procesar_archivo_tipo3(ruta_archivo)