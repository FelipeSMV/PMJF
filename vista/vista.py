from tkinter import Tk, Label, Button, filedialog

class Vista:
    def __init__(self, controlador):
        self.controlador = controlador
        self.ventana = Tk()
        self.ventana.title("Proyecto MVC")

        self.etiqueta = Label(self.ventana, text="Bienvenido al Proyecto MVC")
        self.etiqueta.pack()

        self.boton_subir = Button(self.ventana, text="Subir archivo", command=self.subir_archivo)
        self.boton_subir.pack()

        self.boton_generar = Button(self.ventana, text="Go", command=self.controlador.generar)
        self.boton_generar.pack()

    def iniciar(self):
        self.ventana.mainloop()

    def subir_archivo(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        self.controlador.procesar_archivo(ruta_archivo)
