from tkinter import messagebox
from modelo import Modelo
from vista import Vista

class Controlador:
    def __init__(self, modelo, vista):
        self.modelo = modelo
        self.vista = vista

    def leer_excel(self, ruta_archivo):
        self.modelo.leer_excel(ruta_archivo)

    def generar(self):
        if self.modelo.data is not None:
            total = self.modelo.procesar_datos()
            
            # Imprimir el resultado en la terminal
            print(f"Resultado de la suma: {total}")

            self.modelo.actualizar_planilla(total)
            messagebox.showinfo("Aviso", "Generando planilla standard y guardando copia en el escritorio")
        else:
            messagebox.showerror("Error", "Por favor, suba un archivo Excel")
