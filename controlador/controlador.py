from modelo import Modelo
from vista import Vista

class Controlador:
    def __init__(self):
        self.modelo = Modelo()
        self.vista = Vista(self)

    def iniciar_aplicacion(self):
        self.vista.iniciar()

    def procesar_archivo(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total()
        self.modelo.insertar_en_standard(total_chilefilms)

    def generar(self):
        print("Realizando la generación...")

if __name__ == "__main__":
    controlador = Controlador()
    controlador.iniciar_aplicacion()
