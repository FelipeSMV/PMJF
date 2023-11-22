from modelo import Modelo
from vista import Vista
import tkinter.messagebox as messagebox

class Controlador:
    def __init__(self, vista):
        self.modelo = Modelo(vista)
        self.vista = vista
        self.modelo.vista = vista
        self.vista.establecer_controlador(self)

    def iniciar_aplicacion(self):
        self.vista.iniciar()

    def procesar_archivo_tipo1(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo1()
        self.modelo.insertar_en_standard_tipo1(total_chilefilms)

    def procesar_archivo_tipo2(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo2(total_chilefilms)

    def procesar_archivo_tipo3(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo3()
        self.modelo.insertar_en_standard_tipo3(total_chilefilms)

    def procesar_archivo_tipo4(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_tipo4_rango1, total_tipo4_rango2 = self.modelo.buscar_total_tipo4()
        self.modelo.insertar_en_planilla_de_prueba(total_tipo4_rango1, total_tipo4_rango2)



