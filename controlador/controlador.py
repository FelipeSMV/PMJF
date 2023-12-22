from modelo import Modelo
from vista import Vista
import tkinter.messagebox as messagebox
import pandas as pd


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
        total_chilefilms = self.modelo.buscar_total_resultado()
        self.modelo.insertar_en_standard_Resul(total_chilefilms)

    def procesar_archivo_tipo2(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo2(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado2()
        self.modelo.insertar_en_standard_Resul2(total_chilefilms)

    def procesar_archivo_tipo3(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo3()
        self.modelo.insertar_en_standard_tipo3(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado3()
        self.modelo.insertar_en_standard_Resul3(total_chilefilms)

    def procesar_archivo_tipo4(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo4(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado2()
        self.modelo.insertar_en_standard_Resul4(total_chilefilms)
    
    def procesar_archivo_tipo5(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo3()
        self.modelo.insertar_en_standard_tipo5(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado3()
        self.modelo.insertar_en_standard_Resul5(total_chilefilms)
    
    def procesar_archivo_tipo6(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo6(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado2()
        self.modelo.insertar_en_standard_Resul6(total_chilefilms)

    def procesar_archivo_tipo7(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo7(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado2()
        self.modelo.insertar_en_standard_Resul7(total_chilefilms)

    def procesar_archivo_tipo8(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo4()
        self.modelo.insertar_en_standard_tipo8(total_chilefilms)
        total_chilefilms = self.modelo.buscar_total_resultado4()
        self.modelo.insertar_en_standard_Resul8(total_chilefilms)
    
    def insertar_valor_en_celda(self, fila, columna, valor):
        exito = self.modelo.insertar_valor_en_celda(fila, columna, valor)
        return exito
    
    def insertar_valor_en_celda_R(self, fila, columna, valor):
        exito = self.modelo.insertar_valor_en_celda_R(fila, columna, valor)
        return exito
    
    def limpiar_planilla(self):
        try:
            exito = self.modelo.limpiar_planilla()
            if exito:
                self.vista.mostrar_mensaje("Planilla limpiada correctamente.")
            else:
                self.vista.mostrar_error("Error al limpiar la planilla.")
        except Exception as e:
            mensaje_error = f"Error al limpiar la planilla: {e}"
            print(mensaje_error)
            self.vista.mostrar_error(mensaje_error)

    def descargar_manual(self):
        try:
            self.modelo.descargar_manual()
            self.vista.mostrar_mensaje("Manual descargado correctamente.")
        except Exception as e:
            mensaje_error = f"Error al descargar el manual: {e}"
            print(mensaje_error)
            self.vista.mostrar_error(mensaje_error)