from modelo import Modelo
from vista import Vista
import tkinter.messagebox as messagebox
import pandas as pd
import shutil
import os

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
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo4(total_chilefilms)
    
    def procesar_archivo_tipo5(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo3()
        self.modelo.insertar_en_standard_tipo5(total_chilefilms)
    
    def procesar_archivo_tipo6(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo6(total_chilefilms)

    def procesar_archivo_tipo7(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo2()
        self.modelo.insertar_en_standard_tipo7(total_chilefilms)

    def procesar_archivo_tipo8(self, ruta_archivo):
        self.modelo.cargar_archivo(ruta_archivo)
        total_chilefilms = self.modelo.buscar_total_tipo4()
        self.modelo.insertar_en_standard_tipo8(total_chilefilms)

    def subir_ajuste(self, ruta_archivo):
        self.modelo.subir_ajuste(ruta_archivo)
    
    def eliminar_archivos_ajustes(self):
        self.modelo.eliminar_archivos_ajustes()
    
    def insertar_valor_en_celda(self, fila, columna, valor):
        self.modelo.insertar_valor_en_celda(fila, columna, valor)

    def mostrar_proceso_finalizado(self, tipo):
        self.vista.mostrar_proceso_finalizado(tipo)

    def mostrar_error(self, mensaje):
        self.vista.mostrar_error(mensaje)