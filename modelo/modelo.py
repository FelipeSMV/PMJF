import pandas as pd
import shutil
import os
from openpyxl import load_workbook

class Modelo:
    def __init__(self):
        self.archivo_subido = None
        self.archivo_standard = "recursos/planilla_standard.xlsx"
        self.columnas_a_insertar = [6, 7, 8, 9, 10, 11, 12, 13, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 37, 38, 39, 40, 41, 42, 43, 48, 49, 50, 51, 52, 53, 54, 58, 59, 60, 61, 62, 63, 65]

    def cargar_archivo(self, ruta_archivo):
        try:
            self.archivo_subido = pd.read_excel(ruta_archivo, sheet_name="Estado")
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")

    def buscar_total(self):
        if self.archivo_subido is not None:
            columna_total = self.archivo_subido["TOTAL"]
            return columna_total.iloc[self.columnas_a_insertar]

    def insertar_en_standard(self, total_chilefilms):
        try:
            archivo_standard = pd.read_excel(self.archivo_standard, sheet_name="Estado")

           
            for fila, columna in zip(self.columnas_a_insertar, total_chilefilms.index):
                valor = total_chilefilms[columna]
                archivo_standard.loc[fila, "Chilefilms"] = valor

            
            archivo_standard.to_excel(self.archivo_standard, index=False, sheet_name="Estado", engine='openpyxl')

            
            escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
            nombre_copia = "planilla_standard_modificada.xlsx"
            ruta_copia = os.path.join(escritorio, nombre_copia)

            archivo_standard.to_excel(ruta_copia, index=False)

            print(f"Datos insertados correctamente en el archivo standard.")
            print(f"Copia guardada en el escritorio como {nombre_copia}")
        except Exception as e:
            print(f"Error al insertar en el archivo standard: {e}")
