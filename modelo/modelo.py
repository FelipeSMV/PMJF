import pandas as pd
import shutil
import os
from openpyxl import load_workbook

class Modelo:
    def __init__(self):
        self.archivo_subido = None
        self.archivo_standard = "recursos/planilla_standard.xlsx"
        self.columna_standard_index = 2 
        self.columnas_a_insertar = [4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 35, 36, 37, 38, 39, 40, 41, 43, 46, 47, 48, 49, 50, 51, 52, 56, 57, 58, 59, 60, 61,63]
        self.filas_a_pegar = [6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 37, 38, 39, 40, 41, 42, 43, 45, 48, 49, 50, 51, 52, 53, 54, 58, 59, 60, 61, 62, 63, 65]

    def cargar_archivo(self, ruta_archivo):
        try:
            # Leer el archivo Excel y almacenar el objeto ExcelFile
            self.archivo_subido = pd.ExcelFile(ruta_archivo)
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")

    def buscar_total(self):
        if self.archivo_subido is not None:
            # Verificar si la hoja "Estado" y la columna "TOTAL" existen
            if "Estado" in self.archivo_subido.sheet_names and "TOTAL" in self.archivo_subido.parse("Estado").columns:
                # Extraer los datos de las filas especificadas
                return self.archivo_subido.parse("Estado")["TOTAL"].iloc[self.columnas_a_insertar]
            else:
                print("No se encontró la hoja 'Estado' o la columna 'TOTAL' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def insertar_en_standard(self, total_chilefilms):
        try:

            archivo_standard = load_workbook(self.archivo_standard)

           
            if "Estado" in archivo_standard.sheetnames:
                # Obtener la hoja "Estado"
                hoja_estado = archivo_standard["Estado"]

                # Asegurar que al menos una hoja esté visible
                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                # Obtener la columna en la posición del índice 3 en la planilla estándar
                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 1).value for fila in range(1, hoja_estado.max_row + 1)]

                # Actualizar solo las celdas específicas
                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 1}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 1, value=valor)

                # Guardar el archivo actualizado
                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard.")

                # Crear una copia en el escritorio
                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")
        except Exception as e:
            print(f"Error al insertar en el archivo standard: {e}")
