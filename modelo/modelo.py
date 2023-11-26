import pandas as pd
import shutil
import os
from openpyxl import load_workbook

class Modelo:
    def __init__(self, vista):
        self.archivo_subido = None
        self.archivo_standard = "recursos/planilla_standard.xlsx"
        self.columna_standard_index = 2
        self.columnas_a_insertar = [8, 9, 10, 11, 12, 13, 14, 15, 17, 18, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 39, 40, 41, 42, 43, 44, 45, 47, 50, 51, 52, 53, 54, 55, 56, 60, 61, 62, 63, 64, 65]
        self.columnas_a_insertar_tipo2 = [4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 35, 36, 37, 38, 39, 40, 41, 43, 46, 47, 48, 49, 50, 51, 52, 56, 57, 58, 59, 60, 61, 63 ]
        self.filas_a_pegar = [6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 37, 38, 39, 40, 41, 42, 43, 45, 48, 49, 50, 51, 52, 53, 54, 58, 59, 60, 61, 62, 63, 65]
        self.vista = vista
        

    def cargar_archivo(self, ruta_archivo):
        try:
            
            self.archivo_subido = pd.ExcelFile(ruta_archivo)
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")


    def buscar_total_tipo1(self):
        if self.archivo_subido is not None:
            if "Estado" in self.archivo_subido.sheet_names:
                columna_subida = self.archivo_subido.parse("Estado").iloc[:, 2]  # Columna ubicada en el índice 3 (columna C en Excel)
                return columna_subida[self.columnas_a_insertar]
            else:
                print("No se encontró la hoja 'Estado' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")


    def buscar_total_tipo2(self):
        if self.archivo_subido is not None:
            if "Estado" in self.archivo_subido.sheet_names:
                columna_subida = self.archivo_subido.parse("Estado").iloc[:, 2]  # Columna ubicada en el índice 3 (columna C en Excel)
                return columna_subida[self.columnas_a_insertar_tipo2]
            else:
                print("No se encontró la hoja 'Estado' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def buscar_total_tipo3(self):
        if self.archivo_subido is not None:
            if "Estado" in self.archivo_subido.sheet_names:
                columna_subida = self.archivo_subido.parse("Estado").iloc[:, 24]  # Columna ubicada en el índice 25 (columna Y en Excel)
                return columna_subida[self.columnas_a_insertar_tipo2]
            else:
                print("No se encontró la hoja 'Estado' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def buscar_total_tipo4(self):
        if self.archivo_subido is not None:
            if "Ctas" in self.archivo_subido.sheet_names:
                hoja_ctas = self.archivo_subido.parse("Ctas")

                # Definir el rango 1 hasta la fila 320
                total_tipo4_rango1 = hoja_ctas.loc[:319, ["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO", "T/Cambio"]]

                # Definir el rango 2 desde la fila 320 hasta el final
                total_tipo4_rango2 = hoja_ctas.loc[320:, ["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO", "T/Cambio"]]

                return total_tipo4_rango1, total_tipo4_rango2
            else:
                print("No se encontró la hoja 'Ctas' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")
   
    def buscar_total_tipo5(self):
        if self.archivo_subido is not None:
            if "Ctas" in self.archivo_subido.sheet_names:
                hoja_ctas = self.archivo_subido.parse("Ctas")

                # Definir el rango 1 desde la fila 185 hasta la fila 314
                total_tipo5_rango1 = hoja_ctas.loc[185:314, ["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO", "T/Cambio"]]

                # Definir el rango 2 desde la fila 315 hasta la fila 510
                total_tipo5_rango2 = hoja_ctas.loc[315:510, ["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO", "T/Cambio"]]

                # Eliminar filas con valores NaN en el rango 2
                total_tipo5_rango2 = total_tipo5_rango2.dropna()

                # Eliminar duplicados en el rango 2 basándonos en "SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO"
                total_tipo5_rango2 = total_tipo5_rango2.drop_duplicates(subset=["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO"])

                return total_tipo5_rango1, total_tipo5_rango2
            else:
                print("No se encontró la hoja 'Ctas' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")


    def insertar_en_standard_tipo1(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 1).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 1}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 1, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard para Tipo 1.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")
                
           
            self.vista.mostrar_proceso_finalizado(1)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para Tipo 1: {e}"
            self.vista.mostrar_error(mensaje_error)

    def insertar_en_standard_tipo2(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 2).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar_tipo2, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 2}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 2, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo estándar para Chilefilms.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")

            self.vista.mostrar_proceso_finalizado(2)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para CHF Internacional: {e}"
            self.vista.mostrar_error(mensaje_error)

    def insertar_en_standard_tipo3(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 3).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar_tipo2, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 3}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 3, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard para Chilefilms.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")

            self.vista.mostrar_proceso_finalizado(2)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para CHF Internacional: {e}"
            self.vista.mostrar_error(mensaje_error)

    def insertar_en_standard_tipo4(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 4).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar_tipo2, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 4}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 4, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo estándar para Chilefilms.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")

            self.vista.mostrar_proceso_finalizado(4)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para CHF Internacional: {e}"
            self.vista.mostrar_error(mensaje_error)
    
    def insertar_en_standard_tipo5(self, total_chilefilms):
        
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 5).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar_tipo2, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 5}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 5, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard para Chilefilms.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")

            self.vista.mostrar_proceso_finalizado(2)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para CHF Internacional: {e}"
            self.vista.mostrar_error(mensaje_error)

    def insertar_en_standard_tipo6(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 6).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar_tipo2, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 6}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 6, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo estándar para Chilefilms.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")

            self.vista.mostrar_proceso_finalizado(2)
        except Exception as e:
            mensaje_error = f"Error al insertar en el archivo estándar para CHF Internacional: {e}"
            self.vista.mostrar_error(mensaje_error)