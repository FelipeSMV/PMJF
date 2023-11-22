import pandas as pd
import shutil
import os
from openpyxl import load_workbook

class Modelo:
    def __init__(self):
        self.archivo_subido = None
        self.archivo_standard = "recursos/planilla_standard.xlsx"
        self.columna_standard_index = 2
        self.columnas_a_insertar = [4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 35, 36,
                                    37, 38, 39, 40, 41, 43, 46, 47, 48, 49, 50, 51, 52, 56, 57, 58, 59, 60, 61, 63]
        self.columnas_a_insertar_tipo3 = [4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 35, 36, 37, 38, 39, 40, 41, 42, 44,47, 48, 49, 50, 51, 52, 53, 57, 58, 59, 60, 61, 62, 64]
        self.filas_a_pegar = [6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 37, 38, 39, 40,
                              41, 42, 43, 45, 48, 49, 50, 51, 52, 53, 54, 58, 59, 60, 61, 62, 63, 65]
        

    def cargar_archivo(self, ruta_archivo):
        try:
            
            self.archivo_subido = pd.ExcelFile(ruta_archivo)
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")


    def buscar_total_tipo1(self):
        if self.archivo_subido is not None:
            if "Estado" in self.archivo_subido.sheet_names and "TOTAL" in self.archivo_subido.parse("Estado").columns:
                return self.archivo_subido.parse("Estado")["TOTAL"].iloc[self.columnas_a_insertar]
            else:
                print("No se encontró la hoja 'Estado' o la columna 'TOTAL' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def buscar_total_tipo2(self):
        if self.archivo_subido is not None:
            if "Estado$" in self.archivo_subido.sheet_names and "TOTAL" in self.archivo_subido.parse("Estado$").columns:
                return self.archivo_subido.parse("Estado$")["TOTAL"].iloc[self.columnas_a_insertar]
            else:
                print("No se encontró la hoja 'Estado$' o la columna 'TOTAL' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def buscar_total_tipo3(self):
        if self.archivo_subido is not None:
            if "Estado" in self.archivo_subido.sheet_names:
                
                hoja_estado = self.archivo_subido.parse("Estado")

                
                if len(hoja_estado.columns) > 2:
                    
                    return hoja_estado.iloc[self.columnas_a_insertar_tipo3, 2]
                else:
                    print("No se encontró la columna en la posición del índice 2 en la hoja 'Estado'.")
            else:
                print("No se encontró la hoja 'Estado' en el archivo subido.")
        else:
            print("No se ha cargado ningún archivo.")

    def buscar_total_tipo4(self):
        if self.archivo_subido is not None:
            if "Ctas" in self.archivo_subido.sheet_names:
                hoja_ctas = self.archivo_subido.parse("Ctas")

                
                print("Columnas disponibles en la hoja 'Ctas':", hoja_ctas.columns)

                
                expected_columns = [
                    'Unnamed: 0',  
                    'SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO',
                    'T/Cambio'
                ]

                
                if all(col in hoja_ctas.columns for col in expected_columns):
                    
                    hoja_ctas_filtrada = hoja_ctas.loc[
                        (hoja_ctas['SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO'].notna()) &
                        (hoja_ctas.index < 320)
                    ]

                    return hoja_ctas_filtrada[expected_columns]
                else:
                    print("No se encontraron todas las columnas esperadas en la hoja 'Ctas'.")
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
        except Exception as e:
            print(f"Error al insertar en el archivo standard para Tipo 1: {e}")

    def insertar_en_standard_tipo2(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 3).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 3}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 3, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard para Tipo 2.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")
        except Exception as e:
            print(f"Error al insertar en el archivo standard para Tipo 2: {e}")

    def insertar_en_standard_tipo3(self, total_chilefilms):
        try:
            archivo_standard = load_workbook(self.archivo_standard)

            if "Estado" in archivo_standard.sheetnames:
                hoja_estado = archivo_standard["Estado"]

                if not hoja_estado.sheet_state == "visible":
                    hoja_estado.sheet_state = "visible"

                columna_standard = [hoja_estado.cell(row=fila, column=self.columna_standard_index + 5).value for fila in
                                    range(1, hoja_estado.max_row + 1)]

                for fila_paste, fila_copy, valor in zip(self.filas_a_pegar, self.columnas_a_insertar, total_chilefilms):
                    print(f"Copiando de fila {fila_copy}, pegando en fila {fila_paste}, columna {self.columna_standard_index + 5}, valor: {valor}")
                    hoja_estado.cell(row=fila_paste, column=self.columna_standard_index + 5, value=valor)

                archivo_standard.save(self.archivo_standard)

                print(f"Datos insertados correctamente en el archivo standard para Tipo 3.")

                escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                nombre_copia = "planilla_standard_modificada.xlsx"
                ruta_copia = os.path.join(escritorio, nombre_copia)

                shutil.copy(self.archivo_standard, ruta_copia)

                print(f"Copia guardada en el escritorio como {nombre_copia}")
            else:
                print("No se encontró la hoja 'Estado' en el archivo estándar.")
        except Exception as e:
            print(f"Error al insertar en el archivo standard para Tipo 3: {e}")

    def insertar_en_planilla_de_prueba(self, total_tipo4):
        try:
            archivo_planilla_prueba = load_workbook("recursos/planilladeprueba.xlsx")

            if "Cta Cte  (Acumuladas Consol)" in archivo_planilla_prueba.sheetnames:
                hoja_cta_cte = archivo_planilla_prueba["Cta Cte  (Acumuladas Consol)"]

                if not hoja_cta_cte.sheet_state == "visible":
                    hoja_cta_cte.sheet_state = "visible"

                
                columna_cta_cte = [hoja_cta_cte.cell(row=fila, column=1).value for fila in
                                range(1, hoja_cta_cte.max_row + 1)]
                columna_t_cambio = [hoja_cta_cte.cell(row=fila, column=2).value for fila in
                                    range(1, hoja_cta_cte.max_row + 1)]

                for index, row in total_tipo4.iterrows():
                    svs_estado = row["SVS ESTADO DE SITUACION FINANCIERA CLASIFICADO"]
                    valor_tipo4 = row["T/Cambio"]

                    
                    if svs_estado in columna_cta_cte:
                        fila_paste = columna_cta_cte.index(svs_estado) + 1

                        
                        if len(columna_t_cambio) > fila_paste:
                            print(f"Copiando de fila {index}, columna 'T/Cambio', valor: {valor_tipo4}")
                            hoja_cta_cte.cell(row=fila_paste, column=2, value=valor_tipo4)
                            print(f"Pegando en fila {fila_paste}, columna 2, valor: {valor_tipo4}")
                        else:
                            print(f"No se encontró la fila correspondiente en la columna 'T/Cambio'. "
                                f"No se realizó el copy-paste para fila {index}.")
                
                archivo_planilla_prueba.save("recursos/planilladeprueba.xlsx")
                print("Datos insertados correctamente en la planilla de prueba.")
        except Exception as e:
            print(f"Error al insertar en la planilla de prueba: {e}")

