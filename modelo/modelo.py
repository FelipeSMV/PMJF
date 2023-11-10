import pandas as pd
import os
import shutil

class Modelo:
    def __init__(self):
        self.data = None
        self.planilla_path = "recursos/planilla_standard.xlsx"

    def leer_excel(self, file_path):
        self.data = pd.read_excel(file_path, index_col=0)
        print(self.data)
        

    def procesar_datos(self):
        if self.data is not None and not self.data.empty:
            try:
                c1_value = self.data['Consolidados'].iloc[0]  # Acceder a la primera fila, columna 'C'
                d1_value = self.data['Montos'].iloc[0]  # Acceder a la primera fila, columna 'D'
                total = c1_value + d1_value
                print(f"Resultado de la suma: {total}")
                return total
            except KeyError as e:
                print(f"Error: No se pudo encontrar la celda necesaria en el archivo. Detalles: {e}")
                return 0
        else:
            return 0

    def actualizar_planilla(self, total):
        try:
            planilla = pd.read_excel(self.planilla_path)

            # Actualizar solo la celda E4 con el total
            planilla.at[3, 'E'] = total

            planilla.to_excel(self.planilla_path, index=False)

            # Guardar una copia en el escritorio
            escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
            copia_path = os.path.join(escritorio, "planilla_modificada.xlsx")
            shutil.copy(self.planilla_path, copia_path)
            print(f"Planilla actualizada y copiada en el escritorio: {copia_path}")
        except FileNotFoundError:
            print(f"Error: No se puede encontrar el archivo '{self.planilla_path}'. Verifica la ruta y la existencia del archivo.")