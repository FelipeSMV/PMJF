import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import*
from PIL import Image, ImageTk      
import tkinter.messagebox as messagebox
from customtkinter import*

class Vista:
    def __init__(self):
        self.controlador = None
        self.ventana = tk.Tk()
        self.ventana.title("Proyecto MJF")
        self.ventana.geometry("800x600")
        c_fondo = "#F6EEEE"
        self.ventana.config(bg=c_fondo)
        logo = PhotoImage(file="recursos/ciisaL.png")
        self.notebook = ttk.Notebook(self.ventana)
        self.notebook.pack(expand=True, fill="both")

        # Pestaña Primer Paso
        self.pestaña_primer_paso = ttk.Frame(self.notebook)
        self.notebook.add(self.pestaña_primer_paso, text="Completar archivo financiero")
        self.configurar_pestaña_primer_paso()

        # Pestaña Segundo Paso
        self.pestaña_segundo_paso = ttk.Frame(self.notebook)
        self.notebook.add(self.pestaña_segundo_paso, text="Ajustes financieros hoja Estado")
        self.configurar_pestaña_segundo_paso()

        # Pestaña Cuarto Paso
        self.pestaña_cuarto_paso = ttk.Frame(self.notebook)
        self.notebook.add(self.pestaña_cuarto_paso, text="Ajustes financieros hoja Resultado")
        self.configurar_pestaña_cuarto_paso()

        # Pestaña Tercer Paso
        self.pestaña_tercer_paso = ttk.Frame(self.notebook)
        self.notebook.add(self.pestaña_tercer_paso, text="Extras")
        self.configurar_pestaña_tercer_paso()

        self.ventana.call("wm", "iconphoto", self.ventana._w, logo)

    def iniciar(self):
        self.ventana.mainloop()

    def configurar_pestaña_primer_paso(self,):
        
        titulo = ttk.Label(self.pestaña_primer_paso, text="Bienvenido al NexoFinanciero", font=("Helvetica", 20))
        titulo.pack(pady=10)
        ruta_imagen = 'recursos/contabilidad.png'
        imagen = self.cargar_imagen(ruta_imagen, 400, 200)
        imagen_label = ttk.Label(self.pestaña_primer_paso, image=imagen)
        imagen_label.image = imagen
        imagen_label.pack(pady=10)
        frame_primer_paso = tk.Frame(self.pestaña_primer_paso)
        frame_primer_paso.pack(pady=10)
        self.crear_botones(frame_primer_paso)

    def crear_botones(self, contenedor):
        c_verde = "#008F39"
        self.boton1 = CTkButton(contenedor, text="Chilefilms Matriz", command=self.subir_archivo_tipo1, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton1.grid(row=0, column=0, padx=10, pady=10)

        self.boton2 = CTkButton(contenedor, text="CCE Individual", command=self.subir_archivo_tipo2, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton2.grid(row=0, column=1, padx=5, pady=5)

        self.boton3 = CTkButton(contenedor, text="Conate II Consolidado", command=self.subir_archivo_tipo3, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton3.grid(row=1, column=0, padx=5, pady=5)

        self.boton4 = CTkButton(contenedor, text="CineColor Films", command=self.subir_archivo_tipo4, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton4.grid(row=1, column=1, padx=5, pady=5)

        self.boton5 = CTkButton(contenedor, text="Sonus Consolidado", command=self.subir_archivo_tipo5, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton5.grid(row=2, column=0, padx=5, pady=5)

        self.boton6 = CTkButton(contenedor, text="Servicios Integra Individual", command=self.subir_archivo_tipo6, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton6.grid(row=2, column=1, padx=5, pady=5)

        self.boton7 = CTkButton(contenedor, text="Serviart Individual", command=self.subir_archivo_tipo7, fg_color=c_verde,
                                corner_radius=12, border_width=2, height=50, width=180)
        self.boton7.grid(row=3, column=0, padx=5, pady=5)

        self.boton8 = CTkButton(contenedor, text="CHF Inversiones Consolidado", command=self.subir_archivo_tipo8,
                                fg_color=c_verde, corner_radius=12, border_width=2, height=50, width=180)
        self.boton8.grid(row=3, column=1, padx=5, pady=5)

    def configurar_pestaña_segundo_paso(self):
        
        frame_segundo_paso = CTkFrame(self.pestaña_segundo_paso, width=320, height=360, fg_color="#F2F2F2")
        frame_segundo_paso.place(relx=0.5, rely=0.5, anchor=CENTER)
        
        #ruta_img2 = 'recursos/excel.png'
        #imagen2 = self.cargar_imagen(ruta_img2, 200, 200)
        #imagen_label = ttk.Label(self.pestaña_segundo_paso, image=imagen2)
        #imagen_label.image = imagen2
        #imagen_label.pack(pady=30)
       

        self.lbl_fila = CTkLabel(master=frame_segundo_paso, text="Seleccionar estado de situación financiera:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_fila.pack(pady=5)

        c_lista = "#F2EFFB"
        nombres_filas = {6: "Efectivo y Equivalentes al Efectivo", 7: "Otros activos financieros corrientes", 8: "Otros Activos No Financieros, Corriente", 9: "Deudores comerciales y otras cuentas por cobrar corrientes", 10: "Cuentas por Cobrar a Entidades Relacionadas, Corriente", 11: "Inventarios", 12: "Activos biológicos corrientes", 13: "Activos por impuestos corrientes", 15: "Activos no corrientes clasificados como mantenidos para la venta.", 16: "Activos no corrientes clasificados como mantenidos para distribuir a los propietarios.", 20: "Otros activos financieros no corrientes", 21: "Otros activos no financieros no corrientes", 22: "Derechos por cobrar no corrientes", 23: "Cuentas por Cobrar a Entidades Relacionadas, No Corriente", 24: "Inversiones contabilizadas utilizando el método de la participación", 25: "Activos intangibles distintos de la plusvalía", 26: "Plusvalía", 27: "Propiedades, Planta y Equipo", 28: "Activos biológicos, no corrientes", 29: "Propiedad de inversión", 30: "Activos por impuestos diferidos", 37: "Otros pasivos financieros corrientes", 38: "Cuentas por pagar comerciales y otras cuentas por pagar", 39: "Cuentas por Pagar a Entidades Relacionadas, Corriente", 40: "Otras provisiones a corto plazo", 41: "Pasivos por Impuestos corrientes", 42: "Provisiones corrientes por beneficios a los empleados", 43: "Otros pasivos no financieros corrientes", 45: "Pasivos incluidos en grupos de activos clasificados como mantenidos para la venta", 48: "Otros pasivos financieros no corrientes", 49: "Pasivos no corrientes", 50: "Cuentas por Pagar a Entidades Relacionadas, no corriente", 51: "Otras provisiones a largo plazo", 52: "Pasivo por impuestos diferidos", 53: "Provisiones no corrientes por beneficios a los empleados", 54: "Otros pasivos no financieros no corrientes"}
        self.lista_fila1 = ttk.Combobox(master=frame_segundo_paso, values=list(nombres_filas.values()), height=10, width=75)
        self.lista_fila1.set(list(nombres_filas.values())[0])
        self.lista_fila1.pack(pady=10)


        self.lbl_columna = CTkLabel(master=frame_segundo_paso, text="Seleccionar Empresa:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_columna.pack(pady=5)


        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}
        self.lista_columna1 = ttk.Combobox(master=frame_segundo_paso, values=list(nombres_columnas.values()), height=10, width=40)
        self.lista_columna1.set(list(nombres_columnas.values())[0])
        self.lista_columna1.pack(pady=10)
        

        self.lbl_valor = CTkLabel(master=frame_segundo_paso, text="Ingrese el valor:", font=("Helvetica", 20), text_color="#000000")
        self.lbl_valor.pack(pady=5)

        self.entry_valor1 = CTkEntry(master=frame_segundo_paso, placeholder_text="Valor:", width=150)
        self.entry_valor1.pack(pady=5)

        c_verde = "#008F39"
        self.btn_insertar = CTkButton(master=frame_segundo_paso, text="Insertar", command=self.insertar_valor, fg_color=c_verde, corner_radius=12, border_width=2, height=30, width=150)
        self.btn_insertar.pack(pady=10)


    def insertar_valor(self):
        nombres_filas = {6:"Efectivo y Equivalentes al Efectivo", 7:"Otros activos financieros corrientes",8:"Otros Activos No Financieros, Corriente",9:"Deudores comerciales y otras cuentas por cobrar corrientes", 10:"Cuentas por Cobrar a Entidades Relacionadas, Corriente", 11:"Inventarios", 12:"Activos biológicos corrientes", 13:"Activos por impuestos corrientes", 15:"Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para la venta", 16:"Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para distribuir a los propietarios",20:"Otros activos financieros no corrientes", 21: "Otros activos no financieros no corrientes", 22: "Derechos por cobrar no corrientes", 23:"Cuentas por Cobrar a Entidades Relacionadas, No Corriente", 24:"Inversiones contabilizadas utilizando el método de la participación", 25:"Activos intangibles distintos de la plusvalía", 26:"Plusvalía", 27: "Propiedades, Planta y Equipo", 28: "Activos biológicos, no corrientes", 29: "Propiedad de inversión", 30:"Activos por impuestos diferidos", 37:"Otros pasivos financieros corrientes", 38:"Cuentas por pagar comerciales y otras cuentas por pagar", 39:"Cuentas por Pagar a Entidades Relacionadas, Corriente", 40:"Otras provisiones a corto plazo", 41:"Pasivos por Impuestos corrientes", 42:"Provisiones corrientes por beneficios a los empleados", 43:"Otros pasivos no financieros corrientes", 45:"Pasivos incluidos en grupos de activos para su disposición clasificados como mantenidos para la venta", 48:"Otros pasivos financieros no corrientes", 49:"Pasivos no corrientes", 50:"Cuentas por Pagar a Entidades Relacionadas, no corriente", 51:"Otras provisiones a largo plazo", 52:"Pasivo por impuestos diferidos", 53:"Provisiones no corrientes por beneficios a los empleados", 54:"Otros pasivos no financieros no corrientes"}
       
        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}

        fila_seleccionada = [k for k, v in nombres_filas.items() if v == self.lista_fila1.get()][0]
        columna_seleccionada = [k for k, v in nombres_columnas.items() if v == self.lista_columna1.get()][0]

        if fila_seleccionada is not None and columna_seleccionada is not None:
            valor = self.entry_valor1.get()
            exito = self.controlador.insertar_valor_en_celda(fila_seleccionada, columna_seleccionada, valor)
            
            if exito:
                self.mostrar_mensaje("Valor insertado correctamente.")
            else:
                self.mostrar_error("Error al insertar el valor.")
        elif self.lista_fila1.get() not in nombres_filas.values():
            self.mostrar_error("Fila no encontrada.")
            print(f"Error al insertar valor en celda - Fila: {fila_seleccionada}")
        else:
            self.mostrar_error("Columna no encontrada.")


    def configurar_pestaña_cuarto_paso(self):

        frame_cuarto_paso = CTkFrame(self.pestaña_cuarto_paso, width=320, height=360, fg_color="#F2F2F2")
        frame_cuarto_paso.place(relx=0.5, rely=0.5, anchor=CENTER)
        self.lbl_fila = CTkLabel(master=frame_cuarto_paso, text="Seleccionar estado de situación financiera:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_fila.pack(pady=5)
        c_lista = "#F2EFFB"
        nombres_filas = {5: "Ingresos de actividaes ordinarias", 7: "Costos de ventas", 11: "Gastos de administracion", 12: "Depreciación y/o Amortización del Ejercicio", 15: "Ingresos financieros", 17:"Costos financieros", 19: "Participacion en las ganancias (perdidas) de asociadas y negocios conjunto", 21:"Otros ingresos", 23:"Otros egresos", 25:"Diferencias de cambio", 27:"Resultado por unidades de reajuste", 31:"Gasto por impuestos a las ganancias"}
        self.lista_fila = ttk.Combobox(master=frame_cuarto_paso, values=list(nombres_filas.values()), height=10, width=75)
        self.lista_fila.set(list(nombres_filas.values())[0])
        self.lista_fila.pack(pady=10)
        
        self.lbl_columna = CTkLabel(master=frame_cuarto_paso, text="Seleccionar Empresa:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_columna.pack(pady=5)

        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}
        self.lista_columna = ttk.Combobox(master=frame_cuarto_paso, values=list(nombres_columnas.values()), height=10, width=40)
        self.lista_columna.set(list(nombres_columnas.values())[0])
        self.lista_columna.pack(pady=10)

        self.lbl_valor = CTkLabel(master=frame_cuarto_paso, text="Ingrese el valor:", font=("Helvetica", 20), text_color="#000000")
        self.lbl_valor.pack(pady=5)

        self.entry_valor = CTkEntry(master=frame_cuarto_paso, placeholder_text="Valor:", width=150)
        self.entry_valor.pack(pady=5)

        c_verde = "#008F39"
        self.btn_insertar = CTkButton(master=frame_cuarto_paso, text="Insertar", command=self.insertar_valorR, fg_color=c_verde, corner_radius=12, border_width=2, height=30, width=150)
        self.btn_insertar.pack(pady=10)

    def insertar_valorR(self):
        nombres_filas = {5: "Ingresos de actividaes ordinarias", 7: "Costos de ventas", 11: "Gastos de administracion", 12: "Depreciación y/o Amortización del Ejercicio", 15: "Ingresos financieros", 17:"Costos financieros", 19: "Participacion en las ganancias (perdidas) de asociadas y negocios conjunto", 21:"Otros ingresos", 23:"Otros egresos", 25:"Diferencias de cambio", 27:"Resultado por unidades de reajuste", 31:"Gasto por impuestos a las ganancias"}
        
        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}

        fila_seleccionada = [k for k, v in nombres_filas.items() if v == self.lista_fila.get()][0]
        columna_seleccionada = [k for k, v in nombres_columnas.items() if v == self.lista_columna.get()][0]
        valor_ingresado = self.entry_valor.get()

        try:
            exito = self.controlador.insertar_valor_en_celda_R(fila_seleccionada, columna_seleccionada, valor_ingresado)
            if exito:
                mensaje_info = {"titulo": "Éxito", "contenido": f"Valor '{valor_ingresado}' insertado correctamente en {self.lista_fila.get()}, {self.lista_columna.get()}."}
                self.mostrar_mensaje(mensaje_info)
            else:
                mensaje_info = {"titulo": "Error", "contenido": "No se pudo realizar la inserción. Verifica tu archivo y las coordenadas."}
                self.mostrar_mensaje(mensaje_info)
        except Exception as e:
            mensaje_info = {"titulo": "Error", "contenido": f"Error al insertar el valor: {e}"}
            self.mostrar_mensaje(mensaje_info)

    def configurar_pestaña_tercer_paso(self):
         
        frame_tercer_paso = CTkFrame(self.pestaña_tercer_paso, width=320, height=360, fg_color="#F2F2F2")
        frame_tercer_paso.place(relx=0.5, rely=0.5, anchor=CENTER)
        
        ruta_extra = 'recursos/extra.png'
        imagen3 = self.cargar_imagen(ruta_extra, 200, 150)
        imagen_label = ttk.Label(self.pestaña_tercer_paso, image=imagen3)
        imagen_label.image = imagen3
        imagen_label.pack(pady=30)    
        self.titulo3 = CTkLabel(master=frame_tercer_paso, text="Seleccione un Extra a utilizar:", font=("Helvetica", 35), text_color="#000000")
        self.titulo3.pack(pady=10)        
        self.btn_limpiar_planilla = CTkButton(master=frame_tercer_paso, text="Limpiar planilla", command=self.limpiar_planilla, fg_color="#FF0000", corner_radius=12, border_width=2, height=50, width=180)
        self.btn_limpiar_planilla.pack(pady=10)
        self.btn_descargar_manual = CTkButton(master=frame_tercer_paso, text="Descargar manual de uso", command=self.descargar_manual, fg_color="#0000FF", corner_radius=12, border_width=2, height=50, width=180)
        self.btn_descargar_manual.pack(pady=10)

    def limpiar_planilla(self):
        self.controlador.limpiar_planilla()

    def descargar_manual(self):
        self.controlador.descargar_manual()

    def cargar_imagen(self, ruta, width, height):
        imagen_pil = Image.open(ruta)
        imagen_responsive = imagen_pil.resize((width, height), resample=Image.BICUBIC)
        imagen_tk = ImageTk.PhotoImage(imagen_responsive)      
        return imagen_tk

    def subir_archivo(self, tipo):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        if ruta_archivo:
            if tipo == "tipo1":
                self.controlador.procesar_archivo_tipo1(ruta_archivo)
            elif tipo == "tipo2":
                self.controlador.procesar_archivo_tipo2(ruta_archivo)
            elif tipo == "tipo3":
                self.controlador.procesar_archivo_tipo3(ruta_archivo)
            elif tipo == "tipo4":
                self.controlador.procesar_archivo_tipo4(ruta_archivo)
            elif tipo == "tipo5":
                self.controlador.procesar_archivo_tipo5(ruta_archivo)
            elif tipo == "tipo6":
                self.controlador.procesar_archivo_tipo6(ruta_archivo)
            elif tipo == "tipo7":
                self.controlador.procesar_archivo_tipo7(ruta_archivo)
            elif tipo == "tipo8":
                self.controlador.procesar_archivo_tipo8(ruta_archivo)

    def subir_archivo_tipo1(self):
        self.subir_archivo(tipo="tipo1")

    def subir_archivo_tipo2(self):
        self.subir_archivo(tipo="tipo2")

    def subir_archivo_tipo3(self):
        self.subir_archivo(tipo="tipo3")

    def subir_archivo_tipo4(self):
        self.subir_archivo(tipo="tipo4")
    
    def subir_archivo_tipo5(self):
        self.subir_archivo(tipo="tipo5")
    
    def subir_archivo_tipo6(self):
        self.subir_archivo(tipo="tipo6")
    
    def subir_archivo_tipo7(self):
        self.subir_archivo(tipo="tipo7")

    def subir_archivo_tipo8(self):
        self.subir_archivo(tipo="tipo8")
   

    def establecer_controlador(self, controlador):
        self.controlador = controlador

    def mostrar_mensaje(self, mensaje):
        messagebox.showinfo("Éxito", mensaje)

    
    def mostrar_proceso_finalizado(self, tipo):
        messagebox.showinfo("Información", f"Los datos de Estado se transfirieron correctamente.")
    
    
    def mostrar_proceso_finalizado2(self, tipo):
        messagebox.showinfo("Información", f"Los datos de Resultado se transfirieron correctamente.")
    
    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)