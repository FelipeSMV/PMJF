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
        self.notebook.add(self.pestaña_segundo_paso, text="Ajustes financieros")
        self.configurar_pestaña_segundo_paso()

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
        # Configurar el ancho y alto de los botones
        c_verde = "#008F39"

        # Poder crear y organizar los botones en el contenedor
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

        self.lbl_fila = CTkLabel(master=frame_segundo_paso, text="Seleccionar estado de situación financiera:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_fila.pack(pady=5)

        c_lista = "#F2EFFB"
        nombres_filas = {6: "Efectivo y Equivalentes al Efectivo", 7: "Otros activos financieros corrientes", 8: "Otros Activos No Financieros, Corriente", 9: "Deudores comerciales y otras cuentas por cobrar corrientes", 10: "Cuentas por Cobrar a Entidades Relacionadas, Corriente", 11: "Inventarios", 12: "Activos biológicos corrientes", 13: "Activos por impuestos corrientes", 15: "Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para la venta", 16: "Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para distribuir a los propietarios", 20: "Otros activos financieros no corrientes", 21: "Otros activos no financieros no corrientes", 22: "Derechos por cobrar no corrientes", 23: "Cuentas por Cobrar a Entidades Relacionadas, No Corriente", 24: "Inversiones contabilizadas utilizando el método de la participación", 25: "Activos intangibles distintos de la plusvalía", 26: "Plusvalía", 27: "Propiedades, Planta y Equipo", 28: "Activos biológicos, no corrientes", 29: "Propiedad de inversión", 30: "Activos por impuestos diferidos", 37: "Otros pasivos financieros corrientes", 38: "Cuentas por pagar comerciales y otras cuentas por pagar", 39: "Cuentas por Pagar a Entidades Relacionadas, Corriente", 40: "Otras provisiones a corto plazo", 41: "Pasivos por Impuestos corrientes", 42: "Provisiones corrientes por beneficios a los empleados", 43: "Otros pasivos no financieros corrientes", 45: "Total de pasivos corrientes distintos de los pasivos incluidos en grupos de activos para su disposición clasificados como mantenidos para la venta", 48: "Pasivos incluidos en grupos de activos para su disposición clasificados como mantenidos para la venta", 49: "Otros pasivos financieros no corrientes", 50: "Pasivos no corrientes", 51: "Cuentas por Pagar a Entidades Relacionadas, no corriente", 52: "Otras provisiones a largo plazo", 53: "Pasivo por impuestos diferidos", 54: "Provisiones no corrientes por beneficios a los empleados", 58: "Otros pasivos no financieros no corrientes"}
        self.lista_fila = CTkComboBox(master=frame_segundo_paso, values=list(nombres_filas.values()), height=30, width=400, fg_color=c_lista)
        self.lista_fila.set(list(nombres_filas.values())[0])
        self.lista_fila.pack(pady=10)


        self.lbl_columna = CTkLabel(master=frame_segundo_paso, text="Seleccionar Empresa:", font=("Helvetica", 24), text_color="#000000")
        self.lbl_columna.pack(pady=5)


        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}
        self.lista_columna = CTkComboBox(master=frame_segundo_paso, values=list(nombres_columnas.values()), height=30, width=400, fg_color=c_lista)
        self.lista_columna.set(list(nombres_columnas.values())[0])
        self.lista_columna.pack(pady=10)

        self.lbl_valor = CTkLabel(master=frame_segundo_paso, text="Ingrese el valor:", font=("Helvetica", 20), text_color="#000000")
        self.lbl_valor.pack(pady=5)

        self.entry_valor = CTkEntry(master=frame_segundo_paso, placeholder_text="Valor:", width=150)
        self.entry_valor.pack(pady=5)

        c_verde = "#008F39"
        self.btn_insertar = CTkButton(master=frame_segundo_paso, text="Insertar", command=self.insertar_valor, fg_color=c_verde, corner_radius=12, border_width=2, height=30, width=150)
        self.btn_insertar.pack(pady=10)

        #listas
    def insertar_valor(self):
        nombres_filas = {6: "Efectivo y Equivalentes al Efectivo", 7: "Otros activos financieros corrientes", 8: "Otros Activos No Financieros, Corriente", 9: "Deudores comerciales y otras cuentas por cobrar corrientes", 10: "Cuentas por Cobrar a Entidades Relacionadas, Corriente", 11: "Inventarios", 12: "Activos biológicos corrientes", 13: "Activos por impuestos corrientes", 15: "Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para la venta", 16: "Activos no corrientes o grupos de activos para su disposición clasificados como mantenidos para distribuir a los propietarios", 20: "Otros activos financieros no corrientes", 21: "Otros activos no financieros no corrientes", 22: "Derechos por cobrar no corrientes", 23: "Cuentas por Cobrar a Entidades Relacionadas, No Corriente", 24: "Inversiones contabilizadas utilizando el método de la participación", 25: "Activos intangibles distintos de la plusvalía", 26: "Plusvalía", 27: "Propiedades, Planta y Equipo", 28: "Activos biológicos, no corrientes", 29: "Propiedad de inversión", 30: "Activos por impuestos diferidos", 37: "Otros pasivos financieros corrientes", 38: "Cuentas por pagar comerciales y otras cuentas por pagar", 39: "Cuentas por Pagar a Entidades Relacionadas, Corriente", 40: "Otras provisiones a corto plazo", 41: "Pasivos por Impuestos corrientes", 42: "Provisiones corrientes por beneficios a los empleados", 43: "Otros pasivos no financieros corrientes", 45: "Total de pasivos corrientes distintos de los pasivos incluidos en grupos de activos para su disposición clasificados como mantenidos para la venta", 48: "Pasivos incluidos en grupos de activos para su disposición clasificados como mantenidos para la venta", 49: "Otros pasivos financieros no corrientes", 50: "Pasivos no corrientes", 51: "Cuentas por Pagar a Entidades Relacionadas, no corriente", 52: "Otras provisiones a largo plazo", 53: "Pasivo por impuestos diferidos", 54: "Provisiones no corrientes por beneficios a los empleados", 58: "Otros pasivos no financieros no corrientes"}
        
        nombres_columnas = {12: "Chilefilms", 13: "Cce", 14: "Conate II", 15: "CineColor Films", 16: "Sonus", 17: "Servicios Integra", 18: "Serviart", 19: "CHF Inversiones"}

        fila_seleccionada = [k for k, v in nombres_filas.items() if v == self.lista_fila.get()][0]
        columna_seleccionada = [k for k, v in nombres_columnas.items() if v == self.lista_columna.get()][0]
        valor_ingresado = self.entry_valor.get()

        try:
            exito = self.controlador.insertar_valor_en_celda(fila_seleccionada, columna_seleccionada, valor_ingresado)
            if exito:
                mensaje_exito = f"Valor '{valor_ingresado}' insertado correctamente en " \
                                f"{self.lista_fila.get()}, {self.lista_columna.get()}."
                messagebox.showinfo("Éxito", mensaje_exito)
            else:
                messagebox.showerror("Error", "No se pudo realizar la inserción. Verifica tu archivo y las coordenadas.")
        except Exception as e:
            mensaje_error = f"Error al insertar el valor: {e}"
            messagebox.showerror("Error", mensaje_error)


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
        messagebox.showinfo("Información", mensaje)
    
    def mostrar_proceso_finalizado(self, tipo):
        messagebox.showinfo("Información", f"Proceso finalizado correctamente.")
    
    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)
