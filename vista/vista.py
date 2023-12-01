import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import*
from PIL import Image, ImageTk      
import tkinter.messagebox as messagebox
from customtkinter import CTk, CTkButton, CTkFrame

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

        # Este es el frame para poder modificar los botones
        frame_primer_paso = tk.Frame(self.pestaña_primer_paso)
        frame_primer_paso.pack(pady=10)

        # Crear y organizar los botones en columnas y filas
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
        frame_segundo_paso = tk.Frame(self.pestaña_segundo_paso)
        frame_segundo_paso.pack(pady=10)
        boton_ajuste_segundo_paso = CTkButton(frame_segundo_paso, text="Subir Ajuste", command=self.subir_ajuste_segundo_paso,
                                            fg_color="#008F39", corner_radius=12, border_width=2, height=50, width=180)
        boton_ajuste_segundo_paso.grid(row=0, column=0, padx=10, pady=10)
        
        def eliminar_ajustes():
            try:
                self.controlador.eliminar_archivos_ajustes()
                self.mostrar_mensaje("Archivos de ajustes eliminados correctamente.")
            except Exception as e:
                mensaje_error = f"Error al eliminar archivos de ajustes: {e}"
                self.mostrar_error(mensaje_error)

        boton_eliminar_ajustes = CTkButton(frame_segundo_paso, text="Eliminar Ajustes", command=eliminar_ajustes,
                                           fg_color="red", corner_radius=12, border_width=2, height=50, width=180)
        boton_eliminar_ajustes.grid(row=0, column=1, padx=10, pady=10)

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

    def subir_ajuste_segundo_paso(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        if ruta_archivo:
            # Llama al método correspondiente del controlador
            self.controlador.subir_ajuste(ruta_archivo)
   

    def establecer_controlador(self, controlador):
        self.controlador = controlador

    def mostrar_mensaje(self, mensaje):
        messagebox.showinfo("Información", mensaje)
    
    def mostrar_proceso_finalizado(self, tipo):
        messagebox.showinfo("Información", f"Proceso finalizado correctamente.")
    
    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)