import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import*
from PIL import Image, ImageTk
<<<<<<< HEAD
from customtkinter import CTkFrame, CTkEntry, CTkLabel, CTkButton, CTkCheckBox


class Vista:
    
    def __init__(self, controlador):
        self.controlador = controlador
        self.ventana = tk.Tk()
        self.ventana.title("Proyecto MJF")
        

=======
import tkinter.messagebox as messagebox

class Vista:
    def __init__(self):
        self.controlador = None
        self.ventana = tk.Tk()
        self.ventana.title("Proyecto MJF")
>>>>>>> 9d033156159cb1293c14c3746ec6165c4772bae1
        self.ventana.geometry("800x600")
        
        c_fondo = "#F6EEEE"
        self.ventana.config(bg = c_fondo)
        
        logo = PhotoImage(file = "recursos/ciisaL.png")
        
        self.ventana.call("wm", "iconphoto", self.ventana._w, logo)


        titulo = ttk.Label(self.ventana, text="Bienvenido al Proyecto MJF", font=("Helvetica", 20))
        titulo.pack(pady=10)

        ruta_imagen = 'recursos/contabilidad.png'
        imagen = self.cargar_imagen(ruta_imagen, 400, 200)
        imagen_label = ttk.Label(self.ventana, image=imagen)
        imagen_label.image = imagen
        imagen_label.pack(pady=10)

<<<<<<< HEAD
        
        ancho_botones = 25 
        alto = 2
    
        tk.Button(self.ventana, text="Chilefilms", command=self.subir_archivo_tipo1, width=ancho_botones, height=alto, bg="#008F39", fg="#FFF").pack(pady=5)
        tk.Button(self.ventana, text="CHF internacional", command=self.subir_archivo_tipo2, width=ancho_botones, height=alto, bg="#008F39", fg="#FFF").pack(pady=5)
        tk.Button(self.ventana, text="Cine color", command=self.subir_archivo_tipo3, width=ancho_botones, height=alto, bg="#008F39", fg="#FFF").pack(pady=5)
        tk.Button(self.ventana, text="Matriz Chilefilms", command=self.subir_archivo_tipo4, width=ancho_botones, height=alto, bg="#008F39", fg="#FFF").pack(pady=5)
        
        
=======
        ancho_botones = 20  
        ttk.Button(self.ventana, text="Chilefilms", command=self.subir_archivo_tipo1, width=ancho_botones).pack(pady=5)
        ttk.Button(self.ventana, text="CHF internacional", command=self.subir_archivo_tipo2, width=ancho_botones).pack(pady=5)
        ttk.Button(self.ventana, text="Cine color", command=self.subir_archivo_tipo3, width=ancho_botones).pack(pady=5)
        ttk.Button(self.ventana, text="Matriz Chilefilms", command=self.subir_archivo_tipo4, width=ancho_botones).pack(pady=5)
        ttk.Button(self.ventana, text="Nuevo Tipo", command=self.subir_archivo_tipo5, width=ancho_botones).pack(pady=5)
>>>>>>> 9d033156159cb1293c14c3746ec6165c4772bae1

    def iniciar(self):
        self.ventana.mainloop()

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

    def cargar_imagen(self, ruta, width, height):
        imagen_pil = Image.open(ruta)
        imagen_responsive = imagen_pil.resize((width, height), resample=Image.BICUBIC)
        imagen_tk = ImageTk.PhotoImage(imagen_responsive)
        
        return imagen_tk
<<<<<<< HEAD
 
    
 
  
    

=======

    def establecer_controlador(self, controlador):
        self.controlador = controlador

    def mostrar_mensaje(self, mensaje):
        messagebox.showinfo("Información", mensaje)
    
    def mostrar_proceso_finalizado(self, tipo):
        messagebox.showinfo("Información", f"Proceso para Tipo {tipo} finalizado correctamente.")
    
    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)
>>>>>>> 9d033156159cb1293c14c3746ec6165c4772bae1
