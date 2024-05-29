from customtkinter import *
from tkinter import messagebox
from src.database.conexion import ejecutar_query
from src.models.deuda_vencida import obtener_deudas_vencidas
from src.models.seleccionar_archivos import seleccionar_base, seleccionar_dacxanalista
from src.utils.resource_path import resource_path
import threading


class App_DV():
    def deshabilitar_botones(self):
        self.boton_ejecutar.configure(state="disabled")
        self.boton_base.configure(state="disabled")
        self.boton_dacx.configure(state="disabled")
        self.combobox_analistas.configure(state="disabled")
    
    def habilitar_botones(self):
        self.boton_ejecutar.configure(state="normal")
        self.boton_base.configure(state="normal")
        self.boton_dacx.configure(state="normal")
        self.combobox_analistas.configure(state="normal")
    
    def verificar_thread(self, thread):
        if thread.is_alive():
            self.app.after(1000, self.verificar_thread, thread)
        else:
            self.habilitar_botones()
    
    def iniciar_tarea(self, action):
        self.deshabilitar_botones()
        if action == 1:
            thread = threading.Thread(target=self.ejecutar)
        else:
            return
        thread.start()
        self.app.after(1000, self.verificar_thread, thread)
    
    def ejecutar(self):
        self.progressbar.start()
        query = """SELECT * FROM RUTAS WHERE ID == 0"""
        try:
            datos = ejecutar_query(query)
            ruta_base = datos[0][1]
            ruta_dacxa = datos[0][2]
            ruta_resultado = datos[0][3]
            if ruta_base is None or ruta_dacxa is None or ruta_resultado is None:
                messagebox.showerror("Error", "Por favor, configure las rutas de los archivos.")
            elif not os.path.exists(ruta_base):
                messagebox.showerror("Error", "No se encontraró el archivo BASE en la ruta especificada.")
            elif not os.path.exists(ruta_dacxa):
                messagebox.showerror("Error", "No se encontraró el archivo DACxANALISTA en la ruta especificada.")
            else:
                variables = [
                    (self.var_ope_con_mov, "OPERATIVO CON MOVIMIENTO"),
                    (self.var_ope_sin_mov, "OPERATIVO SIN MOVIMIENTO"),
                    (self.var_proc_liquidacion, "PROCESO DE LIQUIDACIÓN"),
                    (self.var_proc_pre_resolucion, "PROCESO DE PRE RESOLUCION"),
                    (self.var_proc_resolucion, "PROCESO DE RESOLUCIÓN"),
                    (self.var_liquidado, "LIQUIDADO")
                ]
                analista = self.combobox_analistas.get()
                obtener_deudas_vencidas(ruta_base, ruta_dacxa, ruta_resultado, variables, analista)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))
        finally:
            self.progressbar.stop()
    
    def crear_app(self):        
        self.app = CTk()
        self.app.title("Deudas Vencidas C&CD")
        icon_path = resource_path("src/images/icono.ico")
        if os.path.isfile(icon_path):
            self.app.iconbitmap(icon_path)
        else:
            messagebox.showwarning("ADVERTENCIA", "No se encontró el archivo 'icono.ico' en la ruta: " + icon_path)
        self.app.resizable(False, False)
        set_appearance_mode("light")
        
        main_frame = CTkFrame(self.app)
        main_frame.pack_propagate(0)
        main_frame.pack(fill="both", expand=True)
        
        frame_title = CTkFrame(main_frame)
        frame_title.grid(row=0, column=0, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        titulo = CTkLabel(frame_title, text="  DEUDAS VENCIDAS  ", font=("Arial",25,"bold"))
        titulo.pack(fill="both", expand=True, ipady=20, anchor="center")
        
        frame_base = CTkFrame(main_frame)
        frame_base.grid(row=1, column=0, padx=(20, 10), pady=(20, 0), sticky="nsew")
        
        ruta_base = CTkLabel(frame_base, text="Ruta BASE", font=("Calibri",17,"bold"))
        ruta_base.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_base = CTkButton(frame_base, text="Seleccionar", font=("Calibri",17), text_color="black",
                                fg_color="transparent", border_color="#d11515", border_width=3, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: seleccionar_base)
        self.boton_base.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        frame_dacx = CTkFrame(main_frame)
        frame_dacx.grid(row=1, column=1, padx=(10, 20), pady=(20, 0), sticky="nsew")
        
        ruta_dacxa = CTkLabel(frame_dacx, text="Ruta DACxAnalista", font=("Calibri",17,"bold"))
        ruta_dacxa.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_dacx = CTkButton(frame_dacx, text="Seleccionar", font=("Calibri",17), text_color="black",
                                fg_color="transparent", border_color="#d11515", border_width=3, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: seleccionar_dacxanalista)
        self.boton_dacx.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        frame_analista = CTkFrame(main_frame)
        frame_analista.grid(row=2, column=0, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        analistas = ["TODOS", "RAQUEL CAYETANO", "YOLANDA OLIVA", "DIEGO RODRIGUEZ", "JOSE LUIS VALVERDE", 
                    "JUAN CARLOS HUATAY", "WALTER LOPEZ", "REGION NORTE", "REGION SUR"]
        label_analista = CTkLabel(frame_analista, text="Analista Actual: ", font=("Calibri",18,"bold"))
        label_analista.pack(padx=(20, 0), pady=(15, 15), fill="both", expand=True, anchor="w", side="left")
        self.combobox_analistas = CTkComboBox(frame_analista, font=("Calibri",17), width=200, values=analistas, 
                                            state="readonly", border_color="#d11515")
        self.combobox_analistas.pack(padx=(0, 40), pady=(15, 15), fill="both", expand=True, anchor="w", side="right")
        self.combobox_analistas.set("TODOS")
        
        frame_estado = CTkFrame(main_frame)
        frame_estado.grid(row=3, column=0, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        label_estado_dac = CTkLabel(frame_estado, text="Seleccionar Estados: ", font=("Calibri",18,"bold"))
        label_estado_dac.grid(row=0, column=0, columnspan=2, padx=(20, 20), pady=(10, 0), sticky="nsew")
        
        self.var_ope_con_mov = BooleanVar()
        self.var_ope_con_mov.set(True)
        ope_con_mov = CTkCheckBox(frame_estado, text="OP. CON MOVIMIENTO", font=("Calibri",17), 
                                    border_color="#d11515", border_width=2, fg_color="#d11515", 
                                    hover_color="#d11515", variable=self.var_ope_con_mov)
        ope_con_mov.grid(row=1, column=0, padx=(20, 10), pady=(10, 0), sticky="nsew")
        
        self.var_ope_sin_mov = BooleanVar()
        self.var_ope_sin_mov.set(True)
        ope_sin_mov = CTkCheckBox(frame_estado, text="OP. SIN MOVIMIENTO", font=("Calibri",17), 
                                    border_color="#d11515", border_width=2, fg_color="#d11515", 
                                    hover_color="#d11515", variable=self.var_ope_sin_mov)
        ope_sin_mov.grid(row=1, column=1, padx=(10, 20), pady=(10, 0), sticky="nsew")
        
        self.var_proc_resolucion = BooleanVar()
        self.var_proc_resolucion.set(False)
        proc_resolucion = CTkCheckBox(frame_estado, text="PROC. RESOLUCION", font=("Calibri",17), 
                                            border_color="#d11515", border_width=2, fg_color="#d11515", 
                                            hover_color="#d11515", variable=self.var_proc_resolucion)
        proc_resolucion.grid(row=2, column=0, padx=(20, 10), pady=(10, 0), sticky="nsew")
        
        self.var_proc_pre_resolucion = BooleanVar()
        self.var_proc_pre_resolucion.set(False)
        proc_pre_resolucion = CTkCheckBox(frame_estado, text="PROC. PRE RESOLUCION", font=("Calibri",17), 
                                        border_color="#d11515", border_width=2, fg_color="#d11515", 
                                        hover_color="#d11515", variable=self.var_proc_pre_resolucion)
        proc_pre_resolucion.grid(row=2, column=1, padx=(10, 20), pady=(10, 0), sticky="nsew")
        
        self.var_proc_liquidacion = BooleanVar()
        self.var_proc_liquidacion.set(False)
        proc_liquidacion = CTkCheckBox(frame_estado, text="PROC. LIQUIDACION", font=("Calibri",17), 
                                        border_color="#d11515", border_width=2, fg_color="#d11515", 
                                        hover_color="#d11515", variable=self.var_proc_liquidacion)
        proc_liquidacion.grid(row=3, column=0, padx=(20, 10), pady=(10, 20), sticky="nsew")
        
        self.var_liquidado = BooleanVar()
        self.var_liquidado.set(False)
        liquidado = CTkCheckBox(frame_estado, text="LIQUIDADO", font=("Calibri",17), border_color="#d11515", 
                                border_width=2, fg_color="#d11515", hover_color="#d11515", 
                                variable=self.var_liquidado)
        liquidado.grid(row=3, column=1, padx=(10, 20), pady=(10, 20), sticky="nsew")
        
        self.boton_ejecutar = CTkButton(main_frame, text="EJECUTAR", text_color="black", font=("Calibri",25,"bold"), 
                                    border_color="black", border_width=3, fg_color="gray", 
                                    hover_color="red", command=lambda: self.iniciar_tarea(1))
        self.boton_ejecutar.grid(row=4, column=0, columnspan=2, ipady=20, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        self.progressbar = CTkProgressBar(main_frame, mode="indeterminate", orientation="horizontal", 
                                        progress_color="#d11515", height=10, border_width=0)
        self.progressbar.grid(row=5, column=0, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
        
        self.app.mainloop()