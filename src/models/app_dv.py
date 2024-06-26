from src.models.deudas_vencidas import Deudas_Vencidas
from src.models.rutas import verificar_rutas, seleccionar_archivo, seleccionar_carpeta
from src.utils.resource_path import resource_path
from tkinter import messagebox
from customtkinter import *
import threading
import PIL.Image


class App_DV():
    def abrir_manual(self):
        os.startfile(resource_path("src/doc/MANUAL_APP_DEUDAS_VENCIDAS.docx"))
    
    def centrar_ventana(self, ventana):
        screen_width = ventana.winfo_screenwidth()
        screen_height = ventana.winfo_screenheight()
        ventana.update_idletasks()
        ventana_width = ventana.winfo_reqwidth()
        ventana_height = ventana.winfo_reqheight()
        x = (screen_width - ventana_width) // 2
        y = (screen_height - ventana_height) // 2
        ventana.geometry(f"+{x}+{y}")
    
    def deshabilitar_botones(self):
        self.boton_ejecutar.configure(state="disabled")
        self.boton_deudores.configure(state="disabled")
        self.boton_duda.configure(state="disabled")
        self.boton_config.configure(state="disabled")
        self.combobox_analistas.configure(state="disabled")
        self.entry_morosidad.configure(state="disabled")
        self.checkbox_apoyo.configure(state="disabled")
        self.checkbox_hoja.configure(state="disabled")
        self.checkbox_fichero.configure(state="disabled")
        self.ope_con_mov.configure(state="disabled")
        self.ope_sin_mov.configure(state="disabled")
        self.proc_resolucion.configure(state="disabled")
        self.proc_pre_resolucion.configure(state="disabled")
        self.proc_liquidacion.configure(state="disabled")
        self.liquidado.configure(state="disabled")
    
    def habilitar_botones(self):
        self.boton_ejecutar.configure(state="normal")
        self.boton_deudores.configure(state="normal")
        self.boton_duda.configure(state="normal")
        self.boton_config.configure(state="normal")
        self.combobox_analistas.configure(state="readonly")
        self.entry_morosidad.configure(state="normal")
        self.checkbox_hoja.configure(state="normal")
        self.checkbox_fichero.configure(state="normal")
        self.ope_con_mov.configure(state="normal")
        self.ope_sin_mov.configure(state="normal")
        self.proc_resolucion.configure(state="normal")
        self.proc_pre_resolucion.configure(state="normal")
        if self.var_analista.get() not in ["TODOS", "WALTER LOPEZ", "REGION NORTE", "REGION SUR"]:
            self.checkbox_apoyo.configure(state="normal")
        if not self.var_apoyo.get():
            self.proc_liquidacion.configure(state="normal")
            self.liquidado.configure(state="normal")
    
    def verificar_thread(self, thread):
        if thread.is_alive():
            self.app.after(1000, self.verificar_thread, thread)
        else:
            self.habilitar_botones()
    
    def iniciar_tarea(self, action):
        self.deshabilitar_botones()
        if action == 1:
            thread = threading.Thread(target=self.exportar)
        elif action == 2:
            thread = threading.Thread(target=self.ejecutar)
        else:
            return
        thread.start()
        self.app.after(1000, self.verificar_thread, thread)
    
    def exportar(self):
        self.progressbar.start()
        try:
            rutas = verificar_rutas()
            if rutas == False:
                return
            reporte = Deudas_Vencidas(rutas)
            reporte.exportar_deudores()
        except Exception as ex:
            messagebox.showerror("Error", str(ex))
            return
        finally:
            self.progressbar.stop()
    
    def ejecutar(self):
        self.progressbar.start()
        try:
            rutas = verificar_rutas()
            if rutas == False:
                return
            
            analista = self.combobox_analistas.get()
            formato = self.var_fichero_local.get()
            dias_morosidad = self.entry_morosidad.get()
            
            if dias_morosidad == "":
                messagebox.showerror("Error", "Por favor, ingrese los días de morosidad (>=1).")
                return
            elif not dias_morosidad.isdigit():
                messagebox.showerror("Error", "Por favor, ingrese un número válido (>=1).")
                return
            elif int(dias_morosidad) < 1:
                messagebox.showerror("Error", "Por favor, ingrese un número válido (>=1).")
                return
            
            variables = [
                (self.var_ope_con_mov, "OPERATIVO CON MOVIMIENTO"),
                (self.var_ope_sin_mov, "OPERATIVO SIN MOVIMIENTO"),
                (self.var_proc_liquidacion, "PROCESO DE LIQUIDACIÓN"),
                (self.var_proc_pre_resolucion, "PROCESO DE PRE RESOLUCION"),
                (self.var_proc_resolucion, "PROCESO DE RESOLUCIÓN"),
                (self.var_liquidado, "LIQUIDADO")
            ]
            lista_estados = [var[1] for var in variables if var[0].get()]
            if len(lista_estados) == 0:
                messagebox.showerror("Error", "Por favor, seleccione al menos un estado.")
                return
            
            apoyo = self.var_apoyo.get()
            
            reporte = Deudas_Vencidas(rutas)
            reporte.obtener_deudas_vencidas(analista, formato, dias_morosidad, lista_estados, apoyo)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))
            return
        finally:
            self.progressbar.stop()
    
    def confirmar_configuracion(self):
        verificar_rutas()
        self.ventana_config.destroy()
    
    def ventana_rutas(self):
        self.ventana_config =CTkToplevel(self.app)
        self.ventana_config.title("Rutas")
        self.ventana_config.resizable(False, False)
        self.ventana_config.grab_set()
        self.ventana_config.focus_set()
        
        titulo1 = CTkLabel(self.ventana_config, text="Seleccionar Archivos", font=("Calibri",12,"bold"))
        titulo1.pack(fill="both", expand=True, padx=10, pady=0)
        
        frame_botones1 = CTkFrame(self.ventana_config)
        frame_botones1.pack_propagate("True")
        frame_botones1.pack(fill="both", expand=True, padx=10, pady=0)
        
        file_dacxanalista = CTkButton(
            frame_botones1, text="Nuevo_DACxANALISTA", font=("Calibri",12), text_color="black",
            fg_color="transparent", border_color="black", border_width=2, hover_color="#d11515",
            width=100, corner_radius=25, command=lambda: seleccionar_archivo("DACXANALISTA"))
        file_dacxanalista.pack(fill="both", expand=True, padx=10, pady=10, side="left")
        
        titulo2 = CTkLabel(self.ventana_config, text="Seleccionar Carpetas", font=("Calibri",12,"bold"))
        titulo2.pack(fill="both", expand=True, padx=10, pady=0)
        
        frame_botones2 = CTkFrame(self.ventana_config)
        frame_botones2.pack_propagate("True")
        frame_botones2.pack(fill="both", expand=True, padx=10, pady=(0,10))
        
        folder_dv = CTkButton(
            frame_botones2, text="Deudas Vencidas", font=("Calibri",12), text_color="black", 
            fg_color="transparent", border_color="black", border_width=2, hover_color="#d11515", 
            width=100, corner_radius=25, command=lambda: seleccionar_carpeta("DEUDAS VENCIDAS"))
        folder_dv.pack(fill="both", expand=True, padx=10, pady=10, side="left")
        
        folder_vacaciones = CTkButton(
            frame_botones2, text="Vacaciones", font=("Calibri",12), text_color="black", 
            fg_color="transparent", border_color="black", border_width=2, hover_color="#d11515", 
            width=100, corner_radius=25, command=lambda: seleccionar_carpeta("VACACIONES"))
        folder_vacaciones.pack(fill="both", expand=True, padx=(0,10), pady=10, side="left")
        
        boton_confirmar = CTkButton(
            self.ventana_config, text="Confirmar", font=("Calibri",12), text_color="black",
            fg_color="transparent", border_color="black", border_width=2, hover_color="#d11515",
            width=100, height=10, corner_radius=5, command=self.confirmar_configuracion)
        boton_confirmar.pack(ipady=2, padx=10, pady=(0,10))
        
        self.centrar_ventana(self.ventana_config)
    
    def check_apoyo_state(self, *args):
        selected_analista = self.var_analista.get()
        if selected_analista in ["TODOS", "WALTER LOPEZ", "REGION NORTE", "REGION SUR"]:
            self.var_apoyo.set(False)
            self.checkbox_apoyo.configure(state="disabled")
        else:
            self.checkbox_apoyo.configure(state="normal")
    
    def check_estado_state(self, *args):
        if self.var_apoyo.get():
            self.var_liquidado.set(False)
            self.var_proc_liquidacion.set(False)
            self.proc_liquidacion.configure(state="disabled")
            self.liquidado.configure(state="disabled")
        else:
            self.proc_liquidacion.configure(state="normal")
            self.liquidado.configure(state="normal")
    
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
        main_frame.pack_propagate(1)
        main_frame.pack(fill="both", expand=True)
        
        ############## TITULO ##############
        frame_title = CTkFrame(main_frame)
        frame_title.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        titulo = CTkLabel(frame_title, text="DEUDAS VENCIDAS", font=("Arial",20,"bold"))
        titulo.pack(padx=(30,0), fill="both", expand=True, anchor="center", side="left")
        
        image = PIL.Image.open(resource_path("src/images/duda.ico"))
        image_duda = CTkImage(image, size=(15, 15))
        self.boton_duda = CTkButton(
            frame_title, image=image_duda, text="", corner_radius=25, border_color="#d11515",
            fg_color="transparent", hover_color="#d11515", width=50, command=self.abrir_manual)
        self.boton_duda.pack(padx=5, pady=5, ipadx=0, ipady=5, anchor="center", side="left")
        
        image = PIL.Image.open(resource_path("src/images/config.ico"))
        image_config = CTkImage(image, size=(15, 15))
        self.boton_config = CTkButton(
            frame_title, image=image_config, text="", corner_radius=25, border_color="#d11515",
            fg_color="transparent", hover_color="#d11515", width=50, command=self.ventana_rutas)
        self.boton_config.pack(padx=(0,5), pady=5, ipadx=0, ipady=5, anchor="center", side="left")
        
        ############## SELECCIONAR FORMATO SAP ##############
        frame_checkbox = CTkFrame(main_frame)
        frame_checkbox.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        self.boton_deudores = CTkButton(
            frame_checkbox, text="Exportar Deudores", font=("Calibri",15), text_color="black", 
            fg_color="transparent", border_color="#d11515", border_width=3, hover_color="#d11515", 
            width=25, corner_radius=25, command=lambda: self.iniciar_tarea(1))
        self.boton_deudores.pack(padx=(20,30), pady=10, fill="both", expand=True, anchor="center", side="left")
        
        self.var_hoja_calculo = BooleanVar()
        self.var_hoja_calculo.set(False)
        self.var_hoja_calculo.trace("w", lambda *args: self.var_fichero_local.set(not self.var_hoja_calculo.get()))
        self.checkbox_hoja = CTkRadioButton(
            frame_checkbox, text="Hoja", font=("Calibri",15), border_color="black", 
            fg_color="#d11515", hover_color="#d11515", variable=self.var_hoja_calculo)
        self.checkbox_hoja.pack(padx=(10,0), pady=10, fill="y", anchor="center", side="left")
        
        self.var_fichero_local = BooleanVar()
        self.var_fichero_local.set(True)
        self.var_fichero_local.trace("w", lambda *args: self.var_hoja_calculo.set(not self.var_fichero_local.get()))
        self.checkbox_fichero = CTkRadioButton(
            frame_checkbox, text="Fichero", font=("Calibri",15), border_color="black", 
            fg_color="#d11515", hover_color="#d11515", variable=self.var_fichero_local)
        self.checkbox_fichero.pack(padx=(0,0), pady=10, fill="y", anchor="center", side="left")
        
        ############## SELECCION ANALISTA ##############
        frame_analista = CTkFrame(main_frame)
        frame_analista.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        analistas = [
            "TODOS", "RAQUEL CAYETANO", "DIEGO RODRIGUEZ", "JOSE LUIS VALVERDE", "JUAN CARLOS HUATAY", 
            "YOLANDA OLIVA", "WALTER LOPEZ", "REGION NORTE", "REGION SUR"
            ]
        label_analista = CTkLabel(frame_analista, text="Analista Actual: ", font=("Calibri",15,"bold"))
        label_analista.pack(padx=(20, 0), pady=10, fill="both", expand=True, anchor="e", side="left")
        
        self.var_analista = StringVar(value="TODOS")
        self.combobox_analistas = CTkComboBox(
            frame_analista, values=analistas, font=("Calibri",15), border_color="#d11515", width=200, state="readonly")
        self.combobox_analistas.pack(padx=(0, 20), pady=10, fill="y", expand=True, anchor="w", side="left")
        self.combobox_analistas.set("TODOS")
        self.combobox_analistas.configure(variable=self.var_analista, state="readonly")
        
        ############## DIAS DE MOROSIDAD ##############
        frame_morosidad = CTkFrame(main_frame)
        frame_morosidad.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        label_morosidad = CTkLabel(frame_morosidad, text="A partir de", font=("Calibri",15,"bold"))
        label_morosidad.pack(padx=(20, 0), pady=10, fill="y", anchor="center", side="left")
        
        self.entry_morosidad = CTkEntry(frame_morosidad, font=("Calibri",15), width=50, border_color="#d11515")
        self.entry_morosidad.pack(padx=5, pady=10, fill="y", anchor="center", side="left")
        self.entry_morosidad.configure(justify="center")
        self.entry_morosidad.insert(0, "1")
        
        label_morosidad2 = CTkLabel(frame_morosidad, text="días de morosidad", font=("Calibri",15,"bold"))
        label_morosidad2.pack(padx=(0, 10), pady=10, fill="y", anchor="center", side="left")
        
        self.var_apoyo = BooleanVar()
        self.var_apoyo.set(False)
        self.var_apoyo.trace("w", lambda *args: self.var_fichero_local.set(not self.var_hoja_calculo.get()))
        self.checkbox_apoyo = CTkCheckBox(
            frame_morosidad, text="APOYOS", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_apoyo)
        self.checkbox_apoyo.pack(padx=10, pady=10, anchor="center", side="right")
        self.checkbox_apoyo.configure(state="disabled")
        self.var_analista.trace_add("write", self.check_apoyo_state)
        
        ############## SELECCIONAR ESTADOS ##############
        frame_estado = CTkFrame(main_frame)
        frame_estado.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        label_estado_dac = CTkLabel(frame_estado, text="Seleccionar Estados: ", font=("Calibri",15,"bold"))
        label_estado_dac.grid(row=0, column=0, columnspan=2, padx=(20, 20), pady=(5, 0), sticky="nsew")
        
        self.var_ope_con_mov = BooleanVar()
        self.var_ope_con_mov.set(True)
        self.ope_con_mov = CTkCheckBox(
            frame_estado, text="OP. CON MOVIMIENTO", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_ope_con_mov)
        self.ope_con_mov.grid(row=1, column=0, padx=(20, 10), pady=(10, 0), sticky="nsew")
        
        self.var_ope_sin_mov = BooleanVar()
        self.var_ope_sin_mov.set(True)
        self.ope_sin_mov = CTkCheckBox(
            frame_estado, text="OP. SIN MOVIMIENTO", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_ope_sin_mov)
        self.ope_sin_mov.grid(row=1, column=1, padx=(10, 20), pady=(10, 0), sticky="nsew")
        
        self.var_proc_resolucion = BooleanVar()
        self.var_proc_resolucion.set(False)
        self.proc_resolucion = CTkCheckBox(
            frame_estado, text="PROC. RESOLUCION", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_proc_resolucion)
        self.proc_resolucion.grid(row=2, column=0, padx=(20, 10), pady=(10, 0), sticky="nsew")
        
        self.var_proc_pre_resolucion = BooleanVar()
        self.var_proc_pre_resolucion.set(False)
        self.proc_pre_resolucion = CTkCheckBox(
            frame_estado, text="PROC. PRE RESOLUCION", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_proc_pre_resolucion)
        self.proc_pre_resolucion.grid(row=2, column=1, padx=(10, 20), pady=(10, 0), sticky="nsew")
        
        self.var_proc_liquidacion = BooleanVar()
        self.var_proc_liquidacion.set(False)
        self.proc_liquidacion = CTkCheckBox(
            frame_estado, text="PROC. LIQUIDACION", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_proc_liquidacion)
        self.proc_liquidacion.grid(row=3, column=0, padx=(20, 10), pady=10, sticky="nsew")
        self.var_apoyo.trace_add("write", self.check_estado_state)
        
        self.var_liquidado = BooleanVar()
        self.var_liquidado.set(False)
        self.liquidado = CTkCheckBox(
            frame_estado, text="LIQUIDADO", font=("Calibri",15), border_color="#d11515", 
            border_width=2, fg_color="#d11515", hover_color="#d11515", variable=self.var_liquidado)
        self.liquidado.grid(row=3, column=1, padx=(10, 20), pady=10, sticky="nsew")
        self.var_apoyo.trace_add("write", self.check_estado_state)
        
        ############## EJECUTAR ##############
        self.boton_ejecutar = CTkButton(
            main_frame, text="EJECUTAR", text_color="black", font=("Calibri",20,"bold"), fg_color="gray", 
            hover_color="#d11515", border_color="black", border_width=3, corner_radius=25, height=50, 
            command=lambda: self.iniciar_tarea(2))
        self.boton_ejecutar.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        ############## PROGRESSBAR ##############
        self.progressbar = CTkProgressBar(
            main_frame, mode="indeterminate", orientation="horizontal", progress_color="#d11515", height=10, border_width=0)
        self.progressbar.pack(padx=10, pady=(10, 0), fill="both", expand=True, anchor="center", side="top")
        
        ############## © ##############
        label_copyrigth = CTkLabel(main_frame, text="© Creado por Mirko Iriarte (C26823)", font=("Calibri",11), text_color="black")
        label_copyrigth.pack(padx=10, pady=0, fill="both", expand=True, anchor="center", side="bottom")
        
        self.centrar_ventana(self.app)
        self.app.mainloop()