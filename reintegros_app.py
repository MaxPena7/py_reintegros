import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading

try:
    import py_reintegros_pdf as motor
except ImportError:
    messagebox.showerror("Error de Importación", 
                         "No se pudo encontrar el archivo 'py_reintegros_pdf.py'.\n"
                         "Asegúrate de que esté en la misma carpeta que 'reintegros_app.py'.")
    exit()

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Subjefatura de Nóminas - Generador de Reintegros PDF")
        self.iconbitmap(default="reintegro_icono.ico")
        self.geometry("700x1000")
        self.grid_columnconfigure(0, weight=1)

        # Frame scrolleable principal
        self.main_scrollable = ctk.CTkScrollableFrame(self)
        self.main_scrollable.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.main_scrollable.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Cambiar icono de la ventana e icono de barra de tareas
        try:
            self.iconbitmap("reintegro_icono.ico")
        except:
            pass

        self.ruta_anexo_v = ctk.StringVar()
        self.ruta_anexo_vi = ctk.StringVar()
        self.ruta_carpeta_salida = ctk.StringVar()
        self.tipo_reintegro_var = ctk.StringVar(value="TOTAL")
        self.modo_parcial_var = ctk.StringVar(value="DIAS")
        self.modo_concepto_var = ctk.StringVar(value="TOTAL")
        self.plazas_disponibles = []
        self.plazas_seleccionadas = {}
        self.rutas_anexo_v_lista = []
        self.rutas_anexo_vi_lista = []
        self.cargando = False
        self.spinner_frame = None

        # --- 1. SELECCIÓN DE ARCHIVOS ---
        self.frame_archivos = ctk.CTkFrame(self.main_scrollable)
        self.frame_archivos.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.frame_archivos.grid_columnconfigure(1, weight=1)
        
        ctk.CTkButton(self.frame_archivos, text="Anexo V", command=self.seleccionar_anexo_v).grid(row=0, column=0, padx=10, pady=5)
        self.lbl_anexo_v = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_anexo_v, fg_color="white", text_color="black", corner_radius=5)
        self.lbl_anexo_v.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkButton(self.frame_archivos, text="Anexo VI", command=self.seleccionar_anexo_vi).grid(row=1, column=0, padx=10, pady=5)
        self.lbl_anexo_vi = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_anexo_vi, fg_color="white", text_color="black", corner_radius=5)
        self.lbl_anexo_vi.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkButton(self.frame_archivos, text="Carpeta Salida", command=self.seleccionar_salida).grid(row=2, column=0, padx=10, pady=5)
        self.lbl_carpeta_salida = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_carpeta_salida, fg_color="white", text_color="black", corner_radius=5)
        self.lbl_carpeta_salida.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # --- 2. DATOS DEL EMPLEADO ---
        self.frame_datos = ctk.CTkFrame(self.main_scrollable)
        self.frame_datos.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.frame_datos.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.frame_datos, text="RFC:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_rfc = ctk.CTkEntry(self.frame_datos, placeholder_text="Ingrese el RFC a buscar")
        self.entry_rfc.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        self.btn_consultar = ctk.CTkButton(self.frame_datos, text="Consultar Plazas", 
                                           command=self.consultar_plazas, width=100)
        self.btn_consultar.grid(row=0, column=2, padx=10, pady=10)
        
        ctk.CTkLabel(self.frame_datos, text="Nivel Educativo:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_nivel = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej: PREESCOLAR")
        self.entry_nivel.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(self.frame_datos, text="Motivo (Línea 2):").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_motivo_2 = ctk.CTkTextbox(self.frame_datos, height=80)
        self.entry_motivo_2.grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        # --- 3. SELECCIÓN DE PLAZAS ---
        self.frame_plazas = ctk.CTkFrame(self.main_scrollable)
        self.frame_plazas.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.frame_plazas.grid_columnconfigure(0, weight=1)
        self.frame_plazas.grid_remove()
        
        ctk.CTkLabel(self.frame_plazas, text="Seleccionar Plazas:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        
        self.btn_seleccionar_todas = ctk.CTkButton(self.frame_plazas, text="Seleccionar Todas", 
                                                   command=self.seleccionar_todas_plazas, width=120)
        self.btn_seleccionar_todas.grid(row=0, column=1, padx=5, pady=10, sticky="e")
        
        self.btn_deseleccionar_todas = ctk.CTkButton(self.frame_plazas, text="Deseleccionar Todas", 
                                                     command=self.deseleccionar_todas_plazas, width=130)
        self.btn_deseleccionar_todas.grid(row=0, column=2, padx=5, pady=10, sticky="e")
        
        self.scrollable_plazas = ctk.CTkScrollableFrame(self.frame_plazas, height=150)
        self.scrollable_plazas.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.scrollable_plazas.grid_columnconfigure(0, weight=1)

        # --- 4. LÓGICA DE REINTEGRO ---
        self.frame_logica = ctk.CTkFrame(self.main_scrollable)
        self.frame_logica.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        self.frame_logica.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.frame_logica, text="Tipo de Reintegro:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.frame_tipo_reintegro = ctk.CTkFrame(self.frame_logica, fg_color="transparent")
        self.frame_tipo_reintegro.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkRadioButton(self.frame_tipo_reintegro, text="Total", variable=self.tipo_reintegro_var, value="TOTAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        ctk.CTkRadioButton(self.frame_tipo_reintegro, text="Parcial", variable=self.tipo_reintegro_var, value="PARCIAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        
        self.frame_parcial = ctk.CTkFrame(self.frame_logica)
        self.frame_parcial.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.frame_parcial.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.frame_parcial, text="Modo Parcial:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.frame_modo_parcial = ctk.CTkFrame(self.frame_parcial, fg_color="transparent")
        self.frame_modo_parcial.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkRadioButton(self.frame_modo_parcial, text="Por Días", variable=self.modo_parcial_var, value="DIAS", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        ctk.CTkRadioButton(self.frame_modo_parcial, text="Por Concepto", variable=self.modo_parcial_var, value="CONCEPTO", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        
        self.frame_parcial_dias = ctk.CTkFrame(self.frame_parcial, fg_color="transparent")
        self.frame_parcial_dias.grid(row=1, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(self.frame_parcial_dias, text="Días a reintegrar:").pack(side="left", padx=10, pady=5)
        self.entry_parcial_dias = ctk.CTkEntry(self.frame_parcial_dias, width=60, placeholder_text="1-15")
        self.entry_parcial_dias.pack(side="left", padx=5, pady=5)
        
        self.frame_parcial_concepto = ctk.CTkFrame(self.frame_parcial, fg_color="transparent")
        self.frame_parcial_concepto.grid(row=2, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(self.frame_parcial_concepto, text="Concepto(s):").pack(side="left", padx=10, pady=5)
        self.entry_parcial_concepto = ctk.CTkEntry(self.frame_parcial_concepto, width=100, placeholder_text="Ej: 07,09,15")
        self.entry_parcial_concepto.pack(side="left", padx=5, pady=5)
        
        self.frame_modo_concepto = ctk.CTkFrame(self.frame_parcial_concepto, fg_color="transparent")
        self.frame_modo_concepto.pack(side="left", padx=10, pady=5)
        ctk.CTkRadioButton(self.frame_modo_concepto, text="Total (Concepto)", variable=self.modo_concepto_var, value="TOTAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        ctk.CTkRadioButton(self.frame_modo_concepto, text="Por Días", variable=self.modo_concepto_var, value="DIAS", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        
        self.frame_concepto_dias = ctk.CTkFrame(self.frame_parcial_concepto, fg_color="transparent")
        self.frame_concepto_dias.pack(side="left", padx=10, pady=5)
        self.entry_concepto_dias = ctk.CTkEntry(self.frame_concepto_dias, width=60, placeholder_text="1-15")
        self.entry_concepto_dias.pack(side="left", padx=5, pady=5)

        # --- 5. ACCIONES Y ESTADO ---
        self.frame_acciones = ctk.CTkFrame(self.main_scrollable)
        self.frame_acciones.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
        self.frame_acciones.grid_columnconfigure(0, weight=1)
        
        self.btn_generar = ctk.CTkButton(self.frame_acciones, text="Generar Reintegro(s)", 
                                         font=ctk.CTkFont(weight="bold", size=14), 
                                         command=self.iniciar_generacion)
        self.btn_generar.grid(row=0, column=0, padx=10, pady=10, sticky="ew", ipady=10)
        
        self.progress_bar = ctk.CTkProgressBar(self.main_scrollable, mode="determinate")
        self.progress_bar.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        self.progress_bar.set(0)
        self.progress_bar.grid_remove()
        
        self.lbl_status = ctk.CTkLabel(self.main_scrollable, text="Listo.", height=40, text_color="gray", wraplength=680)
        self.lbl_status.grid(row=6, column=0, padx=10, pady=10, sticky="ew")

        self.actualizar_visibilidad()

    def seleccionar_anexo_v(self):
        rutas = filedialog.askopenfilenames(title="Seleccionar Anexo V (puedes seleccionar múltiples)", filetypes=[("Excel", "*.xlsx")])
        if rutas:
            self.rutas_anexo_v_lista = list(rutas)
            if len(rutas) == 1:
                self.ruta_anexo_v.set(rutas[0])
            else:
                self.ruta_anexo_v.set(f"{len(rutas)} archivo(s) seleccionado(s)")

    def seleccionar_anexo_vi(self):
        rutas = filedialog.askopenfilenames(title="Seleccionar Anexo VI (puedes seleccionar múltiples)", filetypes=[("Excel", "*.xlsx")])
        if rutas:
            self.rutas_anexo_vi_lista = list(rutas)
            if len(rutas) == 1:
                self.ruta_anexo_vi.set(rutas[0])
            else:
                self.ruta_anexo_vi.set(f"{len(rutas)} archivo(s) seleccionado(s)")

    def seleccionar_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta de Salida")
        if ruta: self.ruta_carpeta_salida.set(ruta)

    def consultar_plazas(self):
        rfc = self.entry_rfc.get().strip()
        
        if not rfc:
            messagebox.showerror("Error", "Ingrese un RFC para consultar.")
            return
        
        if not self.rutas_anexo_v_lista:
            messagebox.showerror("Error", "Seleccione el Anexo V primero.")
            return
        
        self.cargando = True
        self.mostrar_spinner()
        self.btn_consultar.configure(state="disabled")
        threading.Thread(target=self._consultar_plazas_thread, args=(rfc, self.rutas_anexo_v_lista), daemon=True).start()

    def _consultar_plazas_thread(self, rfc, anexo_v):
        exito, mensaje, plazas = motor.obtener_plazas_por_rfc(rfc, anexo_v)
        
        def _update():
            self.ocultar_spinner()
            self.btn_consultar.configure(state="normal")
            if exito:
                self.plazas_disponibles = plazas
                self.mostrar_plazas(plazas)
                messagebox.showinfo("Éxito", f"Se encontraron {len(plazas)} plaza(s).")
            else:
                messagebox.showerror("Error", mensaje)
                self.frame_plazas.grid_remove()
        
        self.after(0, _update)

    def mostrar_plazas(self, plazas):
        for widget in self.scrollable_plazas.winfo_children():
            widget.destroy()
        
        self.plazas_seleccionadas = {}
        
        for plaza in plazas:
            var = ctk.BooleanVar(value=False)
            self.plazas_seleccionadas[plaza['NO_COMPROBANTE']] = var
            
            texto_plaza = f"Comprobante: {plaza['NO_COMPROBANTE']} | Clave: {plaza['CLAVE_PLAZA']} | Período: {plaza['PERIODO']}"
            
            frame_plaza = ctk.CTkFrame(self.scrollable_plazas, fg_color="transparent")
            frame_plaza.pack(fill="x", padx=10, pady=5)
            
            checkbox = ctk.CTkCheckBox(frame_plaza, text=texto_plaza, variable=var)
            checkbox.pack(side="left", fill="x", expand=True)
        
        self.frame_plazas.grid()

    def seleccionar_todas_plazas(self):
        for var in self.plazas_seleccionadas.values():
            var.set(True)

    def deseleccionar_todas_plazas(self):
        for var in self.plazas_seleccionadas.values():
            var.set(False)

    def actualizar_visibilidad(self):
        if self.tipo_reintegro_var.get() == "PARCIAL":
            self.frame_parcial.grid()
            if self.modo_parcial_var.get() == "DIAS":
                self.frame_parcial_dias.grid(row=1, column=0, columnspan=2, sticky="ew")
                self.frame_parcial_concepto.grid_remove()
            else:
                self.frame_parcial_dias.grid_remove()
                self.frame_parcial_concepto.grid(row=2, column=0, columnspan=2, sticky="ew")
                if self.modo_concepto_var.get() == "DIAS":
                    self.frame_concepto_dias.pack(side="left", padx=10, pady=5)
                else:
                    self.frame_concepto_dias.pack_forget()
        else:
            self.frame_parcial.grid_remove()

    def iniciar_generacion(self):
        plazas_marcadas = [comprobante for comprobante, var in self.plazas_seleccionadas.items() if var.get()]
        
        if not plazas_marcadas:
            messagebox.showerror("Error", "Debe seleccionar al menos una plaza para reintegrar.")
            return
        
        self.btn_generar.configure(text="Procesando...", state="disabled")
        self.lbl_status.configure(text="Iniciando...", text_color="gray")
        self.progress_bar.set(0)
        self.progress_bar.grid()
        
        threading.Thread(target=self.generar_reintegro, args=(plazas_marcadas,), daemon=True).start()

    def generar_reintegro(self, plazas_marcadas):
        rfc = self.entry_rfc.get()
        salida = self.ruta_carpeta_salida.get()

        if not rfc or not self.rutas_anexo_v_lista or not self.rutas_anexo_vi_lista or not salida:
            self.actualizar_status("Error: Todos los campos son obligatorios.", "red")
            return

        datos_manuales = {
            "NIVEL_EDUCATIVO": self.entry_nivel.get(),
            "CAMPO_ABAJO_MOTIVO": self.entry_motivo_2.get("1.0", "end-1c")
        }

        config_reintegro = {}
        tipo = self.tipo_reintegro_var.get()
        config_reintegro['tipo'] = tipo
        if tipo == 'TOTAL':
            config_reintegro.update({'modo': None, 'dias': 15, 'concepto': None, 'por_dias': False})
        else:
            modo_parcial = self.modo_parcial_var.get()
            config_reintegro['modo'] = modo_parcial
            if modo_parcial == 'DIAS':
                try:
                    dias = int(self.entry_parcial_dias.get())
                    if not (1 <= dias <= 15): raise ValueError
                except ValueError:
                    self.actualizar_status("Error: 'Días a reintegrar' debe ser un número entre 1 y 15.", "red")
                    return
                config_reintegro.update({'dias': dias, 'concepto': None, 'por_dias': False})
            else:
                conceptos = self.entry_parcial_concepto.get()
                if not conceptos:
                    self.actualizar_status("Error: Debe ingresar al menos un código de Concepto.", "red")
                    return
                modo_concepto = self.modo_concepto_var.get()
                if modo_concepto == 'DIAS':
                    try:
                        dias = int(self.entry_concepto_dias.get())
                        if not (1 <= dias <= 15): raise ValueError
                    except ValueError:
                        self.actualizar_status("Error: 'Días (Concepto)' debe ser un número entre 1 y 15.", "red")
                        return
                    config_reintegro.update({'dias': dias, 'concepto': conceptos, 'por_dias': True})
                else:
                    config_reintegro.update({'dias': 15, 'concepto': conceptos, 'por_dias': False})

        self.actualizar_status("Cargando Anexos y procesando... esto puede tardar un momento.", "gray")
        
        try:
            exito, mensaje = motor.generar_reintegros_pdf(
                rfc_input=rfc,
                config_reintegro=config_reintegro,
                datos_manuales_input=datos_manuales,
                ruta_anexo_v=self.rutas_anexo_v_lista,
                ruta_anexo_vi=self.rutas_anexo_vi_lista,
                ruta_carpeta_salida=salida,
                no_comprobantes_seleccionados=plazas_marcadas,
                progress_callback=self.actualizar_progreso
            )
            
            if exito:
                self.actualizar_status(f"¡Éxito!\n{mensaje}", "green")
            else:
                self.actualizar_status(f"Error:\n{mensaje}", "red")

        except Exception as e:
            self.actualizar_status(f"Error Inesperado:\n{e}", "red")

    def actualizar_progreso(self, actual, total):
        def _update():
            progreso = actual / total
            self.progress_bar.set(progreso)
            self.lbl_status.configure(
                text=f"Procesando oficio {actual} de {total}... ({int(progreso * 100)}%)",
                text_color="gray"
            )
        self.after(0, _update)

    def actualizar_status(self, mensaje, color):
        def _update():
            self.lbl_status.configure(text=mensaje, text_color=color)
            self.btn_generar.configure(text="Generar Reintegro(s)", state="normal")
            if color == "red":
                self.progress_bar.grid_remove()
                self.progress_bar.set(0)
        
        self.after(0, _update)

    def mostrar_spinner(self):
        """Muestra un indicador de carga"""
        if self.spinner_frame is None:
            self.spinner_frame = ctk.CTkFrame(self.frame_datos, fg_color="transparent")
            self.spinner_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
            
            self.spinner_label = ctk.CTkLabel(
                self.spinner_frame, 
                text="⏳ Consultando plazas...",
                font=ctk.CTkFont(size=12),
                text_color="gray"
            )
            self.spinner_label.pack()
            
            self.animar_spinner()

    def animar_spinner(self):
        """Anima el indicador de carga"""
        if self.cargando and self.spinner_label:
            spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            self.spinner_index = getattr(self, 'spinner_index', 0)
            self.spinner_label.configure(text=f"{spinner_chars[self.spinner_index]} Consultando plazas...")
            self.spinner_index = (self.spinner_index + 1) % len(spinner_chars)
            self.after(100, self.animar_spinner)

    def ocultar_spinner(self):
        """Oculta el indicador de carga"""
        self.cargando = False
        if self.spinner_frame:
            self.spinner_frame.destroy()
            self.spinner_frame = None
            self.spinner_label = None


if __name__ == "__main__":
    if not os.path.exists(motor.PDF_FONDO):
        messagebox.showerror("Error Crítico", 
                             f"No se encontró el archivo '{motor.PDF_FONDO}'.\n"
                             "La aplicación no puede iniciarse sin el PDF de fondo.")
    elif not os.path.exists(motor.RUTA_WKHTMLTOPDF):
        messagebox.showerror("Error Crítico", 
                             f"No se encontró 'wkhtmltopdf.exe' en:\n{motor.RUTA_WKHTMLTOPDF}\n"
                             "La aplicación no puede generar PDFs sin este programa.")
    else:
        app = App()
        app.mainloop()