import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading

# Importar el motor que ya hicimos
try:
    import py_reintegros_pdf as motor
except ImportError:
    messagebox.showerror("Error de Importación", 
                         "No se pudo encontrar el archivo 'py_reintegros_pdf.py'.\n"
                         "Asegúrate de que esté en la misma carpeta que 'reintegros_app.py'.")
    exit()

# --- Configuración de la Apariencia ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuración de la Ventana Principal ---
        self.title("Generador de Reintegros")
        self.geometry("700x780")
        self.grid_columnconfigure(1, weight=1)

        # --- Variables (sin cambios) ---
        self.ruta_anexo_v = ctk.StringVar()
        self.ruta_anexo_vi = ctk.StringVar()
        self.ruta_carpeta_salida = ctk.StringVar()
        self.tipo_reintegro_var = ctk.StringVar(value="TOTAL")
        self.modo_parcial_var = ctk.StringVar(value="DIAS")
        self.modo_concepto_var = ctk.StringVar(value="TOTAL")

        # --- 1. SELECCIÓN DE ARCHIVOS (Sin cambios) ---
        self.frame_archivos = ctk.CTkFrame(self)
        self.frame_archivos.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        self.frame_archivos.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(self.frame_archivos, text="Anexo V", command=self.seleccionar_anexo_v).grid(row=0, column=0, padx=10, pady=5)
        self.lbl_anexo_v = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_anexo_v, fg_color="gray20", corner_radius=5)
        self.lbl_anexo_v.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self.frame_archivos, text="Anexo VI", command=self.seleccionar_anexo_vi).grid(row=1, column=0, padx=10, pady=5)
        self.lbl_anexo_vi = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_anexo_vi, fg_color="gray20", corner_radius=5)
        self.lbl_anexo_vi.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self.frame_archivos, text="Carpeta Salida", command=self.seleccionar_salida).grid(row=2, column=0, padx=10, pady=5)
        self.lbl_carpeta_salida = ctk.CTkLabel(self.frame_archivos, textvariable=self.ruta_carpeta_salida, fg_color="gray20", corner_radius=5)
        self.lbl_carpeta_salida.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # --- 2. DATOS DEL EMPLEADO (Sin cambios) ---
        self.frame_datos = ctk.CTkFrame(self)
        self.frame_datos.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        self.frame_datos.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.frame_datos, text="RFC:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_rfc = ctk.CTkEntry(self.frame_datos, placeholder_text="Ingrese el RFC a buscar")
        self.entry_rfc.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkLabel(self.frame_datos, text="Nivel Educativo:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_nivel = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej: PREESCOLAR")
        self.entry_nivel.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(self.frame_datos, text="Motivo (Línea 2):").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_motivo_2 = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej: NO CORRESPONDE...")
        self.entry_motivo_2.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # --- 3. LÓGICA DE REINTEGRO (Sin cambios) ---
        self.frame_logica = ctk.CTkFrame(self)
        self.frame_logica.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
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
        ctk.CTkLabel(self.frame_parcial_concepto, text="Concepto:").pack(side="left", padx=10, pady=5)
        self.entry_parcial_concepto = ctk.CTkEntry(self.frame_parcial_concepto, width=60, placeholder_text="Ej: 07")
        self.entry_parcial_concepto.pack(side="left", padx=5, pady=5)
        self.frame_modo_concepto = ctk.CTkFrame(self.frame_parcial_concepto, fg_color="transparent")
        self.frame_modo_concepto.pack(side="left", padx=10, pady=5)
        ctk.CTkRadioButton(self.frame_modo_concepto, text="Total (Concepto)", variable=self.modo_concepto_var, value="TOTAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        ctk.CTkRadioButton(self.frame_modo_concepto, text="Por Días", variable=self.modo_concepto_var, value="DIAS", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        self.frame_concepto_dias = ctk.CTkFrame(self.frame_parcial_concepto, fg_color="transparent")
        self.frame_concepto_dias.pack(side="left", padx=10, pady=5)
        self.entry_concepto_dias = ctk.CTkEntry(self.frame_concepto_dias, width=60, placeholder_text="1-15")
        self.entry_concepto_dias.pack(side="left", padx=5, pady=5)

        # --- 4. ACCIONES Y ESTADO ---
        self.frame_acciones = ctk.CTkFrame(self)
        self.frame_acciones.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        self.frame_acciones.grid_columnconfigure(0, weight=1)
        
        self.btn_generar = ctk.CTkButton(self.frame_acciones, text="Generar Reintegro(s)", 
                                         font=ctk.CTkFont(weight="bold", size=14), 
                                         command=self.iniciar_generacion)
        self.btn_generar.grid(row=0, column=0, padx=10, pady=10, sticky="ew", ipady=10)
        
        # <-- MODIFICADO: Barra en modo determinado
        self.progress_bar = ctk.CTkProgressBar(self, mode="determinate")
        self.progress_bar.grid(row=4, column=0, columnspan=2, padx=20, pady=5, sticky="ew")
        self.progress_bar.set(0)  # Inicializar en 0
        self.progress_bar.grid_remove()
        
        self.lbl_status = ctk.CTkLabel(self, text="Listo.", height=40, text_color="gray", wraplength=680)
        self.lbl_status.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # --- Estado inicial ---
        self.actualizar_visibilidad()

    # --- 5. FUNCIONES DE LA INTERFAZ ---

    def seleccionar_anexo_v(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Anexo V", filetypes=[("Excel", "*.xlsx")])
        if ruta: self.ruta_anexo_v.set(ruta)

    def seleccionar_anexo_vi(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Anexo VI", filetypes=[("Excel", "*.xlsx")])
        if ruta: self.ruta_anexo_vi.set(ruta)

    def seleccionar_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta de Salida")
        if ruta: self.ruta_carpeta_salida.set(ruta)

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
                    self.frame_concepto_dias.pack_remove()
        else:
            self.frame_parcial.grid_remove()

    def iniciar_generacion(self):
        # <-- MODIFICADO: Preparar la barra
        self.btn_generar.configure(text="Procesando...", state="disabled")
        self.lbl_status.configure(text="Iniciando...", text_color="gray")
        self.progress_bar.set(0)
        self.progress_bar.grid()
        
        threading.Thread(target=self.generar_reintegro, daemon=True).start()

    def generar_reintegro(self):
        # --- 1. Validar Entradas (sin cambios) ---
        rfc = self.entry_rfc.get()
        anexo_v = self.ruta_anexo_v.get()
        anexo_vi = self.ruta_anexo_vi.get()
        salida = self.ruta_carpeta_salida.get()

        if not all([rfc, anexo_v, anexo_vi, salida]):
            self.actualizar_status("Error: Todos los campos de RFC y rutas son obligatorios.", "red")
            return

        # --- 2. Construir Diccionario 'datos_manuales_input' (sin cambios) ---
        datos_manuales = {
            "NIVEL_EDUCATIVO": self.entry_nivel.get(),
            "CAMPO_ABAJO_MOTIVO": self.entry_motivo_2.get()
        }

        # --- 3. Construir Diccionario 'config_reintegro' (sin cambios) ---
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
                concepto = self.entry_parcial_concepto.get()
                if not concepto:
                    self.actualizar_status("Error: Debe ingresar un código de Concepto.", "red")
                    return
                modo_concepto = self.modo_concepto_var.get()
                if modo_concepto == 'DIAS':
                    try:
                        dias = int(self.entry_concepto_dias.get())
                        if not (1 <= dias <= 15): raise ValueError
                    except ValueError:
                        self.actualizar_status("Error: 'Días (Concepto)' debe ser un número entre 1 y 15.", "red")
                        return
                    config_reintegro.update({'dias': dias, 'concepto': concepto, 'por_dias': True})
                else:
                    config_reintegro.update({'dias': 15, 'concepto': concepto, 'por_dias': False})

        # --- 4. Llamar al motor ---
        self.actualizar_status("Cargando Anexos y procesando... esto puede tardar un momento.", "gray")
        
        try:
            # <-- MODIFICADO: Pasar el callback
            exito, mensaje = motor.generar_reintegros_pdf(
                rfc_input=rfc,
                config_reintegro=config_reintegro,
                datos_manuales_input=datos_manuales,
                ruta_anexo_v=anexo_v,
                ruta_anexo_vi=anexo_vi,
                ruta_carpeta_salida=salida,
                progress_callback=self.actualizar_progreso  # <-- AGREGADO
            )
            
            if exito:
                self.actualizar_status(f"¡Éxito!\n{mensaje}", "green")
            else:
                self.actualizar_status(f"Error:\n{mensaje}", "red")

        except Exception as e:
            self.actualizar_status(f"Error Inesperado:\n{e}", "red")

    # <-- AGREGADO: Función para actualizar la barra de progreso
    def actualizar_progreso(self, actual, total):
        """Callback que recibe el progreso desde el motor"""
        def _update():
            progreso = actual / total
            self.progress_bar.set(progreso)
            self.lbl_status.configure(
                text=f"Procesando oficio {actual} de {total}... ({int(progreso * 100)}%)",
                text_color="gray"
            )
        self.after(0, _update)

    def actualizar_status(self, mensaje, color):
        """Función segura para actualizar la UI desde el hilo"""
        def _update():
            self.lbl_status.configure(text=mensaje, text_color=color)
            self.btn_generar.configure(text="Generar Reintegro(s)", state="normal")
            # <-- MODIFICADO: No ocultar ni resetear la barra si terminó exitosamente
            if color == "red":  # Solo ocultar si hubo error
                self.progress_bar.grid_remove()
                self.progress_bar.set(0)
        
        self.after(0, _update)


# --- 6. EJECUTAR LA APLICACIÓN ---
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