import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
import json
from pathlib import Path
import traceback
from datetime import datetime
import sys

# Importar el motor
try:
    import py_reintegros_pdf as motor
except ImportError:
    messagebox.showerror("Error de Importación", 
                         "No se pudo encontrar el archivo 'py_reintegros_pdf.py'.\n"
                         "Asegúrate de que esté en la misma carpeta que 'reintegros_app.py'.")
    exit()

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Subjefatura de Nóminas - Generador de Reintegros PDF")
        self.geometry("700x1000")
        self.grid_columnconfigure(0, weight=1)
        
        try:
            self.iconbitmap(resource_path("reintegro_icono.ico"))
        except:
            pass

        # Frame scrolleable principal
        self.main_scrollable = ctk.CTkScrollableFrame(self)
        self.main_scrollable.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.main_scrollable.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Variables ---
        self.ruta_anexo_v = ctk.StringVar()
        self.ruta_anexo_vi = ctk.StringVar()
        self.ruta_carpeta_salida = ctk.StringVar()
        self.tipo_reintegro_var = ctk.StringVar(value="TOTAL")
        self.modo_parcial_var = ctk.StringVar(value="DIAS")
        self.modo_concepto_var = ctk.StringVar(value="TOTAL")
        self.plazas_seleccionadas = {}
        self.rutas_anexo_v_lista = []
        self.rutas_anexo_vi_lista = []
        self.cargando = False
        self.spinner_frame = None
        
        # Variable para el checkbox manual
        self.check_manual_var = ctk.BooleanVar(value=False)

        self.cargar_configuracion_guardada()

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

        self.btn_consultar = ctk.CTkButton(self.frame_datos, text="Consultar Plazas", command=self.consultar_plazas, width=100)
        self.btn_consultar.grid(row=0, column=2, padx=10, pady=10)

        ctk.CTkLabel(self.frame_datos, text="Nivel Educativo:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_nivel = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej: PREESCOLAR")
        self.entry_nivel.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self.frame_datos, text="Motivo:").grid(row=2, column=0, padx=10, pady=5, sticky="e")

        self.entry_motivo_2 = ctk.CTkTextbox(self.frame_datos, height=80)
        self.entry_motivo_2.grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        # --- FIX CORRECTO PARA SCROLL ---
        # Cuando el mouse entra al textbox, desactivamos el scroll de afuera
        self.entry_motivo_2.bind("<Enter>", lambda e: self._activar_scroll_motivo(True))
        # Cuando sale, lo reactivamos
        self.entry_motivo_2.bind("<Leave>", lambda e: self._activar_scroll_motivo(False))

       # --- 3. SELECCIÓN DE PLAZAS ---
        self.f_plazas = ctk.CTkFrame(self.main_scrollable)
        self.f_plazas.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.f_plazas.grid_columnconfigure(0, weight=1)
        self.f_plazas.grid_remove()
        
        # Título y Botones
        f_controles_plazas = ctk.CTkFrame(self.f_plazas, fg_color="transparent")
        f_controles_plazas.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ctk.CTkLabel(f_controles_plazas, text="Seleccionar Plazas:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkButton(f_controles_plazas, text="Todas", command=self.seleccionar_todas_plazas, width=60, height=24).pack(side="right", padx=5)
        ctk.CTkButton(f_controles_plazas, text="Ninguna", command=self.deseleccionar_todas_plazas, width=60, height=24).pack(side="right", padx=5)
        
        # --- ENCABEZADOS DE LA TABLA ---
        self.f_headers = ctk.CTkFrame(self.f_plazas, fg_color="gray80", corner_radius=5, height=30)
        self.f_headers.grid(row=1, column=0, sticky="ew", padx=5, pady=(5,0))
        # Configuramos las columnas del encabezado (mismos pesos que usaremos abajo)
        self.f_headers.grid_columnconfigure(0, weight=0) # Checkbox
        self.f_headers.grid_columnconfigure(1, weight=2) # RFC
        self.f_headers.grid_columnconfigure(2, weight=2) # Comprobante
        self.f_headers.grid_columnconfigure(3, weight=3) # Plaza
        self.f_headers.grid_columnconfigure(4, weight=2) # CCT
        self.f_headers.grid_columnconfigure(5, weight=2) # Periodo

        # Etiquetas de los encabezados
        ctk.CTkLabel(self.f_headers, text="✔", width=30, text_color="black").grid(row=0, column=0, padx=2, pady=2)
        ctk.CTkLabel(self.f_headers, text="RFC", font=("Arial", 11, "bold"), text_color="black").grid(row=0, column=1, sticky="ew", padx=5)
        ctk.CTkLabel(self.f_headers, text="NO. COMP", font=("Arial", 11, "bold"), text_color="black").grid(row=0, column=2, sticky="ew", padx=5)
        ctk.CTkLabel(self.f_headers, text="PLAZA", font=("Arial", 11, "bold"), text_color="black").grid(row=0, column=3, sticky="ew", padx=5)
        ctk.CTkLabel(self.f_headers, text="CCT", font=("Arial", 11, "bold"), text_color="black").grid(row=0, column=4, sticky="ew", padx=5)
        ctk.CTkLabel(self.f_headers, text="PERIODO", font=("Arial", 11, "bold"), text_color="black").grid(row=0, column=5, sticky="ew", padx=5)
        
        # Área Scrollable (Ahora solo contiene los datos)
        self.scroll_plazas = ctk.CTkScrollableFrame(self.f_plazas, height=200)
        self.scroll_plazas.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self.scroll_plazas.grid_columnconfigure(0, weight=0) # Checkbox
        self.scroll_plazas.grid_columnconfigure(1, weight=2) # RFC
        self.scroll_plazas.grid_columnconfigure(2, weight=2) # Comp
        self.scroll_plazas.grid_columnconfigure(3, weight=3) # Plaza
        self.scroll_plazas.grid_columnconfigure(4, weight=2) # CCT
        self.scroll_plazas.grid_columnconfigure(5, weight=2) # Periodo
        
        self.scroll_plazas.bind("<Enter>", lambda e: self._activar_scroll_plazas(True))
        self.scroll_plazas.bind("<Leave>", lambda e: self._activar_scroll_plazas(False))

        # --- 4. LÓGICA DE REINTEGRO (Configuración) ---
        self.frame_logica = ctk.CTkFrame(self.main_scrollable)
        self.frame_logica.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        self.frame_logica.grid_columnconfigure(1, weight=1)
        
        # Título y RadioButtons Principales
        ctk.CTkLabel(self.frame_logica, text="Tipo de Reintegro:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.frame_tipo_reintegro = ctk.CTkFrame(self.frame_logica, fg_color="transparent")
        self.frame_tipo_reintegro.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkRadioButton(self.frame_tipo_reintegro, text="Total", variable=self.tipo_reintegro_var, value="TOTAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        ctk.CTkRadioButton(self.frame_tipo_reintegro, text="Parcial", variable=self.tipo_reintegro_var, value="PARCIAL", command=self.actualizar_visibilidad).pack(side="left", padx=5)
        
        # --- Frame PARCIAL (Opciones dinámicas) ---
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

        # --- SECCIÓN DE MONTO MANUAL ---
        ctk.CTkFrame(self.frame_logica, height=2, fg_color="gray70").grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(15,5))
        
        self.check_monto_manual = ctk.CTkCheckBox(
            self.frame_logica, 
            text="Ingresar Monto Manual (Omitir cálculo automático)", 
            variable=self.check_manual_var, 
            command=self.toggle_manual_entry,
            font=ctk.CTkFont(weight="bold")
        )
        self.check_monto_manual.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        
        self.frame_input_manual = ctk.CTkFrame(self.frame_logica, fg_color="transparent")
        self.frame_input_manual.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(self.frame_input_manual, text="Importe Total a Reintegrar: $").pack(side="left", padx=(25, 5))
        self.entry_monto_manual = ctk.CTkEntry(self.frame_input_manual, placeholder_text="0.00", state="disabled", width=120)
        self.entry_monto_manual.pack(side="left", padx=5)

        # --- 5. ACCIONES Y ESTADO ---
        self.frame_acciones = ctk.CTkFrame(self.main_scrollable)
        self.frame_acciones.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
        self.frame_acciones.grid_columnconfigure(0, weight=1)
        
        self.btn_generar = ctk.CTkButton(self.frame_acciones, text="Generar Reintegro(s)", 
                                         font=ctk.CTkFont(weight="bold", size=14), 
                                         command=self.iniciar_generacion)
        self.btn_generar.grid(row=0, column=0, padx=10, pady=10, sticky="ew", ipady=10)
        
        # --- BARRA DE PROGRESO ---
        self.progress_bar = ctk.CTkProgressBar(self.main_scrollable, mode="indeterminate")
        self.progress_bar.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        self.progress_bar.grid_remove() # Oculta al inicio
        
        self.lbl_status = ctk.CTkLabel(
            self.main_scrollable, 
            text="Listo. Atajos: Enter=Consultar | Ctrl+G=Generar | Ctrl+Q=Salir | Esc=Limpiar", 
            height=40, text_color="gray", wraplength=680,
            font=ctk.CTkFont(size=11)
        )
        self.lbl_status.grid(row=6, column=0, padx=10, pady=10, sticky="ew")

        self.actualizar_visibilidad()
        self.configurar_atajos_teclado()

    # --- FUNCIONES UI MANUAL ---
    def toggle_manual_entry(self):
        """Habilita/Deshabilita el campo manual y las opciones de cálculo"""
        if self.check_manual_var.get():
            self.entry_monto_manual.configure(state="normal")
            self.entry_monto_manual.focus()
        else:
            self.entry_monto_manual.configure(state="disabled")
            self.entry_monto_manual.delete(0, 'end')

    # --- FUNCIONES DE SELECCIÓN DE PLAZAS ---
    def seleccionar_todas_plazas(self):
        for var in self.plazas_seleccionadas.values():
            var.set(True)

    def deseleccionar_todas_plazas(self):
        for var in self.plazas_seleccionadas.values():
            var.set(False)

    def _activar_scroll_plazas(self, active):
        if active:
            self.scroll_plazas.bind_all("<MouseWheel>", self._scroll_plazas)
        else:
            self.scroll_plazas.unbind_all("<MouseWheel>")
            self.main_scrollable.bind_all("<MouseWheel>", self._scroll_principal)
            
    def _activar_scroll_motivo(self, active):
        """Controla el scroll cuando el mouse está sobre el motivo"""
        if active:
            # Si el mouse está sobre el motivo, quitamos el control a la ventana principal
            # El textbox se scrolleará solo por defecto al no tener competencia
            self.main_scrollable.unbind_all("<MouseWheel>")
        else:
            # Si el mouse sale, le devolvemos el control a la ventana principal
            self.main_scrollable.bind_all("<MouseWheel>", self._scroll_principal)

    def _scroll_principal(self, event):
        try:
            self.main_scrollable._parent_canvas.yview_scroll(int(-1 * (event.delta / 10)), "units")
        except Exception: pass

    def _scroll_plazas(self, event):
        try:
            self.scroll_plazas._parent_canvas.yview_scroll(int(-1 * (event.delta / 10)), "units")
            return "break"
        except Exception: pass

    # --- FUNCIONES DE ARCHIVOS ---
    def seleccionar_anexo_v(self):
        rutas = filedialog.askopenfilenames(title="Seleccionar Anexo V", filetypes=[("Excel", "*.xlsx")])
        if rutas:
            self.rutas_anexo_v_lista = list(rutas)
            self.ruta_anexo_v.set(rutas[0] if len(rutas)==1 else f"{len(rutas)} archivos")
            self.actualizar_configuracion_rutas()

    def seleccionar_anexo_vi(self):
        rutas = filedialog.askopenfilenames(title="Seleccionar Anexo VI", filetypes=[("Excel", "*.xlsx")])
        if rutas:
            self.rutas_anexo_vi_lista = list(rutas)
            self.ruta_anexo_vi.set(rutas[0] if len(rutas)==1 else f"{len(rutas)} archivos")
            self.actualizar_configuracion_rutas()

    def seleccionar_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta de Salida")
        if ruta: 
            self.ruta_carpeta_salida.set(ruta)
            self.actualizar_configuracion_rutas()

    def actualizar_configuracion_rutas(self):
        try:
            cfg = motor.cargar_config()
            cfg.update({
                'ultima_ruta_anexo_v': self.rutas_anexo_v_lista,
                'ultima_ruta_anexo_vi': self.rutas_anexo_vi_lista,
                'ultima_carpeta_salida': self.ruta_carpeta_salida.get()
            })
            motor.guardar_config(cfg)
        except: pass

    def cargar_configuracion_guardada(self):
        try:
            cfg = motor.cargar_config()
            if 'ultima_ruta_anexo_v' in cfg:
                v = cfg['ultima_ruta_anexo_v']
                self.rutas_anexo_v_lista = v if isinstance(v, list) else [v]
                if self.rutas_anexo_v_lista: self.ruta_anexo_v.set(self.rutas_anexo_v_lista[0] if len(self.rutas_anexo_v_lista)==1 else f"{len(self.rutas_anexo_v_lista)} archivos")
            if 'ultima_ruta_anexo_vi' in cfg:
                vi = cfg['ultima_ruta_anexo_vi']
                self.rutas_anexo_vi_lista = vi if isinstance(vi, list) else [vi]
                if self.rutas_anexo_vi_lista: self.ruta_anexo_vi.set(self.rutas_anexo_vi_lista[0] if len(self.rutas_anexo_vi_lista)==1 else f"{len(self.rutas_anexo_vi_lista)} archivos")
            if 'ultima_carpeta_salida' in cfg: 
                self.ruta_carpeta_salida.set(cfg['ultima_carpeta_salida'])
        except: pass

    # --- FUNCIONES DE UI GENERAL ---
    def configurar_atajos_teclado(self):
        self.entry_rfc.bind('<Return>', lambda e: self.consultar_plazas())
        self.bind('<Control-g>', lambda e: self.iniciar_generacion())
        self.bind('<Control-G>', lambda e: self.iniciar_generacion())
        self.bind('<Escape>', lambda e: self.limpiar_interface())

    def limpiar_interface(self):
        if not self.cargando and self.btn_generar.cget('state') == 'normal':
            self.entry_rfc.delete(0, 'end')
            self.entry_nivel.delete(0, 'end')
            self.entry_motivo_2.delete('1.0', 'end')
            self.check_manual_var.set(False)
            self.toggle_manual_entry()
            self.frame_plazas.grid_remove()
            self.actualizar_status("Interfaz limpiada.", "gray")

    def consultar_plazas(self):
        rfc = self.entry_rfc.get().strip()
        if not rfc or not self.rutas_anexo_v_lista:
            return messagebox.showerror("Error", "Faltan datos (RFC o Anexo V)")
        self.cargando = True
        self.show_spinner()
        self.btn_consultar.configure(state="disabled")
        threading.Thread(target=self._th_consulta, args=(rfc,), daemon=True).start()

    def _th_consulta(self, rfc):
        ok, msg, plazas = motor.obtener_plazas_por_rfc(rfc, self.rutas_anexo_v_lista)
        self.after(0, lambda: self._post_consulta(ok, msg, plazas))

    def _post_consulta(self, ok, msg, plazas):
        self.hide_spinner()
        self.btn_consultar.configure(state="normal")
        
        if not ok: 
            return messagebox.showerror("Error", msg)
            
        # Limpiar lista anterior
        for w in self.scroll_plazas.winfo_children(): 
            w.destroy()
            
        self.plazas_seleccionadas = {}
        
        # Dibujar filas
        for idx, p in enumerate(plazas):
            var = ctk.BooleanVar(value=False)
            self.plazas_seleccionadas[p['NO_COMPROBANTE']] = var
            
            # 1. Checkbox (CORREGIDO: Se eliminó sticky="c")
            c = ctk.CTkCheckBox(self.scroll_plazas, text="", variable=var, width=24)
            c.grid(row=idx, column=0, padx=2, pady=2) 
            
            # 2. RFC
            ctk.CTkLabel(self.scroll_plazas, text=p['RFC'], anchor="w").grid(row=idx, column=1, padx=5, pady=2, sticky="ew")
            
            # 3. Comprobante
            ctk.CTkLabel(self.scroll_plazas, text=p['NO_COMPROBANTE'], anchor="w").grid(row=idx, column=2, padx=5, pady=2, sticky="ew")
            
            # 4. Plaza
            ctk.CTkLabel(self.scroll_plazas, text=p['CLAVE_PLAZA'], anchor="w").grid(row=idx, column=3, padx=5, pady=2, sticky="ew")
            
            # 5. CCT
            ctk.CTkLabel(self.scroll_plazas, text=p.get('CCT', 'S/D'), anchor="w").grid(row=idx, column=4, padx=5, pady=2, sticky="ew")

            #6. Periodo
            ctk.CTkLabel(self.scroll_plazas, text=p['PERIODO'], anchor="w").grid(row=idx, column=5, padx=5, pady=2, sticky="ew")
       
        self.f_plazas.grid()
        messagebox.showinfo("Éxito", f"{len(plazas)} plazas encontradas")

    def actualizar_visibilidad(self):
        self.toggle_manual_entry() # <--- Asegurarse que esto se llame al actualizar
        if self.tipo_reintegro_var.get() == "PARCIAL":
            self.frame_parcial.grid()
            if self.modo_parcial_var.get() == "DIAS":
                self.frame_parcial_dias.grid(row=1, column=0, columnspan=2, sticky="ew")
                self.frame_parcial_concepto.grid_remove()
            else:
                self.frame_parcial_dias.grid_remove()
                self.frame_parcial_concepto.grid(row=2, column=0, columnspan=2, sticky="ew")
                if self.modo_concepto_var.get() == "DIAS":
                    self.frame_concepto_dias.pack(side="left", padx=10)
                else:
                    self.frame_concepto_dias.pack_forget()
        else:
            self.frame_parcial.grid_remove()

    def iniciar_generacion(self):
        sel = [k for k,v in self.plazas_seleccionadas.items() if v.get()]
        if not sel: return messagebox.showerror("Error", "Seleccione al menos una plaza")
        self.btn_generar.configure(state="disabled", text="Procesando...")
        # --- BARRA DE PROGRESO ACTIVADA ---
        self.progress_bar.grid()
        self.progress_bar.start()
        threading.Thread(target=self.generar_reintegro, args=(sel,), daemon=True).start()

    def generar_reintegro(self, plazas_marcadas):
        rfc = self.entry_rfc.get()
        salida = self.ruta_carpeta_salida.get()
        if not all([rfc, self.rutas_anexo_v_lista, self.rutas_anexo_vi_lista, salida]):
            self.actualizar_status("Faltan campos obligatorios", "red")
            return

        # Lectura Monto Manual
        monto_manual = None
        if self.check_manual_var.get():
            try:
                monto_manual = float(self.entry_monto_manual.get())
                if monto_manual <= 0: raise ValueError
            except:
                self.actualizar_status("Error: Monto manual inválido", "red")
                return

        # Configuración Automática
        conf = {'tipo': self.tipo_reintegro_var.get()}
        if conf['tipo'] == 'TOTAL':
            conf.update({'modo': None, 'dias': 15, 'concepto': None, 'por_dias': False})
        else:
            modo = self.modo_parcial_var.get()
            conf['modo'] = modo
            if modo == 'DIAS':
                try: d = int(self.entry_parcial_dias.get())
                except: d = 0
                conf.update({'dias': d, 'concepto': None, 'por_dias': False})
            else:
                c = self.entry_parcial_concepto.get()
                if not c and not monto_manual:
                    self.actualizar_status("Error: Falta concepto", "red")
                    return
                try: d = int(self.entry_concepto_dias.get())
                except: d = 15
                conf.update({'concepto': c, 'dias': d, 'por_dias': self.modo_concepto_var.get()=='DIAS'})

        datos = {
            "NIVEL_EDUCATIVO": self.entry_nivel.get(),
            "CAMPO_ABAJO_MOTIVO": self.entry_motivo_2.get("1.0", "end-1c")
        }

        self.actualizar_status("Generando PDFs...", "gray")
        try:
            ok, msg = motor.generar_reintegros_pdf(
                rfc, conf, datos,
                self.rutas_anexo_v_lista, self.rutas_anexo_vi_lista, salida,
                no_comprobantes_seleccionados=plazas_marcadas,
                monto_manual_override=monto_manual
            )
            self.after(0, lambda: self.end_gen(ok, msg))
        except Exception as e:
            self.after(0, lambda: self.end_gen(False, str(e)))

    def end_gen(self, ok, msg):
        # --- BARRA DE PROGRESO DESACTIVADA ---
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        
        self.btn_generar.configure(state="normal", text="Generar Reintegro(s)")
        if ok:
            self.lbl_status.configure(text="¡Éxito!", text_color="green")
            messagebox.showinfo("Éxito", msg)
        else:
            self.lbl_status.configure(text="Error", text_color="red")
            messagebox.showerror("Error", msg)

    def actualizar_status(self, msg, col):
        def _upd():
            self.lbl_status.configure(text=msg, text_color=col)
            self.btn_generar.configure(state="normal", text="Generar Reintegro(s)")
            if col == "red": # Si hay error, detener barra
                self.progress_bar.stop()
                self.progress_bar.grid_remove()
        self.after(0, _upd)

    def show_spinner(self):
        self.spinner_frame = ctk.CTkFrame(self.frame_datos, fg_color="transparent")
        self.spinner_frame.grid(row=3, column=0, columnspan=3)
        self.spinner_lbl = ctk.CTkLabel(self.spinner_frame, text="⏳ Cargando...", text_color="gray")
        self.spinner_lbl.pack()
        self.anim()

    def anim(self):
        if self.cargando and self.spinner_lbl:
            chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            self.idx = getattr(self, 'idx', 0)
            self.spinner_lbl.configure(text=f"{chars[self.idx]} Buscando...")
            self.idx = (self.idx + 1) % len(chars)
            self.after(100, self.anim)

    def hide_spinner(self):
        self.cargando = False
        if self.spinner_frame: self.spinner_frame.destroy()

if __name__ == "__main__":
    if not os.path.exists(motor.PDF_FONDO): messagebox.showerror("Falta archivo", "Falta fondo_reintegro.pdf")
    else: App().mainloop()