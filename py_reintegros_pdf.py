import pandas as pd
from pathlib import Path
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
import PyPDF2
import math
import traceback
import json  # <-- agregado para persistencia de rutas
import logging
import sys

def resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para desarrollo y para PyInstaller"""
    try:
        # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# --- CONFIGURACIÓN DE LOGGING ---
def setup_logging():
    """Configura el sistema de logging para el motor"""
    logger = logging.getLogger('reintegros_motor')
    logger.setLevel(logging.INFO)
    
    # Evitar logs duplicados
    if not logger.handlers:
        # Formato de los logs
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Handler para archivo
        file_handler = logging.FileHandler('motor_reintegros.log', encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        
        # Handler para consola (solo durante desarrollo)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.WARNING)  # Solo warnings y errores en consola
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    
    return logger

# Crear el logger global
logger = setup_logging()

# --- CONFIGURACIÓN DE RUTAS (MODIFICADA) ---
ARCHIVO_PLANTILLA_HTML = resource_path("plantilla.html")
RUTA_WKHTMLTOPDF = resource_path("wkhtmltopdf/bin/wkhtmltopdf.exe")  #CAMBIADO
PDF_FONDO = resource_path("fondo_reintegro.pdf")
ARCHIVO_CONFIG = Path(resource_path("config.json"))

# Verificar si wkhtmltopdf existe, si no, buscar en otras ubicaciones
if not os.path.exists(RUTA_WKHTMLTOPDF):
    # Intentar otras rutas comunes
    posibles_rutas = [
        r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
        r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
        os.path.join(os.path.dirname(__file__), "wkhtmltopdf", "bin", "wkhtmltopdf.exe")
    ]
    
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            RUTA_WKHTMLTOPDF = ruta
            break
    else:
        logger.error("No se encontró wkhtmltopdf en ninguna ubicación")
options = {
    'page-size': 'Letter',
    'encoding': "UTF-8",
    'no-outline': None,
    'enable-local-file-access': None,
    'margin-top': '35mm',
    'margin-right': '15mm',
    'margin-bottom': '15mm',
    'margin-left': '15mm',
    'disable-smart-shrinking': None,
    'dpi': 300,
    'print-media-type': None
}

# --- 1b. FUNCIONES DE CONFIGURACIÓN ---
def cargar_config():
    """Cargar configuración desde archivo JSON"""
    if ARCHIVO_CONFIG.exists():
        try:
            with open(ARCHIVO_CONFIG, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"No se pudo leer el archivo de configuración: {e}")
    return {}

def guardar_config(config):
    """Guardar configuración en archivo JSON"""
    try:
        with open(ARCHIVO_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logger.warning(f"No se pudo guardar la configuración: {e}")

# --- 2. FUNCIONES AUXILIARES ---
def obtener_pdf_fondo():
    if Path(PDF_FONDO).exists():
        return PDF_FONDO
    else:
        logger.error(f"No se encontro {PDF_FONDO} en: {os.getcwd()}")
        return None

def superponer_pdfs(pdf_contenido, pdf_fondo, pdf_salida):
    try:
        pdf_writer = PyPDF2.PdfWriter()
        with open(pdf_fondo, 'rb') as pdf_fondo_file, \
             open(pdf_contenido, 'rb') as pdf_contenido_file:
            
            pdf_fondo_reader = PyPDF2.PdfReader(pdf_fondo_file)
            pagina_fondo = pdf_fondo_reader.pages[0]
            pdf_contenido_reader = PyPDF2.PdfReader(pdf_contenido_file)
            
            for pagina_contenido in pdf_contenido_reader.pages:
                pagina_nueva = PyPDF2.PageObject.create_blank_page(
                    width=pagina_fondo.mediabox.width, 
                    height=pagina_fondo.mediabox.height
                )
                pagina_nueva.merge_page(pagina_fondo)
                pagina_nueva.merge_page(pagina_contenido)
                pdf_writer.add_page(pagina_nueva)
            
        with open(pdf_salida, 'wb') as f:
            pdf_writer.write(f)
        return True, f"PDF final guardado en: {pdf_salida}"
    
    except Exception as e:
        logger.error(f"No se pudo superponer PDFs: {e}", exc_info=True)
        return False, f"Error al superponer PDFs: {e}"

def convertir_html_a_pdf(html_string, ruta_pdf_salida):
    try:
        pdfkit.from_string(
            html_string,
            ruta_pdf_salida,
            options=options,
            configuration=pdfkit.configuration(wkhtmltopdf=RUTA_WKHTMLTOPDF)
        )
        return True, "PDF temporal creado."
    except Exception as e:
        logger.error(f"ERROR al crear PDF '{ruta_pdf_salida}': {e}")
        return False, f"Error de pdfkit: {e}"

def truncar_a_2_decimales(valor):
    return math.trunc(float(valor) * 100) / 100.0

# Agregar estas funciones en py_reintegros_pdf.py

def verificar_estructura_anexos(df_anexo_v, df_anexo_vi):
    """
    Verifica que los DataFrames tengan las columnas esperadas
    Retorna: (bool, str) - (éxito, mensaje de error)
    """
    try:
        # Columnas requeridas para Anexo V
        columnas_requeridas_v = [
            'RFC', 'NO_COMPROBANTE', 'CLAVE_PLAZA', 'PERIODO', 
            'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO', 'NOMBRE(S)', 
            'CCT', 'FECHA_INICIO', 'FECHA_TERMINO'
        ]
        
        # Columnas requeridas para Anexo VI
        columnas_requeridas_vi = [
            'NO_COMPROBANTE', 'TIPO_CONCEPTO', 'COD_CONCEPTO', 
            'DESC_CONCEPTO', 'IMPORTE'
        ]
        
        errores = []
        
        # Verificar Anexo V
        columnas_faltantes_v = [col for col in columnas_requeridas_v if col not in df_anexo_v.columns]
        if columnas_faltantes_v:
            errores.append(f"ANEXO V - Columnas faltantes: {', '.join(columnas_faltantes_v)}")
            errores.append(f"Columnas encontradas: {', '.join(df_anexo_v.columns.tolist())}")
        
        # Verificar Anexo VI
        columnas_faltantes_vi = [col for col in columnas_requeridas_vi if col not in df_anexo_vi.columns]
        if columnas_faltantes_vi:
            errores.append(f"ANEXO VI - Columnas faltantes: {', '.join(columnas_faltantes_vi)}")
            errores.append(f"Columnas encontradas: {', '.join(df_anexo_vi.columns.tolist())}")
        
        if errores:
            return False, "Errores en la estructura de los archivos:\n" + "\n".join(errores)
        
        return True, "Estructura de archivos verificada correctamente"
        
    except Exception as e:
        return False, f"Error al verificar estructura: {str(e)}"

def guardar_log_estructura(error_mensaje, rutas_anexo_v, rutas_anexo_vi):
    """Guarda errores de estructura en el log"""
    try:
        log_filename = "estructura_errores.log"
        with open(log_filename, "a", encoding="utf-8") as log_file:
            log_file.write("=" * 80 + "\n")
            log_file.write(f"ERROR DE ESTRUCTURA - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log_file.write("=" * 80 + "\n")
            log_file.write(f"Anexo V: {rutas_anexo_v}\n")
            log_file.write(f"Anexo VI: {rutas_anexo_vi}\n")
            log_file.write(f"Error: {error_mensaje}\n")
            log_file.write("=" * 80 + "\n\n")
        
        return log_filename
    except Exception as e:
        logger.error(f"No se pudo guardar log de estructura: {e}")
        return None


# --- 3. FUNCIÓN: OBTENER PLAZAS POR RFC ---
def obtener_plazas_por_rfc(rfc_input, rutas_anexo_v):
    if isinstance(rutas_anexo_v, str):
        rutas_anexo_v = [rutas_anexo_v]
    
    plazas = []
    try:
        for ruta in rutas_anexo_v:
            logger.info(f"Leyendo Anexo V desde: {ruta}")
            df_anexo_v = pd.read_excel(ruta, dtype=str)
            filtro = df_anexo_v['RFC'].astype(str).str.upper().str.contains(str(rfc_input).upper().strip())

            oficios = df_anexo_v[filtro]
            
            if not oficios.empty:
                for idx, fila in oficios.iterrows():
                    plaza_info = {
                        'RFC': fila.get('RFC', '').strip(),
                        'NO_COMPROBANTE': fila.get('NO_COMPROBANTE', ''),
                        'CLAVE_PLAZA': fila.get('CLAVE_PLAZA', ''),
                        'PERIODO': fila.get('PERIODO', ''),
                        'nombre_completo': f"{fila.get('PRIMER_APELLIDO', '')} {fila.get('SEGUNDO_APELLIDO', '')} {fila.get('NOMBRE(S)', '')}".strip()
                    }
                    plazas.append(plaza_info)
                logger.info(f"Encontradas {len(oficios)} plazas para RFC: {rfc_input}")
    except Exception as e:
        logger.error(f"Error al leer los Anexos V: {e}", exc_info=True)
        return False, f"Error al leer los Anexos V: {e}", []
    
    if not plazas:
        logger.warning(f"No se encontró ningún RFC que coincida con: {rfc_input}")
        return False, f"No se encontró ningún RFC que coincida con: {rfc_input}", []
    
    return True, "Plazas encontradas", plazas

# --- 4. FUNCIÓN PRINCIPAL DEL MOTOR ---
def generar_reintegros_pdf(
        rfc_input, 
        config_reintegro, 
        datos_manuales_input,
        ruta_anexo_v,
        ruta_anexo_vi,
        ruta_carpeta_salida,
        no_comprobantes_seleccionados=None,
        progress_callback=None
    ):
    
    # --- Persistencia: guardar rutas seleccionadas ---
    config = cargar_config()
    config['ultima_ruta_anexo_v'] = str(ruta_anexo_v) if isinstance(ruta_anexo_v, Path) else ruta_anexo_v
    config['ultima_ruta_anexo_vi'] = str(ruta_anexo_vi) if isinstance(ruta_anexo_vi, Path) else ruta_anexo_vi
    config['ultima_carpeta_salida'] = str(ruta_carpeta_salida) if isinstance(ruta_carpeta_salida, Path) else ruta_carpeta_salida
    guardar_config(config)

    # Convertir a listas si son strings
    if isinstance(ruta_anexo_v, str):
        ruta_anexo_v = [ruta_anexo_v]
    if isinstance(ruta_anexo_vi, str):
        ruta_anexo_vi = [ruta_anexo_vi]
    
    logger.info("Iniciando 'py_reintegros' (Modo PDF con capas)...")

    # --- Validar rutas ---
    for ruta in ruta_anexo_v:
        if not Path(ruta).exists():
            logger.error(f"No se encontró el Anexo V en: {ruta}")
            return False, f"Error: No se encontró el Anexo V en:\n{ruta}"
    for ruta in ruta_anexo_vi:
        if not Path(ruta).exists():
            logger.error(f"No se encontró el Anexo VI en: {ruta}")
            return False, f"Error: No se encontró el Anexo VI en:\n{ruta}"
    
    if not Path(ruta_carpeta_salida).exists():
        try:
            os.makedirs(ruta_carpeta_salida)
            logger.info(f"Carpeta creada: {ruta_carpeta_salida}")
        except Exception as e:
            logger.error(f"No se pudo crear la carpeta de salida: {e}")
            return False, f"Error: No se pudo crear la carpeta de salida:\n{e}"

    # --- Cargar Anexos ---
    logger.info("Cargando Anexos (esto puede tardar)...")
    try:
        df_anexo_v_list = [pd.read_excel(ruta, dtype=str) for ruta in ruta_anexo_v]
        df_anexo_v = pd.concat(df_anexo_v_list, ignore_index=True)
        
        df_anexo_vi_list = [pd.read_excel(ruta, dtype=str) for ruta in ruta_anexo_vi]
        df_anexo_vi = pd.concat(df_anexo_vi_list, ignore_index=True)
        
        # VERIFICAR ESTRUCTURA DE ANEXOS
        exito_estructura, mensaje_estructura = verificar_estructura_anexos(df_anexo_v, df_anexo_vi)
        if not exito_estructura:
            # Guardar en log de estructura
            log_file = guardar_log_estructura(mensaje_estructura, ruta_anexo_v, ruta_anexo_vi)
            return False, f"{mensaje_estructura}\n\nSe ha guardado un log en: {log_file}"
        
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'], errors='coerce').fillna(0.0)
        logger.info("Anexos cargados exitosamente")
    except Exception as e:
        logger.error(f"ERROR al leer los archivos Excel: {e}", exc_info=True)
        return False, f"ERROR al leer los archivos Excel: {e}"
    # --- Obtener PDF de fondo ---
    pdf_fondo = obtener_pdf_fondo()
    if not pdf_fondo:
        msg = "ERROR: No se encontró 'fondo_reintegro.pdf'."
        logger.error(msg)
        return False, msg

    # --- Configurar Jinja2 ---con el resource_path
    try:
        # Obtener el directorio donde están los recursos
        base_path = os.path.dirname(resource_path("plantilla.html"))
        env = Environment(loader=FileSystemLoader(base_path))
        template = env.get_template("plantilla.html")  # ← Nombre directo del archivo
    except Exception as e:
        msg = f"ERROR: No se encontró la plantilla 'plantilla.html'. Error: {e}"
        logger.error(msg)
        return False, msg  

    # --- Filtrar Anexo V por RFC ---
    if no_comprobantes_seleccionados:
        filtro = df_anexo_v['NO_COMPROBANTE'].astype(str).str.strip().isin(no_comprobantes_seleccionados)
        oficios_encontrados = df_anexo_v[filtro]
    else:
        filtro = df_anexo_v['RFC'].astype(str).str.strip() == str(rfc_input).strip()
        oficios_encontrados = df_anexo_v[filtro]


    if oficios_encontrados.empty:
        msg = f"--- ERROR: No se encontró ningún registro que coincida con el RFC: {rfc_input} ---"
        logger.error(msg)
        return False, msg

    # --- Filtrar solo los comprobantes seleccionados ---
    if no_comprobantes_seleccionados:
        oficios_encontrados = oficios_encontrados[
            oficios_encontrados['NO_COMPROBANTE'].isin(no_comprobantes_seleccionados)
        ]
        if oficios_encontrados.empty:
            msg = "ERROR: Ninguno de los comprobantes seleccionados fueron encontrados."
            logger.error(msg)
            return False, msg

    logger.info(f"¡Coincidencia encontrada! Se procesarán {len(oficios_encontrados)} oficio(s) para este RFC.")
    
    # Preparar fecha
    ahora = datetime.now()
    dias_semana = {
        0: "lunes", 1: "martes", 2: "miércoles",
        3: "jueves", 4: "viernes", 5: "sábado", 6: "domingo"
    }
    meses_nombre = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
        7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    fecha_hoy_str = f"{dias_semana[ahora.weekday()].capitalize()}, {ahora.day:02d} de {meses_nombre[ahora.month]} de {ahora.year}"

    mensajes_exito = []
    total_oficios = len(oficios_encontrados)

    # --- Iterar sobre CADA oficio encontrado ---
    for idx, (index, fila_v) in enumerate(oficios_encontrados.iterrows(), start=1):
        
        no_comprobante = fila_v.get('NO_COMPROBANTE', 'S/C')
        rfc = fila_v.get('RFC', 'S/RFC')
        logger.info(f"Procesando: {rfc} - Comprobante: {no_comprobante}")
        nombre_completo = f"{fila_v.get('PRIMER_APELLIDO', '')} {fila_v.get('SEGUNDO_APELLIDO', '')} {fila_v.get('NOMBRE(S)', '')}".strip()
        qna = fila_v.get('PERIODO', '')
        detalles = df_anexo_vi[df_anexo_vi['NO_COMPROBANTE'] == no_comprobante]
        percepciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'P']
        deducciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'D']
        total_percepciones = percepciones_df['IMPORTE'].sum()
        total_deducciones = deducciones_df['IMPORTE'].sum()
        total_liquido = truncar_a_2_decimales(total_percepciones - total_deducciones)
        total_a_reintegrar = 0.0
        
        if config_reintegro['tipo'] == 'TOTAL':
            total_a_reintegrar = total_liquido
        else:
            if config_reintegro['modo'] == 'DIAS':
                total_a_reintegrar = truncar_a_2_decimales((total_liquido / 15.0) * float(config_reintegro['dias']))
            else:
                conceptos_str = str(config_reintegro.get('concepto', '')).strip()
                if conceptos_str:
                    conceptos_lista = [c.strip() for c in conceptos_str.split(',') if c.strip()]
                    total_concepto = 0.0
                    
                    for concepto_buscar in conceptos_lista:
                        s = percepciones_df['COD_CONCEPTO'].astype(str).str.strip()
                        concepto_df = percepciones_df[s == concepto_buscar]
                        if concepto_df.empty:
                            concepto_df = percepciones_df[s.str.lstrip('0') == concepto_buscar.lstrip('0')]
                        if not concepto_df.empty:
                            total_concepto += concepto_df['IMPORTE'].astype(float).sum()
                    
                    total_a_reintegrar = truncar_a_2_decimales(
                        (total_concepto / 15.0) * float(config_reintegro['dias']) if config_reintegro.get('por_dias') else total_concepto
                    )
        
        contexto_datos = {
            "nombre_completo": nombre_completo, "rfc": rfc, "no_comprobante": no_comprobante,
            "clave_cobro": fila_v.get('CLAVE_PLAZA', ''), "cct": fila_v.get('CCT', ''),
            "campo_abajo_motivo": datos_manuales_input.get("CAMPO_ABAJO_MOTIVO", ""),
            "desde": fila_v.get('FECHA_INICIO', ''), "hasta": fila_v.get('FECHA_TERMINO', ''),
            "qna": qna, "fecha_hoy": fecha_hoy_str,
            "percepciones": percepciones_df.to_dict('records'),
            "deducciones": deducciones_df.to_dict('records'),
            "total_percepciones": float(truncar_a_2_decimales(total_percepciones)),
            "total_deducciones": float(truncar_a_2_decimales(total_deducciones)),
            "total_liquido": float(total_liquido),
            "total_a_reintegrar": float(total_a_reintegrar),
            "motivo_reintegro": config_reintegro['tipo'],
            "nivel_educativo": datos_manuales_input.get("NIVEL_EDUCATIVO", "")
        }
        html_final_renderizado = template.render(contexto_datos)
        nombre_archivo_temporal = f"{rfc}_{no_comprobante}_temp.pdf"
        ruta_temporal = Path(ruta_carpeta_salida) / nombre_archivo_temporal
        exito_pdf, mensaje_pdf = convertir_html_a_pdf(html_final_renderizado, ruta_temporal)
        if not exito_pdf:
            return False, mensaje_pdf
        nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
        ruta_final = Path(ruta_carpeta_salida) / nombre_archivo_final
        exito_superponer, mensaje_superponer = superponer_pdfs(ruta_temporal, pdf_fondo, ruta_final)
        try: os.remove(ruta_temporal)
        except OSError: pass
        if not exito_superponer:
            return False, mensaje_superponer
        mensajes_exito.append(f"Éxito: {nombre_archivo_final}")

        # Reportar progreso
        if progress_callback:
            progress_callback(idx, total_oficios)

    return True, "\n".join(mensajes_exito)

# --- 5. BLOQUE DE PRUEBA ---
if __name__ == "__main__":
    logger.info("--- MODO DE PRUEBA DEL MOTOR ---")