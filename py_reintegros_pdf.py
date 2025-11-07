import pandas as pd
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
import PyPDF2
import math
import traceback

# --- 1. CONFIGURACIÓN DE RUTAS (CONSTANTES) ---
ARCHIVO_PLANTILLA_HTML = "plantilla.html"
RUTA_WKHTMLTOPDF = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
PDF_FONDO = "fondo_reintegro.pdf"

options = {
    'page-size': 'Letter',
    'encoding': "UTF-8",
    'no-outline': None,
    'enable-local-file-access': None,
    'margin-top': '35mm',
    'margin-right': '15mm',
    'margin-bottom': '15mm',
    'margin-left': '15mm',
}

# --- 2. FUNCIONES AUXILIARES ---

def obtener_pdf_fondo():
    if os.path.exists(PDF_FONDO):
        return PDF_FONDO
    else:
        print(f"[ERROR] No se encontro {PDF_FONDO} en: {os.getcwd()}")
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
        print(f"[ERROR] No se pudo superponer PDFs: {e}")
        traceback.print_exc()
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
        print(f"   -> ERROR al crear PDF '{ruta_pdf_salida}': {e}")
        return False, f"Error de pdfkit: {e}"

def truncar_a_2_decimales(valor):
    return math.trunc(float(valor) * 100) / 100.0

# --- 3. FUNCIÓN: OBTENER PLAZAS POR RFC ---

def obtener_plazas_por_rfc(rfc_input, rutas_anexo_v):
    """
    Obtiene las plazas (NO_COMPROBANTE) disponibles para un RFC específico.
    rutas_anexo_v: puede ser una ruta (str) o una lista de rutas (list)
    Retorna una lista de diccionarios con información de cada plaza.
    """
    # Convertir a lista si es una sola ruta
    if isinstance(rutas_anexo_v, str):
        rutas_anexo_v = [rutas_anexo_v]
    
    plazas = []
    
    try:
        for ruta in rutas_anexo_v:
            df_anexo_v = pd.read_excel(ruta, dtype=str)
            
            # Filtrar por RFC
            filtro = (df_anexo_v['RFC'].astype(str).str.strip() == str(rfc_input).strip())
            oficios = df_anexo_v[filtro]
            
            if not oficios.empty:
                for idx, fila in oficios.iterrows():
                    plaza_info = {
                        'NO_COMPROBANTE': fila.get('NO_COMPROBANTE', ''),
                        'CLAVE_PLAZA': fila.get('CLAVE_PLAZA', ''),
                        'PERIODO': fila.get('PERIODO', ''),
                        'nombre_completo': f"{fila.get('PRIMER_APELLIDO', '')} {fila.get('SEGUNDO_APELLIDO', '')} {fila.get('NOMBRE(S)', '')}".strip()
                    }
                    plazas.append(plaza_info)
    except Exception as e:
        return False, f"Error al leer los Anexos V: {e}", []
    
    if not plazas:
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
    """
    Motor principal.
    ruta_anexo_v: puede ser str o lista de str
    ruta_anexo_vi: puede ser str o lista de str
    no_comprobantes_seleccionados: lista de comprobantes a procesar. Si es None, procesa todos.
    progress_callback: función que recibe (actual, total) para reportar progreso
    """
    
    # Convertir a listas si son strings
    if isinstance(ruta_anexo_v, str):
        ruta_anexo_v = [ruta_anexo_v]
    if isinstance(ruta_anexo_vi, str):
        ruta_anexo_vi = [ruta_anexo_vi]
    
    print("Iniciando 'py_reintegros' (Modo PDF con capas)...")

    # --- Validar rutas ---
    for ruta in ruta_anexo_v:
        if not os.path.exists(ruta):
            return False, f"Error: No se encontró el Anexo V en:\n{ruta}"
    for ruta in ruta_anexo_vi:
        if not os.path.exists(ruta):
            return False, f"Error: No se encontró el Anexo VI en:\n{ruta}"
    
    if not os.path.exists(ruta_carpeta_salida):
        try:
            os.makedirs(ruta_carpeta_salida)
        except Exception as e:
            return False, f"Error: No se pudo crear la carpeta de salida:\n{e}"

    # --- Cargar Anexos (múltiples) ---
    print("Cargando Anexos (esto puede tardar)...")
    try:
        df_anexo_v_list = [pd.read_excel(ruta, dtype=str) for ruta in ruta_anexo_v]
        df_anexo_v = pd.concat(df_anexo_v_list, ignore_index=True)
        
        df_anexo_vi_list = [pd.read_excel(ruta, dtype=str) for ruta in ruta_anexo_vi]
        df_anexo_vi = pd.concat(df_anexo_vi_list, ignore_index=True)
        
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'], errors='coerce').fillna(0.0)
    except Exception as e:
        msg = f"ERROR al leer los archivos Excel: {e}"
        print(msg)
        return False, msg
    print("Anexos cargados.")

    # --- Obtener PDF de fondo ---
    pdf_fondo = obtener_pdf_fondo()
    if not pdf_fondo:
        msg = "ERROR: No se encontró 'fondo_reintegro.pdf'."
        print(msg)
        return False, msg

    # --- Configurar Jinja2 ---
    try:
        script_dir = os.path.dirname(__file__)
        env = Environment(loader=FileSystemLoader(script_dir if script_dir else '.'))
        template = env.get_template(ARCHIVO_PLANTILLA_HTML)
    except Exception as e:
        msg = f"ERROR: No se encontró la plantilla '{ARCHIVO_PLANTILLA_HTML}'. Error: {e}"
        print(msg)
        return False, msg

    # --- Filtrar Anexo V por RFC ---
    filtro = (df_anexo_v['RFC'].astype(str).str.strip() == str(rfc_input).strip())
    oficios_encontrados = df_anexo_v[filtro]

    if oficios_encontrados.empty:
        msg = f"--- ERROR: No se encontró ningún registro que coincida con el RFC: {rfc_input} ---"
        print(msg)
        return False, msg

    # --- Filtrar solo los comprobantes seleccionados ---
    if no_comprobantes_seleccionados:
        oficios_encontrados = oficios_encontrados[
            oficios_encontrados['NO_COMPROBANTE'].isin(no_comprobantes_seleccionados)
        ]
        if oficios_encontrados.empty:
            msg = "ERROR: Ninguno de los comprobantes seleccionados fueron encontrados."
            print(msg)
            return False, msg

    print(f"¡Coincidencia encontrada! Se procesarán {len(oficios_encontrados)} oficio(s) para este RFC.")
    
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
        print(f"\nProcesando: {rfc} - Comprobante: {no_comprobante}")
        nombre_completo = f"{fila_v.get('PRIMER_APELLIDO', '')} {fila_v.get('SEGUNDO_APELLIDO', '')} {fila_v.get('NOMBRE(S)', '')}".strip()
        qna = fila_v.get('PERIODO', '')
        detalles = df_anexo_vi[df_anexo_vi['NO_COMPROBANTE'] == no_comprobante]
        percepciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'P']
        deducciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'D']
        total_percepciones = percepciones_df['IMPORTE'].sum()
        total_deducciones = deducciones_df['IMPORTE'].sum()
        total_liquido = total_percepciones - total_deducciones
        total_liquido = truncar_a_2_decimales(total_liquido)
        total_a_reintegrar = 0.0
        
        if config_reintegro['tipo'] == 'TOTAL':
            total_a_reintegrar = total_liquido
        else:
            if config_reintegro['modo'] == 'DIAS':
                total_a_reintegrar = truncar_a_2_decimales((total_liquido / 15.0) * float(config_reintegro['dias']))
            else:
                # MODO CONCEPTO - MÚLTIPLES CONCEPTOS
                conceptos_str = str(config_reintegro.get('concepto', '')).strip()
                if not conceptos_str:
                    total_a_reintegrar = 0.0
                else:
                    # Separar conceptos por coma y limpiar espacios
                    conceptos_lista = [c.strip() for c in conceptos_str.split(',') if c.strip()]
                    total_concepto = 0.0
                    
                    for concepto_buscar in conceptos_lista:
                        # Buscar el concepto en percepciones
                        s = percepciones_df['COD_CONCEPTO'].astype(str).str.strip()
                        concepto_df = percepciones_df[s == concepto_buscar]
                        if concepto_df.empty:
                            concepto_df = percepciones_df[s.str.lstrip('0') == concepto_buscar.lstrip('0')]
                        
                        if not concepto_df.empty:
                            importe_concepto = concepto_df['IMPORTE'].astype(float).sum()
                            total_concepto += importe_concepto
                    
                    if config_reintegro.get('por_dias'):
                        total_a_reintegrar = truncar_a_2_decimales((total_concepto / 15.0) * float(config_reintegro['dias']))
                    else:
                        total_a_reintegrar = truncar_a_2_decimales(total_concepto)
        
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
        ruta_temporal = os.path.join(ruta_carpeta_salida, nombre_archivo_temporal)
        exito_pdf, mensaje_pdf = convertir_html_a_pdf(html_final_renderizado, ruta_temporal)
        if not exito_pdf:
            return False, mensaje_pdf
        nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
        ruta_final = os.path.join(ruta_carpeta_salida, nombre_archivo_final)
        exito_superponer, mensaje_superponer = superponer_pdfs(ruta_temporal, pdf_fondo, ruta_final)
        try: os.remove(ruta_temporal)
        except OSError: pass
        if not exito_superponer:
            return False, mensaje_superponer
        mensajes_exito.append(f"Éxito: {nombre_archivo_final}")

        # Reportar progreso
        if progress_callback:
            progress_callback(idx, total_oficios)

    # Si todo salió bien
    return True, "\n".join(mensajes_exito)


# --- 5. BLOQUE DE PRUEBA ---
if __name__ == "__main__":
    print("--- MODO DE PRUEBA DEL MOTOR ---")