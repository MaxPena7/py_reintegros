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

# --- 2. FUNCIONES AUXILIARES (Sin cambios) ---

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

# --- 3. FUNCIÓN PRINCIPAL DEL MOTOR ---

def generar_reintegros_pdf(
        rfc_input, 
        config_reintegro, 
        datos_manuales_input,
        ruta_anexo_v,
        ruta_anexo_vi,
        ruta_carpeta_salida,
        progress_callback=None  # <-- AGREGADO
    ):
    """
    Motor principal.
    progress_callback: función que recibe (actual, total) para reportar progreso
    """
    
    print("Iniciando 'py_reintegros' (Modo PDF con capas)...")

    # --- Validar rutas (sin cambios) ---
    if not os.path.exists(ruta_anexo_v):
        return False, f"Error: No se encontró el Anexo V en:\n{ruta_anexo_v}"
    if not os.path.exists(ruta_anexo_vi):
        return False, f"Error: No se encontró el Anexo VI en:\n{ruta_anexo_vi}"
    if not os.path.exists(ruta_carpeta_salida):
        try:
            os.makedirs(ruta_carpeta_salida)
        except Exception as e:
            return False, f"Error: No se pudo crear la carpeta de salida:\n{e}"

    # --- Cargar Anexos (sin cambios) ---
    print("Cargando Anexos (esto puede tardar)...")
    try:
        df_anexo_v = pd.read_excel(ruta_anexo_v, dtype=str)
        df_anexo_vi = pd.read_excel(ruta_anexo_vi, dtype=str)
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'], errors='coerce').fillna(0.0)
    except Exception as e:
        msg = f"ERROR al leer los archivos Excel: {e}"
        print(msg)
        return False, msg
    print("Anexos cargados.")

    # --- Obtener PDF de fondo (sin cambios) ---
    pdf_fondo = obtener_pdf_fondo()
    if not pdf_fondo:
        msg = "ERROR: No se encontró 'fondo_reintegro.pdf'."
        print(msg)
        return False, msg

    # --- Configurar Jinja2 (sin cambios) ---
    try:
        script_dir = os.path.dirname(__file__)
        env = Environment(loader=FileSystemLoader(script_dir if script_dir else '.'))
        template = env.get_template(ARCHIVO_PLANTILLA_HTML)
    except Exception as e:
        msg = f"ERROR: No se encontró la plantilla '{ARCHIVO_PLANTILLA_HTML}'. Error: {e}"
        print(msg)
        return False, msg

    # --- Filtrar Anexo V por RFC (sin cambios) ---
    filtro = (df_anexo_v['RFC'].astype(str).str.strip() == str(rfc_input).strip())
    oficios_encontrados = df_anexo_v[filtro]

    if oficios_encontrados.empty:
        msg = f"--- ERROR: No se encontró ningún registro que coincida con el RFC: {rfc_input} ---"
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
    total_oficios = len(oficios_encontrados)  # <-- AGREGADO

    # --- Iterar sobre CADA oficio encontrado ---
    for idx, (index, fila_v) in enumerate(oficios_encontrados.iterrows(), start=1):  # <-- MODIFICADO
        
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
                concepto_buscar = str(config_reintegro.get('concepto', '')).strip()
                if not concepto_buscar:
                    total_concepto = 0.0
                else:
                    s = percepciones_df['COD_CONCEPTO'].astype(str).str.strip()
                    concepto_df = percepciones_df[s == concepto_buscar]
                    if concepto_df.empty:
                        concepto_df = percepciones_df[s.str.lstrip('0') == concepto_buscar.lstrip('0')]
                    if concepto_df.empty:
                        total_concepto = 0.0
                    else:
                        total_concepto = concepto_df['IMPORTE'].astype(float).sum()
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

        # <-- AGREGADO: Reportar progreso
        if progress_callback:
            progress_callback(idx, total_oficios)

    # Si todo salió bien
    return True, "\n".join(mensajes_exito)


# --- 4. BLOQUE DE PRUEBA ---
def solicitar_tipo_reintegro():
    while True:
        print("\n=== TIPO DE REINTEGRO ===")
        tipo = input("Ingrese tipo de reintegro (TOTAL/PARCIAL): ").strip().upper()
        if tipo not in ['TOTAL', 'PARCIAL']:
            print("Error: Ingrese TOTAL o PARCIAL")
            continue
        if tipo == 'TOTAL':
            return {'tipo': 'TOTAL','modo': None,'dias': 15,'concepto': None,'por_dias': False}
        print("\n=== MODO DE REINTEGRO PARCIAL ===")
        print("1. Por días")
        print("2. Por concepto")
        modo = input("Seleccione una opción (1/2): ").strip()
        if modo == '1':
            while True:
                try:
                    dias = int(input("\nIngrese número de días a reintegrar (1-15): "))
                    if 1 <= dias <= 15:
                        return {'tipo': 'PARCIAL','modo': 'DIAS','dias': dias,'concepto': None,'por_dias': False}
                    print("Error: El número de días debe estar entre 1 y 15")
                except ValueError:
                    print("Error: Ingrese un número válido")
        elif modo == '2':
            concepto = input("\nIngrese el concepto a reintegrar (ejemplo: 07): ").strip()
            por_dias = input("¿El reintegro del concepto es por días? (S/N): ").strip().upper()
            if por_dias == 'S':
                while True:
                    try:
                        dias = int(input("Ingrese número de días para el concepto (1-15): "))
                        if 1 <= dias <= 15:
                            return {'tipo': 'PARCIAL','modo': 'CONCEPTO','dias': dias,'concepto': concepto,'por_dias': True}
                        print("Error: El número de días debe estar entre 1 y 15")
                    except ValueError:
                        print("Error: Ingrese un número válido")
            elif por_dias == 'N':
                return {'tipo': 'PARCIAL','modo': 'CONCEPTO','dias': 15,'concepto': concepto,'por_dias': False}
        print("Opción no válida. Intente de nuevo.")


if __name__ == "__main__":
    print("--- MODO DE PRUEBA DEL MOTOR ---")
    RUTA_ANEXO_V_PRUEBA = r"C:\Users\Maxruso7\Desktop\ANEXOS\R06_202518_O1_AnexoV.xlsx"
    RUTA_ANEXO_VI_PRUEBA = r"C:\Users\Maxruso7\Desktop\ANEXOS\R06_202518_O1_AnexoVI.xlsx"
    RUTA_SALIDA_PRUEBA = r"C:\Users\Maxruso7\Desktop\py_reintegros"
    
    rfc_prueba = input("RFC para prueba (ej: BISA841115H59): ").strip().upper()
    if not rfc_prueba:
        rfc_prueba = "BISA841115H59"
    
    config_prueba = solicitar_tipo_reintegro()
    
    datos_manuales_prueba = {
        "CAMPO_ABAJO_MOTIVO": 'Prueba de motor desde __main__',
        "NIVEL_EDUCATIVO": "PREESCOLAR (PRUEBA)"
    }
    
    def callback_prueba(actual, total):
        print(f"Progreso: {actual}/{total} ({actual*100//total}%)")
    
    exito, mensaje = generar_reintegros_pdf(
        rfc_prueba, 
        config_prueba, 
        datos_manuales_prueba,
        RUTA_ANEXO_V_PRUEBA,
        RUTA_ANEXO_VI_PRUEBA,
        RUTA_SALIDA_PRUEBA,
        progress_callback=callback_prueba
    )
    
    if exito:
        print("\n--- PRUEBA EXITOSA ---")
        print(mensaje)
    else:
        print("\n--- PRUEBA FALLIDA ---")
        print(mensaje)