import pandas as pd
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
import locale
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import PyPDF2

# --- 1. CONFIGURACIÓN DE ARCHIVOS Y RUTA ---
ARCHIVO_PLANTILLA_HTML = "plantilla.html"
ARCHIVO_ANEXO_V = r"C:\Users\NominaAdmin\Desktop\ANEXOS\ANEXOS V Y VI\ANEXOS V Y VI 2025\202516\R06_202516_O1_AnexoV.xlsx"
ARCHIVO_ANEXO_VI = r"C:\Users\NominaAdmin\Desktop\ANEXOS\ANEXOS V Y VI\ANEXOS V Y VI 2025\202516\R06_202516_O1_AnexoVI.xlsx"
CARPETA_SALIDA = r"C:\Users\NominaAdmin\Desktop\reintegros_prueba"

# Configurar la ruta de wkhtmltopdf
RUTA_WKHTMLTOPDF = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# Opciones para pdfkit
options = {
    'page-size': 'Letter',
    'encoding': "UTF-8",
    'no-outline': None,
    'enable-local-file-access': None,
    'margin-top': '35mm',   # antes ~5mm por defecto, ahora 55mm = 5.5 cm
    'margin-right': '15mm',
    'margin-bottom': '15mm',
    'margin-left': '15mm',
}

# --- 2. DATOS DE ENTRADA ---
print("--- Configuracion de Busqueda ---")
INPUT_RFC = "BISA841115H59"
print(f"Buscando todos los reintegros para el RFC: {INPUT_RFC}\n")

DATOS_MANUALES = {
    "MOTIVO_REINTEGRO": "REINTEGRO TOTAL",
    "CAMPO_ABAJO_MOTIVO": 'NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS, NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS, NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS',
    "NIVEL_EDUCATIVO": "PREESCOLAR"
}

# --- 3. FUNCIÓN PARA CREAR PDF DE FONDO CON IMÁGENES ---
def obtener_pdf_fondo():
    """
    Obtiene el PDF de fondo pre-diseñado
    """
    pdf_fondo_path = "fondo_reintegro.pdf"
    
    if os.path.exists(pdf_fondo_path):
        print(f"[OK] PDF de fondo encontrado: {pdf_fondo_path}")
        return pdf_fondo_path
    else:
        print(f"[ERROR] No se encontro {pdf_fondo_path} en la carpeta del proyecto")
        print(f"[ERROR] Por favor, asegurate de que 'fondo_reintegro.pdf' este en: {os.getcwd()}")
        return None

# --- 4. FUNCIÓN PARA SUPERPONER PDFS ---
def superponer_pdfs(pdf_contenido, pdf_fondo, pdf_salida):
    """
    Superpone el PDF de contenido sobre el PDF de fondo
    """
    try:
        pdf_writer = PyPDF2.PdfWriter()
        
        # Leer PDF de fondo
        pdf_fondo_file = open(pdf_fondo, 'rb')
        pdf_fondo_reader = PyPDF2.PdfReader(pdf_fondo_file)
        pagina_fondo = pdf_fondo_reader.pages[0]
        
        # Leer PDF de contenido
        pdf_contenido_file = open(pdf_contenido, 'rb')
        pdf_contenido_reader = PyPDF2.PdfReader(pdf_contenido_file)
        
        # Iterar sobre cada página del contenido
        for pagina_contenido in pdf_contenido_reader.pages:
            # Crear página nueva en blanco
            pagina_nueva = PyPDF2.PageObject.create_blank_page(None, 612, 792)
            # Primero agregar fondo
            pagina_nueva.merge_page(pagina_fondo)
            # Luego agregar contenido encima
            pagina_nueva.merge_page(pagina_contenido)
            pdf_writer.add_page(pagina_nueva)
        
        # Guardar resultado
        with open(pdf_salida, 'wb') as f:
            pdf_writer.write(f)
        
        # Cerrar archivos
        pdf_fondo_file.close()
        pdf_contenido_file.close()
        
        print(f"[OK] PDFs superpuestos correctamente: {pdf_salida}")
        return True
    
    except Exception as e:
        print(f"[ERROR] No se pudo superponer PDFs: {e}")
        import traceback
        traceback.print_exc()
        return False

# --- 5. FUNCIÓN PARA CONVERTIR HTML A PDF ---
def convertir_html_a_pdf(html_string, ruta_pdf_salida):
    """
    Convierte una cadena de texto HTML a un archivo PDF usando wkhtmltopdf.
    """
    try:
        pdfkit.from_string(
            html_string,
            ruta_pdf_salida,
            options=options,
            configuration=pdfkit.configuration(wkhtmltopdf=RUTA_WKHTMLTOPDF)
        )
        print(f"   -> Exito! PDF Guardado en: {ruta_pdf_salida}")
    
    except Exception as e:
        print(f"   -> ERROR al crear PDF '{ruta_pdf_salida}': {e}")


# --- 6. FUNCIÓN PRINCIPAL ---
def generar_reintegros_pdf():
    print("Iniciando 'py_reintegros' (Modo PDF con capas)...")
    
    # Crear carpeta de salida si no existe
    if not os.path.exists(CARPETA_SALIDA):
        try:
            os.makedirs(CARPETA_SALIDA)
            print(f"Carpeta '{CARPETA_SALIDA}' creada.")
        except Exception as e:
            print(f"ERROR: No se pudo crear la carpeta de salida: {e}")
            return

    print("[DEBUG] Carpeta de salida lista")

    # Cargar los datos de los Anexos
    try:
        print("[DEBUG] Leyendo Anexo V...")
        df_anexo_v = pd.read_excel(ARCHIVO_ANEXO_V, dtype=str)
        print("[DEBUG] Leyendo Anexo VI...")
        df_anexo_vi = pd.read_excel(ARCHIVO_ANEXO_VI, dtype=str)
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'])
        print("[DEBUG] Anexos cargados correctamente")
    except FileNotFoundError as e:
        print(f"ERROR: No se encontro el archivo {e.filename}. Revisa las rutas.")
        return
    except Exception as e:
        print(f"ERROR al leer los archivos Excel: {e}")
        return

    print(f"Anexo V cargado: {len(df_anexo_v)} registros totales.")
    print(f"Anexo VI cargado: {len(df_anexo_vi)} registros totales.")
    
    # --- Obtener PDF de fondo pre-diseñado ---
    print("\n--- Obteniendo PDF de fondo pre-diseñado ---")
    pdf_fondo = obtener_pdf_fondo()
    print(f"[DEBUG] PDF fondo: {pdf_fondo}")
    
    # --- Configurar Jinja2 ---
    print("[DEBUG] Configurando Jinja2...")
    env = Environment(loader=FileSystemLoader('.'))
    try:
        template = env.get_template(ARCHIVO_PLANTILLA_HTML)
        print("[DEBUG] Plantilla cargada correctamente")
    except Exception as e:
        print(f"ERROR: No se encontro la plantilla '{ARCHIVO_PLANTILLA_HTML}'. Error: {e}")
        return

    # --- Filtrar Anexo V ---
    print(f"[DEBUG] Buscando RFC: {INPUT_RFC}")
    filtro = (df_anexo_v['RFC'].astype(str) == str(INPUT_RFC))
    oficios_encontrados = df_anexo_v[filtro]
    print(f"[DEBUG] Registros encontrados: {len(oficios_encontrados)}")

    if oficios_encontrados.empty:
        print(f"\n--- ERROR: No se encontro ningun registro que coincida con el RFC: {INPUT_RFC} ---")
        return

    print(f"\nCoincidencia encontrada! Se procesaran {len(oficios_encontrados)} oficio(s) para este RFC.")

    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'es_MX.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
            except:
                locale.setlocale(locale.LC_TIME, '')

    fecha_hoy_str = datetime.now().strftime("%A, %d de %B de %Y").capitalize()
    print(f"[DEBUG] Fecha configurada: {fecha_hoy_str}")

    # Iterar sobre CADA resultado encontrado para ese RFC
    for index, fila_v in oficios_encontrados.iterrows():
        
        no_comprobante = fila_v.get('NO_COMPROBANTE', 'S/C')
        rfc = fila_v.get('RFC', 'S/RFC')
        print(f"\n[DEBUG] Procesando: {rfc} - Comprobante: {no_comprobante}")

        # --- A. Preparar todos los datos ---
        nombre_completo = f"{fila_v.get('PRIMER_APELLIDO', '')} {fila_v.get('SEGUNDO_APELLIDO', '')} {fila_v.get('NOMBRE(S)', '')}"
        qna = fila_v.get('PERIODO', '')
        
        # Filtrar detalles del Anexo VI
        detalles = df_anexo_vi[df_anexo_vi['NO_COMPROBANTE'] == no_comprobante]
        print(f"[DEBUG] Detalles encontrados: {len(detalles)}")
        
        # Separar DataFrames para calcular totales
        percepciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'P']
        deducciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'D']
        
        # Calcular Totales
        total_percepciones = percepciones_df['IMPORTE'].sum()
        total_deducciones = deducciones_df['IMPORTE'].sum()
        total_liquido = total_percepciones - total_deducciones

        # Convertir a diccionarios para Jinja
        percepciones_list = percepciones_df.to_dict('records')
        deducciones_list = deducciones_df.to_dict('records')
        
        # Calcular Observaciones
        observaciones = (
            f"{DATOS_MANUALES['MOTIVO_REINTEGRO']}, "
            f"{DATOS_MANUALES['CAMPO_ABAJO_MOTIVO']}, "
            f"{qna}"
        )

        # Este diccionario tiene TODA la información que la plantilla HTML necesita
        contexto_datos = {
            "fecha_hoy": fecha_hoy_str,
            "nombre_completo": nombre_completo.strip(),
            "no_comprobante": no_comprobante,
            "qna": qna,
            "rfc": rfc,
            "clave_cobro": fila_v.get('CLAVE_PLAZA', ''),
            "cct": fila_v.get('CCT', ''),
            "motivo_reintegro": DATOS_MANUALES["MOTIVO_REINTEGRO"],
            "campo_abajo_motivo": DATOS_MANUALES["CAMPO_ABAJO_MOTIVO"],
            "desde": fila_v.get('FECHA_INICIO', ''),
            "hasta": fila_v.get('FECHA_TERMINO', ''),
            "nivel_educativo": DATOS_MANUALES["NIVEL_EDUCATIVO"],
            "observaciones": observaciones,
            "percepciones": percepciones_list,
            "deducciones": deducciones_list,
            "total_percepciones": total_percepciones,
            "total_deducciones": total_deducciones,
            "total_liquido": total_liquido
        }

        # --- B. Renderizar el HTML ---
        print("[DEBUG] Renderizando HTML...")
        html_final_renderizado = template.render(contexto_datos)
        
        # --- C. Guardar PDF de contenido (temporal) ---
        nombre_archivo_temporal = f"{rfc}_{no_comprobante}_temp.pdf"
        ruta_temporal = os.path.join(CARPETA_SALIDA, nombre_archivo_temporal)
        print(f"[DEBUG] Generando PDF temporal: {ruta_temporal}")
        convertir_html_a_pdf(html_final_renderizado, ruta_temporal)
        
        # --- D. Superponer con PDF de fondo ---
        if pdf_fondo and os.path.exists(pdf_fondo):
            print(f"[DEBUG] Superponiendo PDFs...")
            nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
            ruta_final = os.path.join(CARPETA_SALIDA, nombre_archivo_final)
            superponer_pdfs(ruta_temporal, pdf_fondo, ruta_final)
            os.remove(ruta_temporal)  # Eliminar temporal
            print(f"[DEBUG] PDF final generado: {ruta_final}")
        else:
            print(f"[ADVERTENCIA] No se superpuso fondo, renombrando a final...")
            nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
            ruta_final = os.path.join(CARPETA_SALIDA, nombre_archivo_final)
            os.rename(ruta_temporal, ruta_final)
            print(f"[DEBUG] PDF sin fondo generado: {ruta_final}")
    
    # Limpiar archivos temporales
    if pdf_fondo and os.path.exists(pdf_fondo):
        # NO eliminar el fondo, es el que diseño el usuario
        print("[OK] Proceso completado, PDF de fondo NO eliminado (es permanente)")

    print("\n--- Proceso completado ---")

# --- 7. Ejecutar el script ---
if __name__ == "__main__":
    generar_reintegros_pdf()