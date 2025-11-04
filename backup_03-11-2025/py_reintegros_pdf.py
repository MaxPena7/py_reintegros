import pandas as pd
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
import locale

# --- 1. CONFIGURACIÓN DE ARCHIVOS Y RUTA ---
# Ajusta los nombres de tus archivos
ARCHIVO_PLANTILLA_HTML = "plantilla.html"  # La plantilla HTML que creamos
ARCHIVO_ANEXO_V = r"C:\Users\Maxruso7\Desktop\ANEXOS\R06_202518_O1_AnexoV.xlsx"
ARCHIVO_ANEXO_VI = r"C:\Users\Maxruso7\Desktop\ANEXOS\R06_202518_O1_AnexoVI.xlsx"
CARPETA_SALIDA = r"C:\Users\Maxruso7\Desktop\py_reintegros"

# Configurar la ruta de wkhtmltopdf
RUTA_WKHTMLTOPDF = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# Opciones para pdfkit
options = {
    'page-size': 'Letter',
    'encoding': "UTF-8",
    'no-outline': None,
    'enable-local-file-access': None,
}

# --- 2. DATOS DE ENTRADA (AQUÍ INGRESAS LOS DATOS) ---
print("--- Configuración de Búsqueda ---")
INPUT_RFC = "EASM911113LG1"  # Escribe el RFC a buscar
print(f"Buscando todos los reintegros para el RFC: {INPUT_RFC}\n")

# Datos que se ingresarán manualmente para el formato
DATOS_MANUALES = {
    "MOTIVO_REINTEGRO": "REINTEGRO TOTAL",
    "CAMPO_ABAJO_MOTIVO": 'NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS, NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS, NO CORRESPONDE TODA LA QUINCENA POR INCOMPATIBILIDAD LABORAL DE HORARIOS',
    "NIVEL_EDUCATIVO": "PREESCOLAR"
}

# --- 3. FUNCIÓN PARA CONVERTIR HTML A PDF (pdfkit + wkhtmltopdf) ---
def convertir_html_a_pdf(html_string, ruta_pdf_salida):
    """
    Convierte una cadena de texto HTML a un archivo PDF usando wkhtmltopdf.
    """
    try:
        # Convertir HTML string a PDF usando pdfkit
        pdfkit.from_string(
            html_string,
            ruta_pdf_salida,
            options=options,
            configuration=pdfkit.configuration(wkhtmltopdf=RUTA_WKHTMLTOPDF)
        )
        print(f"   -> ¡Éxito! PDF Guardado en: {ruta_pdf_salida}")
    
    except Exception as e:
        print(f"   -> ERROR al crear PDF '{ruta_pdf_salida}': {e}")


# --- 4. FUNCIÓN PRINCIPAL ---
def generar_reintegros_pdf():
    print("Iniciando 'py_reintegros' (Modo PDF con wkhtmltopdf)...")
    
    # Crear carpeta de salida si no existe
    if not os.path.exists(CARPETA_SALIDA):
        try:
            os.makedirs(CARPETA_SALIDA)
            print(f"Carpeta '{CARPETA_SALIDA}' creada.")
        except Exception as e:
            print(f"ERROR: No se pudo crear la carpeta de salida: {e}")
            return

    # Cargar los datos de los Anexos
    try:
        df_anexo_v = pd.read_excel(ARCHIVO_ANEXO_V, dtype=str)
        df_anexo_vi = pd.read_excel(ARCHIVO_ANEXO_VI, dtype=str)
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'])
    except FileNotFoundError as e:
        print(f"ERROR: No se encontró el archivo {e.filename}. Revisa las rutas.")
        return
    except Exception as e:
        print(f"ERROR al leer los archivos Excel: {e}")
        return

    print(f"Anexo V cargado: {len(df_anexo_v)} registros totales.")
    print(f"Anexo VI cargado: {len(df_anexo_vi)} registros totales.")
    
    # --- Configurar Jinja2 ---
    # Le decimos a Jinja que busque plantillas en la carpeta actual ('.')
    env = Environment(loader=FileSystemLoader('.'))
    try:
        template = env.get_template(ARCHIVO_PLANTILLA_HTML)
    except Exception as e:
        print(f"ERROR: No se encontró la plantilla '{ARCHIVO_PLANTILLA_HTML}'. Error: {e}")
        return

    # --- Filtrar Anexo V ---
    filtro = (df_anexo_v['RFC'].astype(str) == str(INPUT_RFC))
    oficios_encontrados = df_anexo_v[filtro]

    if oficios_encontrados.empty:
        print(f"\n--- ERROR: No se encontró ningún registro que coincida con el RFC: {INPUT_RFC} ---")
        return

    print(f"\n¡Coincidencia encontrada! Se procesarán {len(oficios_encontrados)} oficio(s) para este RFC.")

    try:
        # Intentamos establecer el idioma español según el sistema operativo
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Linux / macOS
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'es_MX.UTF-8')  # México (si está disponible)
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain')  # Windows (España)
            except:
                locale.setlocale(locale.LC_TIME, '')  # Usa configuración regional del sistema

    # Crear la fecha con formato largo y en español
    fecha_hoy_str = datetime.now().strftime("%A, %d de %B de %Y").capitalize()

    # Iterar sobre CADA resultado encontrado para ese RFC
    for index, fila_v in oficios_encontrados.iterrows():
        
        no_comprobante = fila_v.get('NO_COMPROBANTE', 'S/C')
        rfc = fila_v.get('RFC', 'S/RFC')
        print(f"\nProcesando: {rfc} - Comprobante: {no_comprobante}")

        # --- A. Preparar todos los datos (el "Contexto") ---
        nombre_completo = f"{fila_v.get('PRIMER_APELLIDO', '')} {fila_v.get('SEGUNDO_APELLIDO', '')} {fila_v.get('NOMBRE(S)', '')}"
        qna = fila_v.get('PERIODO', '')
        
        # Filtrar detalles del Anexo VI
        detalles = df_anexo_vi[df_anexo_vi['NO_COMPROBANTE'] == no_comprobante]
        
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
        # Inyecta los datos del 'contexto_datos' en la plantilla
        html_final_renderizado = template.render(contexto_datos)
        
        # --- C. Guardar el PDF ---
        nombre_archivo_salida = f"{rfc}_{no_comprobante}.pdf"
        ruta_salida = os.path.join(CARPETA_SALIDA, nombre_archivo_salida)
        
        # Usamos nuestra función para convertir el HTML a PDF
        convertir_html_a_pdf(html_final_renderizado, ruta_salida)

    print("\n--- Proceso completado ---")

# --- 5. Ejecutar el script ---
if __name__ == "__main__":
    generar_reintegros_pdf()