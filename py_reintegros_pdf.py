import pandas as pd
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
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
# DEFINICION DEL TIPO DE REINTEGRO
def solicitar_tipo_reintegro():
    """Solicita y valida el tipo de reintegro y sus parámetros."""
    while True:
        print("\n=== TIPO DE REINTEGRO ===")
        tipo = input("Ingrese tipo de reintegro (TOTAL/PARCIAL): ").strip().upper()
        if tipo not in ['TOTAL', 'PARCIAL']:
            print("Error: Ingrese TOTAL o PARCIAL")
            continue

        if tipo == 'TOTAL':
            return {
                'tipo': 'TOTAL',
                'modo': None,
                'dias': 15,
                'concepto': None,
                'por_dias': False
            }

        # Manejo de reintegro PARCIAL
        print("\n=== MODO DE REINTEGRO PARCIAL ===")
        print("1. Por días")
        print("2. Por concepto")
        modo = input("Seleccione una opción (1/2): ").strip()

        if modo == '1':  # PARCIAL POR DÍAS
            while True:
                try:
                    dias = int(input("\nIngrese número de días a reintegrar (1-15): "))
                    if 1 <= dias <= 15:
                        return {
                            'tipo': 'PARCIAL',
                            'modo': 'DIAS',
                            'dias': dias,
                            'concepto': None,
                            'por_dias': False
                        }
                    print("Error: El número de días debe estar entre 1 y 15")
                except ValueError:
                    print("Error: Ingrese un número válido")

        elif modo == '2':  # PARCIAL POR CONCEPTO
            concepto = input("\nIngrese el concepto a reintegrar (ejemplo: 07): ").strip()
            por_dias = input("¿El reintegro del concepto es por días? (S/N): ").strip().upper()
            
            if por_dias == 'S':
                while True:
                    try:
                        dias = int(input("Ingrese número de días para el concepto (1-15): "))
                        if 1 <= dias <= 15:
                            return {
                                'tipo': 'PARCIAL',
                                'modo': 'CONCEPTO',
                                'dias': dias,
                                'concepto': concepto,
                                'por_dias': True
                            }
                        print("Error: El número de días debe estar entre 1 y 15")
                    except ValueError:
                        print("Error: Ingrese un número válido")
            elif por_dias == 'N':
                return {
                    'tipo': 'PARCIAL',
                    'modo': 'CONCEPTO',
                    'dias': 15,
                    'concepto': concepto,
                    'por_dias': False
                }
            
        print("Opción no válida. Intente de nuevo.")


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
    import math

    print("Iniciando 'py_reintegros' (Modo PDF con capas)...")

    # Solicitar configuración de reintegro (devuelve dict con keys: tipo, modo, dias, concepto, por_dias)
    config_reintegro = solicitar_tipo_reintegro()
    print(f"\nProcesando reintegro: {config_reintegro}")

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
        # Asegurar que IMPORTE es numérico
        df_anexo_vi['IMPORTE'] = pd.to_numeric(df_anexo_vi['IMPORTE'], errors='coerce').fillna(0.0)
        print("[DEBUG] Anexos cargados correctamente")
    except FileNotFoundError as e:
        print(f"ERROR: No se encontro el archivo {e.filename}. Revisa las rutas.")
        return
    except Exception as e:
        print(f"ERROR al leer los archivos Excel: {e}")
        return

    print(f"Anexo V cargado: {len(df_anexo_v)} registros totales.")
    print(f"Anexo VI cargado: {len(df_anexo_vi)} registros totales.")

    # Obtener PDF de fondo
    print("\n--- Obteniendo PDF de fondo pre-diseñado ---")
    pdf_fondo = obtener_pdf_fondo()
    print(f"[DEBUG] PDF fondo: {pdf_fondo}")

    # Configurar Jinja2
    print("[DEBUG] Configurando Jinja2...")
    env = Environment(loader=FileSystemLoader('.'))
    try:
        template = env.get_template(ARCHIVO_PLANTILLA_HTML)
        print("[DEBUG] Plantilla cargada correctamente")
    except Exception as e:
        print(f"ERROR: No se encontro la plantilla '{ARCHIVO_PLANTILLA_HTML}'. Error: {e}")
        return

    # Filtrar Anexo V por RFC
    print(f"[DEBUG] Buscando RFC: {INPUT_RFC}")
    filtro = (df_anexo_v['RFC'].astype(str) == str(INPUT_RFC))
    oficios_encontrados = df_anexo_v[filtro]
    print(f"[DEBUG] Registros encontrados: {len(oficios_encontrados)}")

    if oficios_encontrados.empty:
        print(f"\n--- ERROR: No se encontro ningun registro que coincida con el RFC: {INPUT_RFC} ---")
        return

    print(f"\nCoincidencia encontrada! Se procesaran {len(oficios_encontrados)} oficio(s) para este RFC.")

    # Fecha en español (sin locale)
    dias_semana = {
        0: "lunes", 1: "martes", 2: "miércoles",
        3: "jueves", 4: "viernes", 5: "sábado", 6: "domingo"
    }
    meses_nombre = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    ahora = datetime.now()
    fecha_hoy_str = f"{dias_semana[ahora.weekday()].capitalize()}, {ahora.day:02d} de {meses_nombre[ahora.month]} de {ahora.year}"
    print(f"[DEBUG] Fecha configurada: {fecha_hoy_str}")

    # Helper para truncar (no redondear) a 2 decimales
    def truncar_a_2_decimales(valor):
        return math.trunc(float(valor) * 100) / 100.0

    # Iterar sobre cada oficio encontrado
    for index, fila_v in oficios_encontrados.iterrows():

        no_comprobante = fila_v.get('NO_COMPROBANTE', 'S/C')
        rfc = fila_v.get('RFC', 'S/RFC')
        print(f"\n[DEBUG] Procesando: {rfc} - Comprobante: {no_comprobante}")

        # Preparar datos básicos
        nombre_completo = f"{fila_v.get('PRIMER_APELLIDO', '')} {fila_v.get('SEGUNDO_APELLIDO', '')} {fila_v.get('NOMBRE(S)', '')}".strip()
        qna = fila_v.get('PERIODO', '')

        # Filtrar detalles del Anexo VI por comprobante
        detalles = df_anexo_vi[df_anexo_vi['NO_COMPROBANTE'] == no_comprobante]
        print(f"[DEBUG] Detalles encontrados: {len(detalles)}")

        # Separar percepciones y deducciones
        percepciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'P']
        deducciones_df = detalles[detalles['TIPO_CONCEPTO'] == 'D']

        # Totales
        total_percepciones = percepciones_df['IMPORTE'].sum()
        total_deducciones = deducciones_df['IMPORTE'].sum()
        total_liquido = total_percepciones - total_deducciones
        total_liquido = truncar_a_2_decimales(total_liquido)

        # Calcular total_a_reintegrar según config_reintegro
        if config_reintegro['tipo'] == 'TOTAL':
            total_a_reintegrar = total_liquido
        else:
            # PARCIAL
            if config_reintegro['modo'] == 'DIAS':
                total_a_reintegrar = truncar_a_2_decimales((total_liquido / 15.0) * float(config_reintegro['dias']))
            else:
                # Por concepto: buscar en la columna 'COD_CONCEPTO' de Anexo VI y sumar 'IMPORTE'
                concepto_buscar = str(config_reintegro.get('concepto', '')).strip()
                if not concepto_buscar:
                    print("[ADVERTENCIA] No se especificó concepto para reintegro por concepto.")
                    total_concepto = 0.0
                else:
                    col = 'COD_CONCEPTO'
                    if col not in percepciones_df.columns:
                        print(f"[ADVERTENCIA] La columna '{col}' no existe en percepciones. Columnas: {list(percepciones_df.columns)}")
                        total_concepto = 0.0
                    else:
                        s = percepciones_df[col].astype(str).str.strip()
                        # 1) coincidencia exacta
                        concepto_df = percepciones_df[s == concepto_buscar]
                        # 2) intentar sin ceros a la izquierda
                        if concepto_df.empty:
                            concepto_df = percepciones_df[s.str.lstrip('0') == concepto_buscar.lstrip('0')]
                        # 3) comparación numérica
                        if concepto_df.empty:
                            try:
                                num_buscar = float(concepto_buscar)
                                mask = pd.to_numeric(s, errors='coerce') == num_buscar
                                concepto_df = percepciones_df[mask.fillna(False)]
                            except:
                                pass
                        # 4) contains (por si el usuario ingresó parte del código)
                        if concepto_df.empty:
                            concepto_df = percepciones_df[s.str.contains(concepto_buscar, na=False)]
                        if concepto_df.empty:
                            muestra = list(pd.unique(s))[:20]
                            print(f"[ADVERTENCIA] No se encontró el concepto '{concepto_buscar}'. Ejemplos disponibles (hasta 20): {muestra}")
                            total_concepto = 0.0
                        else:
                            total_concepto = concepto_df['IMPORTE'].astype(float).sum()
                if config_reintegro.get('por_dias'):
                    total_a_reintegrar = truncar_a_2_decimales((total_concepto / 15.0) * float(config_reintegro['dias']))
                else:
                    total_a_reintegrar = truncar_a_2_decimales(total_concepto)

        # Preparar listas para Jinja
        percepciones_list = percepciones_df.to_dict('records')
        deducciones_list = deducciones_df.to_dict('records')

        # Tipo rintegro texto
        motivo_texto = f"{config_reintegro['tipo']}"
        

        # Contexto para la plantilla (mantener números como float)
        contexto_datos = {
            "nombre_completo": nombre_completo,
            "rfc": rfc,
            "no_comprobante": no_comprobante,
            "clave_cobro": fila_v.get('CLAVE_PLAZA', ''),
            "cct": fila_v.get('CCT', ''),
            "campo_abajo_motivo": DATOS_MANUALES.get("CAMPO_ABAJO_MOTIVO", ""),
            "desde": fila_v.get('FECHA_INICIO', ''),
            "hasta": fila_v.get('FECHA_TERMINO', ''),
            "qna": qna,
            "fecha_hoy": fecha_hoy_str,
            "percepciones": percepciones_list,
            "deducciones": deducciones_list,
            "total_percepciones": float(truncar_a_2_decimales(total_percepciones)),
            "total_deducciones": float(truncar_a_2_decimales(total_deducciones)),
            "total_liquido": float(total_liquido),
            "total_a_reintegrar": float(total_a_reintegrar),
            "motivo_reintegro": motivo_texto,
            "nivel_educativo": DATOS_MANUALES.get("NIVEL_EDUCATIVO", "")
        }

        # DEBUG: revisar contexto antes de render
        print("DEBUG contexto_datos:", {k: (v if k not in ('percepciones','deducciones') else f"len={len(v)}") for k,v in contexto_datos.items()})
        print("[DEBUG] Renderizando HTML...")
        html_final_renderizado = template.render(contexto_datos)

        # Guardar PDF temporal y superponer con fondo
        nombre_archivo_temporal = f"{rfc}_{no_comprobante}_temp.pdf"
        ruta_temporal = os.path.join(CARPETA_SALIDA, nombre_archivo_temporal)
        print(f"[DEBUG] Generando PDF temporal: {ruta_temporal}")
        convertir_html_a_pdf(html_final_renderizado, ruta_temporal)

        # Superponer fondo si existe
        if pdf_fondo and os.path.exists(pdf_fondo):
            print(f"[DEBUG] Superponiendo PDFs...")
            nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
            ruta_final = os.path.join(CARPETA_SALIDA, nombre_archivo_final)
            superponer_pdfs(ruta_temporal, pdf_fondo, ruta_final)
            try:
                os.remove(ruta_temporal)
            except OSError:
                pass
            print(f"[DEBUG] PDF final generado: {ruta_final}")
        else:
            print(f"[ADVERTENCIA] No se superpuso fondo, renombrando a final...")
            nombre_archivo_final = f"{rfc}_{no_comprobante}.pdf"
            ruta_final = os.path.join(CARPETA_SALIDA, nombre_archivo_final)
            os.rename(ruta_temporal, ruta_final)
            print(f"[DEBUG] PDF sin fondo generado: {ruta_final}")

    # Al finalizar
    if pdf_fondo and os.path.exists(pdf_fondo):
        print("[OK] Proceso completado, PDF de fondo NO eliminado (es permanente)")

    print("\n--- Proceso completado ---")

# --- 7. Ejecutar el script ---
if __name__ == "__main__":
    generar_reintegros_pdf()