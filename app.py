import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from fpdf import FPDF

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de PDI", page_icon="📄", layout="centered")

st.title("📄 Generador de Planes de Desarrollo Individual (PDI)")
st.write("Esta aplicación genera un PDI en PDF para cada empleado a partir de los datos de una hoja de cálculo de Google.")

# --- CONEXIÓN SEGURA CON GOOGLE SHEETS ---
def conectar_google_sheets():
    """Conecta con Google Sheets usando los secretos de Streamlit."""
    try:
        scopes = [
            "https://spreadsheets.google.com/feeds",
            'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
        # Carga las credenciales desde los secretos de Streamlit (para despliegue)
        # o desde un archivo local (para pruebas)
        if "gcp_service_account" in st.secrets:
            creds_dict = st.secrets["gcp_service_account"]
        else:
            # Para ejecutarlo localmente, asegúrate de que credentials.json está en la misma carpeta
            import json
            with open("credentials.json", "r") as f:
                creds_dict = json.load(f)

        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        st.error("No se encontró el archivo 'credentials.json'. Asegúrate de que esté en la misma carpeta que app.py para pruebas locales.")
        return None
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        st.info("Si la app está desplegada, asegúrate de haber configurado los secretos correctamente en Streamlit Community Cloud.")
        return None

# --- GENERACIÓN DE PDF (VERSIÓN COMPLETA) ---
def generar_pdf(datos_empleado):
    """Genera un archivo PDF completo con el formato del PDI."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    # --- TÍTULO ---
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "PLAN DE DESARROLLO INDIVIDUAL (PDI)", ln=True, align='C')
    pdf.ln(10)

    # Función auxiliar para agregar campos y manejar valores que no existen
    def agregar_campo(etiqueta, valor_columna):
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(65, 6, f"{etiqueta}:", border=0)
        pdf.set_font("Arial", '', 10)
        valor = str(datos_empleado.get(valor_columna, 'N/A')) # Usa N/A si la columna no existe
        pdf.multi_cell(0, 6, valor, border=0)

    # --- SECCIONES DEL PDI ---
    def agregar_seccion(titulo, campos):
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, titulo, ln=True, align='L')
        pdf.set_font("Arial", size=10)
        for etiqueta, columna in campos.items():
            agregar_campo(etiqueta, columna)
        pdf.ln(5)

    agregar_seccion("1. Datos Personales y Laborales", {
        "Apellido y nombre": "Apellido y nombre", "DNI": "DNI",
        "Correo electrónico": "Correo electrónico", "Número de contacto": "Número de contacto",
        "Edad": "Edad", "Posición actual": "Posición actual",
        "Fecha de ingreso": "Fecha de ingreso", "Superior inmediato": "Superior inmediato"
    })

    agregar_seccion("2. Formación y Nivel Educativo", {
        "Nivel educativo alcanzado": "Nivel educativo alcanzado",
        "Carrera / Título obtenido": "Carrera / Título obtenido",
        "Otras capacitaciones": "Otras capacitaciones realizadas fuera de la empresa"
    })

    agregar_seccion("3. Interés de Desarrollo", {
        "¿Interés en desarrollar carrera?": "¿Le interesa desarrollar su carrera dentro de la empresa?",
        "Áreas de interés": "Áreas de interés para desarrollarse",
        "Puesto al que aspira": "Tipo de puesto al que aspira",
        "Motivaciones para cambiar": "Motivaciones para cambiar de posición"
    })

    agregar_seccion("4. Necesidades de Capacitación", {
        "Competencias a capacitar": "Competencias en las que desea capacitarse",
        "Interés formativo específico": "Especificación del interés formativo"
    })

    agregar_seccion("5. Fortalezas y Obstáculos", {
        "Fortalezas profesionales": "Principales fortalezas profesionales",
        "Obstáculos para el desarrollo": "Obstáculos para el desarrollo profesional"
    })

    agregar_seccion("6. Proyección y Asesoramiento", {
        "¿Desea recibir asesoramiento?": "¿Desea recibir asesoramiento sobre su desarrollo profesional?",
        "¿Dispuesto a nuevos desafíos?": "¿Está dispuesto a asumir nuevos desafíos/responsabilidades?",
        "Comentarios adicionales": "Comentarios adicionales sobre su desarrollo"
    })

    # --- SECCIONES PARA COMPLETAR MANUALMENTE ---
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "7. Síntesis de la entrevista (Para completar por RRHH)", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.multi_cell(0, 8, "Percepción del entrevistado:\n\nExpectativas:\n\nPotencial detectado:\n", border=1)
    pdf.ln(5)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "8. Plan de Acción", ln=True)
    pdf.set_font("Arial", 'B', 9)
    # Encabezados de la tabla
    col_widths = [45, 45, 25, 25, 25, 25]
    headers = ['Objetivo', 'Acción', 'Responsable', 'Inicio', 'Revisión', 'Estado']
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 7, header, 1, 0, 'C')
    pdf.ln()
    # Filas vacías de la tabla
    pdf.set_font("Arial", '', 9)
    for _ in range(3):
        for width in col_widths:
            pdf.cell(width, 10, '', 1, 0, 'C')
        pdf.ln()

    # --- FINALIZACIÓN ---
    return pdf.output(dest='S').encode('latin-1')

# --- INTERFAZ DE STREAMLIT ---
url_google_sheet = st.text_input("Ingresa la URL de tu Hoja de Cálculo de Google:", key="url_input")

if url_google_sheet:
    client = conectar_google_sheets()
    if client:
        try:
            # REEMPLAZA LA LÍNEA ANTERIOR POR ESTA
            nombre_de_la_hoja = "Respuestas de formulario 1"  # <-- ¡CAMBIA ESTO!
            sheet = client.open_by_url(url_google_sheet).worksheet(nombre_de_la_hoja)
            datos = sheet.get_all_records()
            df = pd.DataFrame(datos)

            st.success("¡Conexión exitosa! ✅")
            
            if "Apellido y nombre" in df.columns:
                empleados = df["Apellido y nombre"].dropna().unique()
                empleado_seleccionado = st.selectbox("Selecciona un empleado para generar su PDI:", empleados)

                if empleado_seleccionado:
                    datos_empleado = df[df["Apellido y nombre"] == empleado_seleccionado].iloc[0].to_dict()
                    
                    st.write(f"### Previsualización de datos para: {empleado_seleccionado}")
                    st.json(datos_empleado, expanded=False)

                    if st.button(f"Generar PDF para {empleado_seleccionado}"):
                        pdf_bytes = generar_pdf(datos_empleado)
                        st.download_button(
                            label="📥 Descargar PDF",
                            data=pdf_bytes,
                            file_name=f"PDI_{empleado_seleccionado.replace(' ', '_')}.pdf",
                            mime="application/octet-stream"
                        )
            else:
                st.error("Error: La hoja de cálculo debe tener una columna llamada 'Apellido y nombre'.")
                st.write("Columnas encontradas:", df.columns.tolist())

        except gspread.exceptions.SpreadsheetNotFound:
            st.error("No se encontró la hoja de cálculo. Verifica la URL y asegúrate de haberla compartido con el correo de la cuenta de servicio.")
        except Exception as e:

            st.error(f"Ocurrió un error inesperado: {e}")
