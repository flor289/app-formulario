import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from fpdf import FPDF
import json

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de PDI", page_icon="📄", layout="centered")
st.title("📄 Generador de Planes de Desarrollo Individual (PDI)")
st.write("Esta aplicación genera un PDI en PDF a partir de los datos de una hoja de cálculo de Google.")

# --- CONEXIÓN SEGURA CON GOOGLE SHEETS (VERSIÓN ACTUALIZADA) ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scopes = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        
        # CAMBIO CLAVE: Leemos el JSON desde un string simple en los secretos
        if "google_creds_json" in st.secrets:
            creds_dict = json.loads(st.secrets["google_creds_json"])
        else:
            with open("credentials.json", "r") as f:
                creds_dict = json.load(f)

        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        st.error("No se encontró el archivo 'credentials.json'. Asegúrate de que esté en la misma carpeta que app.py.")
        return None
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        return None

# --- GENERACIÓN DE PDF (Sin cambios) ---
def generar_pdf(datos_empleado):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "PLAN DE DESARROLLO INDIVIDUAL (PDI)", ln=True, align='C')
    pdf.ln(10)
    def agregar_campo(etiqueta, valor_columna):
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(0, 6, f"{etiqueta}:", 0, 'L')
        pdf.set_font("Arial", '', 10)
        valor = str(datos_empleado.get(valor_columna, 'N/A'))
        pdf.multi_cell(0, 6, valor, 0, 'L')
        pdf.ln(2)
    def agregar_seccion(titulo, campos):
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, titulo, ln=True, align='L')
        for etiqueta, columna in campos.items():
            agregar_campo(etiqueta, columna)
        pdf.ln(5)
    agregar_seccion("1. Datos Personales y Laborales", {"Apellido y Nombre": "Apellido y Nombre", "DNI": "DNI", "Correo electrónico": "Correo electrónico", "Número de contacto": "Número de contacto", "Edad": "Edad", "Posición actual": "Posición actual", "Fecha de ingreso": "Fecha de ingreso a la empresa", "Lugar de trabajo": "Lugar de trabajo"})
    agregar_seccion("2. Formación y Nivel Educativo", {"Nivel educativo": "Nivel educativo alcanzado", "Título obtenido": "Título obtenido (si corresponde)", "Otras capacitaciones": "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)", "Puesto relacionado con formación": "Su puesto actual ¿está relacionado con su formación académica?"})
    agregar_seccion("3. Interés de Desarrollo", {"Interesado en desarrollar carrera": "¿Le interesaría desarrollar su carrera dentro de la empresa?", "Área de interés futura": "¿En qué área de la empresa le gustaría desarrollarse en el futuro?", "Puesto al que aspira": "¿Qué tipo de puesto aspira ocupar en el futuro?", "Motivaciones para cambiar": "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición dentro de la empresa? (Seleccione hasta 3 opciones)"})
    agregar_seccion("4. Necesidades de Capacitación", {"Competencias a capacitar": "¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo?", "Especificación de interés": "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse"})
    agregar_seccion("5. Fortalezas y Obstáculos", {"Fortalezas profesionales": "¿Cuáles considera que son sus principales fortalezas profesionales?", "Obstáculos para el desarrollo": "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?"})
    agregar_seccion("6. Proyección y Crecimiento", {"Desea recibir asesoramiento": "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa?", "Dispuesto a nuevos desafíos": "¿Estaría dispuesto a asumir nuevas responsabilidades o desafíos para avanzar en su carrera dentro de la empresa?", "Comentarios adicionales": "Si desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"})
    return pdf.output(dest='S').encode('latin-1')

# --- INTERFAZ DE STREAMLIT (Sin cambios) ---
url_google_sheet = st.text_input(
    "Ingresa la URL de tu Hoja de Cálculo de Google:",
    "https://docs.google.com/spreadsheets/d/1Uo5qFK34s94xGLO6WF3QuBGRLZi7XDYxcYVGZx79z_M/edit?usp=sharing"
)

if url_google_sheet:
    client = conectar_google_sheets()
    if client:
        try:
            spreadsheet = client.open_by_url(url_google_sheet)
            nombre_de_la_hoja = "Respuestas de formulario 1"
            sheet = spreadsheet.worksheet(nombre_de_la_hoja)
            all_values = sheet.get_all_values()
            
            if len(all_values) < 2:
                st.warning("La hoja de cálculo está vacía o solo tiene la fila de títulos.")
            else:
                df = pd.DataFrame(all_values[1:], columns=all_values[0])
                st.success("¡Datos cargados correctamente! ✅")

                columna_nombre = "Apellido y Nombre"
                if columna_nombre in df.columns:
                    empleados = df[columna_nombre].dropna().unique()
                    empleado_seleccionado = st.selectbox("Selecciona un empleado:", empleados)

                    if empleado_seleccionado:
                        datos_empleado = df[df[columna_nombre] == empleado_seleccionado].iloc[0].to_dict()
                        if st.button(f"Generar PDF para {empleado_seleccionado}"):
                            pdf_bytes = generar_pdf(datos_empleado)
                            st.download_button(
                                label="📥 Descargar PDF",
                                data=pdf_bytes,
                                file_name=f"PDI_{empleado_seleccionado.replace(' ', '_')}.pdf",
                                mime="application/octet-stream"
                            )
                else:
                    st.error(f"Error: No se encontró la columna '{columna_nombre}' en la hoja.")
                    st.write("Columnas encontradas:", df.columns.tolist())

        except gspread.exceptions.WorksheetNotFound:
            st.error(f"No se encontró la pestaña llamada '{nombre_de_la_hoja}'. Revisa el nombre en tu Google Sheet.")
        except Exception as e:
            st.error(f"Ocurrió un error inesperado al procesar la hoja: {e}")
