import streamlit as st
import pandas as pd
from io import BytesIO

# Importamos las librerías de reportlab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de PDI", page_icon="📄", layout="centered")
st.title("📄 Generador de Planes de Desarrollo Individual (PDI)")
st.write("Esta aplicación genera un PDI en PDF a partir de un archivo Excel que subas.")

# --- CARGADOR DE ARCHIVO EXCEL ---
uploaded_file = st.file_uploader(
    "Sube tu archivo Excel con los datos de los empleados",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("¡Archivo Excel cargado correctamente! ✅")

        # --- GENERACIÓN DE PDF (VERSIÓN FINAL CON REPORTLAB) ---
        def generar_pdf(datos_empleado):
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
            
            styles = getSampleStyleSheet()
            story = []

            # --- Título ---
            p_titulo = Paragraph("PLAN DE DESARROLLO INDIVIDUAL (PDI)", styles['h1'])
            story.append(p_titulo)
            story.append(Spacer(1, 24))

            # --- Función para agregar secciones ---
            def agregar_seccion(titulo, campos):
                story.append(Paragraph(titulo, styles['h2']))
                story.append(Spacer(1, 12))
                
                for etiqueta, columna in campos.items():
                    valor = str(datos_empleado.get(columna, 'N/A'))
                    # Usamos Paragraph para que el texto se ajuste automáticamente
                    p_etiqueta = Paragraph(f"<b>{etiqueta}:</b>", styles['Normal'])
                    p_valor = Paragraph(valor, styles['Normal'])
                    story.append(p_etiqueta)
                    story.append(p_valor)
                    story.append(Spacer(1, 6))
                story.append(Spacer(1, 18))

            # --- Contenido del PDI ---
            agregar_seccion("1. Datos Personales y Laborales", {"Apellido y Nombre": "Apellido y Nombre", "DNI": "DNI", "Correo electrónico": "Correo electrónico", "Número de contacto": "Número de contacto", "Edad": "Edad", "Posición actual": "Posición actual", "Fecha de ingreso": "Fecha de ingreso a la empresa", "Lugar de trabajo": "Lugar de trabajo"})
            agregar_seccion("2. Formación y Nivel Educativo", {"Nivel educativo": "Nivel educativo alcanzado", "Título obtenido": "Título obtenido (si corresponde)", "Otras capacitaciones": "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)", "Puesto relacionado con formación": "Su puesto actual ¿está relacionado con su formación académica?"})
            agregar_seccion("3. Interés de Desarrollo", {"Interesado en desarrollar carrera": "¿Le interesaría desarrollar su carrera dentro de la empresa?", "Área de interés futura": "¿En qué área de la empresa le gustaría desarrollarse en el futuro?", "Puesto al que aspira": "¿Qué tipo de puesto aspira ocupar en el futuro?", "Motivaciones para cambiar": "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición dentro de la empresa? (Seleccione hasta 3 opciones)"})
            agregar_seccion("4. Necesidades de Capacitación", {"Competencias a capacitar": "¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo?", "Especificación de interés": "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse"})
            agregar_seccion("5. Fortalezas y Obstáculos", {"Fortalezas profesionales": "¿Cuáles considera que son sus principales fortalezas profesionales?", "Obstáculos para el desarrollo": "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?"})
            agregar_seccion("6. Proyección y Crecimiento", {"Desea recibir asesoramiento": "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa?", "Dispuesto a nuevos desafíos": "¿Estaría dispuesto a asumir nuevos desafíos/responsabilidades?", "Comentarios adicionales": "Si desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"})

            doc.build(story)
            buffer.seek(0)
            return buffer

        # --- INTERFAZ PARA SELECCIONAR EMPLEADO ---
        columna_nombre = "Apellido y Nombre"
        if columna_nombre in df.columns:
            empleados = df[columna_nombre].dropna().unique()
            empleado_seleccionado = st.selectbox("Selecciona un empleado:", empleados)

            if empleado_seleccionado:
                datos_empleado = df[df[columna_nombre] == empleado_seleccionado].iloc[0].to_dict()
                if st.button(f"Generar PDF para {empleado_seleccionado}"):
                    pdf_buffer = generar_pdf(datos_empleado)
                    st.download_button(
                        label="📥 Descargar PDF",
                        data=pdf_buffer,
                        file_name=f"PDI_{empleado_seleccionado.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )
        else:
            st.error(f"Error: No se encontró la columna '{columna_nombre}' en tu archivo Excel.")
            st.write("Columnas encontradas:", df.columns.tolist())

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo Excel: {e}")
