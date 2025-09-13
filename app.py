import streamlit as st
import pandas as pd
from io import BytesIO

# Importamos las librerías de reportlab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas # <-- ¡ESTA ES LA LÍNEA QUE FALTABA!

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

        # --- GENERACIÓN DE PDF (VERSIÓN FINAL CON ESTÉTICA MEJORADA) ---
        def generar_pdf(datos_empleado):
            buffer = BytesIO()
            
            # --- CLASE PARA AÑADIR PIE DE PÁGINA ---
            class PieDePaginaCanvas(canvas.Canvas):
                def __init__(self, *args, **kwargs):
                    canvas.Canvas.__init__(self, *args, **kwargs)
                    self._saved_page_states = []

                def showPage(self):
                    self._saved_page_states.append(dict(self.__dict__))
                    self._startPage()

                def save(self):
                    num_pages = len(self._saved_page_states)
                    for state in self._saved_page_states:
                        self.__dict__.update(state)
                        self.draw_footer(num_pages)
                        canvas.Canvas.showPage(self)
                    canvas.Canvas.save(self)

                def draw_footer(self, page_count):
                    self.saveState()
                    self.setFont('Helvetica', 9)
                    self.drawString(inch, 0.75 * inch, f"Página {self._pageNumber} de {page_count}")
                    self.restoreState()

            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=inch)
            
            # --- ESTILOS PERSONALIZADOS ---
            styles = getSampleStyleSheet()
            color_azul = colors.HexColor("#2a5caa") # Un azul corporativo

            styles.add(ParagraphStyle(name='TituloPrincipal', parent=styles['h1'], textColor=color_azul, alignment=TA_CENTER))
            styles.add(ParagraphStyle(name='TituloSeccion', parent=styles['h2'], textColor=color_azul))
            styles.add(ParagraphStyle(name='CheckboxStyle', parent=styles['Normal'], alignment=TA_LEFT))
            
            story = []

            # --- Título ---
            p_titulo = Paragraph("PLAN DE DESARROLLO INDIVIDUAL (PDI)", styles['TituloPrincipal'])
            story.append(p_titulo)
            story.append(Spacer(1, 24))

            # --- Función para Checkboxes (sin cambios) ---
            def crear_checkbox(pregunta, opciones, respuesta):
                marcado, no_marcado = "☒", "☐"
                texto_final = f"<b>{pregunta}:</b><br/>"
                lineas = [f"{marcado} {op}" if str(respuesta).strip().lower() == op.strip().lower() else f"{no_marcado} {op}" for op in opciones]
                texto_final += " ".join(lineas)
                return Paragraph(texto_final, styles['CheckboxStyle'])

            # --- Función para Secciones (actualizada con nuevos estilos) ---
            def agregar_seccion(titulo, campos, es_checkbox=False, opciones_checkbox=None, pregunta_checkbox_actual=None):
                bloque_seccion = [Paragraph(titulo, styles['TituloSeccion']), Spacer(1, 12)]
                for etiqueta, columna in campos.items():
                    valor = str(datos_empleado.get(columna, 'N/A'))
                    if es_checkbox and etiqueta == pregunta_checkbox_actual:
                        bloque_seccion.append(crear_checkbox(etiqueta, opciones_checkbox, valor))
                    else:
                        bloque_seccion.extend([Paragraph(f"<b>{etiqueta}:</b>", styles['Normal']), Paragraph(valor, styles['Normal'])])
                    bloque_seccion.append(Spacer(1, 6))
                story.append(KeepTogether(bloque_seccion))
                story.append(Spacer(1, 18))

            # --- Contenido del PDI ---
            agregar_seccion("1. Datos Personales y Laborales", {"Apellido y Nombre": "Apellido y Nombre", "DNI": "DNI", "Correo electrónico": "Correo electrónico", "Número de contacto": "Número de contacto", "Edad": "Edad", "Posición actual": "Posición actual", "Fecha de ingreso": "Fecha de ingreso a la empresa", "Lugar de trabajo": "Lugar de trabajo"})
            agregar_seccion("2. Formación y Nivel Educativo", {"Nivel educativo": "Nivel educativo alcanzado", "Título obtenido": "Título obtenido (si corresponde)", "Otras capacitaciones": "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)", "Puesto relacionado con formación": "Su puesto actual ¿está relacionado con su formación académica?"})
            
            pregunta_interes = "¿Le interesaría desarrollar su carrera dentro de la empresa?"
            agregar_seccion("3. Interés de Desarrollo", {pregunta_interes: pregunta_interes, "Área de interés futura": "¿En qué área de la empresa le gustaría desarrollarse en el futuro?", "Puesto al que aspira": "¿Qué tipo de puesto aspira ocupar en el futuro?", "Motivaciones para cambiar": "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición dentro de la empresa? (Seleccione hasta 3 opciones)"}, es_checkbox=True, opciones_checkbox=["Sí", "No"], pregunta_checkbox_actual=pregunta_interes)
            
            agregar_seccion("4. Necesidades de Capacitación", {"Competencias a capacitar": "¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo?", "Especificación de interés": "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse"})
            agregar_seccion("5. Fortalezas y Obstáculos", {"Fortalezas profesionales": "¿Cuáles considera que son sus principales fortalezas profesionales?", "Obstáculos para el desarrollo": "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?"})
            
            pregunta_asesoramiento = "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa?"
            agregar_seccion("6. Proyección y Crecimiento", {pregunta_asesoramiento: pregunta_asesoramiento, "Dispuesto a nuevos desafíos": "¿Estaría dispuesto a asumir nuevos desafíos/responsabilidades?", "Comentarios adicionales": "Si desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"}, es_checkbox=True, opciones_checkbox=["Sí", "No"], pregunta_checkbox_actual=pregunta_asesoramiento)
            
            # --- SECCIONES FINALES MANUALES Y TABLA ---
            story.append(PageBreak())
            story.append(Paragraph("7. Síntesis de la entrevista", styles['TituloSeccion']))
            story.append(Paragraph("(Para completar por el responsable de RRHH o desarrollo)", styles['Italic']))
            story.append(Spacer(1, 12))
            story.append(Paragraph("Percepción del entrevistado:", styles['Normal']))
            story.append(Spacer(1, 48)) # Espacio para escribir
            story.append(Paragraph("Expectativas y motivaciones:", styles['Normal']))
            story.append(Spacer(1, 48))
            story.append(Paragraph("Potencial detectado:", styles['Normal']))
            story.append(Spacer(1, 48))
            
            story.append(Spacer(1, 24))
            story.append(Paragraph("8. Plan de Acción", styles['TituloSeccion']))
            story.append(Spacer(1, 12))

            # --- Creación de la Tabla "Plan de Acción" ---
            data_tabla = [["Objetivo de Desarrollo", "Acción a realizar", "Responsable", "Fecha de inicio", "Fecha de revisión", "Estado"]]
            for i in range(4): # 4 filas vacías
                data_tabla.append(["", "", "", "", "", ""])

            tabla = Table(data_tabla, colWidths=[1.5*inch, 1.5*inch, 1*inch, 1*inch, 1*inch, 1*inch])
            estilo_tabla = TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4a78c2")), # Fondo azul para el encabezado
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND', (0,1), (-1,-1), colors.beige),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('ROWHEIGHT', (1,1), (-1,-1), 30) # Alto de las filas de datos
            ])
            tabla.setStyle(estilo_tabla)
            story.append(tabla)
            
            doc.build(story, canvasmaker=PieDePaginaCanvas)
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
        st.error("Sugerencia: Revisa que los nombres de las columnas en tu Excel coincidan con los del código.")

