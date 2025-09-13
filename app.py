import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile  # Importamos la librería para crear archivos ZIP

# Importamos las librerías de reportlab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

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

        # --- GENERACIÓN DE PDF (Sin cambios) ---
        def generar_pdf(datos_empleado):
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=inch)
            styles = getSampleStyleSheet()
            color_azul = colors.HexColor("#2a5caa")
            color_gris = colors.HexColor("#808080")
            styles.add(ParagraphStyle(name='TituloPrincipal', parent=styles['h1'], textColor=color_azul, alignment=TA_CENTER, fontSize=18))
            styles.add(ParagraphStyle(name='TituloSeccion', parent=styles['h2'], textColor=color_azul, spaceAfter=6))
            styles.add(ParagraphStyle(name='NormalJustificado', parent=styles['Normal'], alignment=TA_JUSTIFY))
            styles.add(ParagraphStyle(name='Etiqueta', parent=styles['Normal'], fontName='Helvetica-Bold'))
            story = []
            p_titulo = Paragraph("PLAN DE DESARROLLO INDIVIDUAL (PDI)", styles['TituloPrincipal'])
            story.append(p_titulo)
            story.append(Spacer(1, 24))
            def crear_checkbox(pregunta, opciones, respuesta):
                marcado, no_marcado = "☒", "☐"
                texto_final = f"<b>{pregunta}:</b><br/>"
                lineas = [f"{marcado} <b>{op}</b>" if str(respuesta).strip().lower() == op.strip().lower() else f"<font color='{color_gris}'>{no_marcado} {op}</font>" for op in opciones]
                texto_final += " &nbsp; ".join(lineas)
                return Paragraph(texto_final, styles['Normal'])
            def format_as_list(text):
                if isinstance(text, str) and ',' in text:
                    items = [item.strip() for item in text.split(',')]
                    return "<br/>".join(f"- {item}" for item in items)
                return text
            def agregar_seccion(titulo, campos):
                bloque_seccion = [Paragraph(titulo, styles['TituloSeccion']), Spacer(1, 6)]
                for etiqueta, config in campos.items():
                    columna = config['col']
                    valor = str(datos_empleado.get(columna, 'N/A'))
                    if config.get('type') == 'checkbox':
                        bloque_seccion.append(crear_checkbox(etiqueta, config['options'], valor))
                    elif config.get('type') == 'list':
                        p_etiqueta = Paragraph(f"<b>{etiqueta}:</b>", styles['Etiqueta'])
                        p_valor = Paragraph(format_as_list(valor), styles['NormalJustificado'])
                        bloque_seccion.extend([p_etiqueta, p_valor])
                    else:
                        p_etiqueta = Paragraph(f"<b>{etiqueta}:</b>", styles['Etiqueta'])
                        p_valor = Paragraph(valor, styles['NormalJustificado'])
                        bloque_seccion.extend([p_etiqueta, p_valor])
                    bloque_seccion.append(Spacer(1, 10))
                story.append(KeepTogether(bloque_seccion))
            
            # --- Contenido del PDI ---
            agregar_seccion("1. Datos Personales y Laborales", {"Apellido y Nombre": {'col': "Apellido y Nombre"}, "DNI": {'col': "DNI"}, "Correo electrónico": {'col': "Correo electrónico"}, "Número de contacto": {'col': "Número de contacto"}, "Edad": {'col': "Edad"}, "Posición actual": {'col': "Posición actual"}, "Fecha de ingreso": {'col': "Fecha de ingreso a la empresa"}, "Lugar de trabajo": {'col': "Lugar de trabajo"}})
            agregar_seccion("2. Formación y Nivel Educativo", {"Nivel educativo": {'col': "Nivel educativo alcanzado"}, "Título obtenido": {'col': "Título obtenido (si corresponde)"}, "Otras capacitaciones": {'col': "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)"}, "Relación entre puesto actual y formación académica": {'col': "Su puesto actual ¿está relacionado con su formación académica?", 'type': 'checkbox', 'options': ["Totalmente", "Parcialmente", "No"]}})
            agregar_seccion("3. Interés de Desarrollo", {"¿Le interesaría desarrollar su carrera dentro de la empresa?": {'col': "¿Le interesaría desarrollar su carrera dentro de la empresa?", 'type': 'checkbox', 'options': ["Sí", "No"]}, "Área de interés futura": {'col': "¿En qué área de la empresa le gustaría desarrollarse en el futuro?", 'type': 'list'}, "Puesto al que aspira": {'col': "¿Qué tipo de puesto aspira ocupar en el futuro?"}, "Motivaciones para cambiar": {'col': "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición dentro de la empresa? (Seleccione hasta 3 opciones)", 'type': 'list'}})
            agregar_seccion("4. Necesidades de Capacitación", {"Competencias a capacitar": {'col': "¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo?", 'type': 'list'}, "Especificación de interés": {'col': "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse"}})
            agregar_seccion("5. Fortalezas y Obstáculos", {"Fortalezas profesionales": {'col': "¿Cuáles considera que son sus principales fortalezas profesionales?", 'type': 'list'}, "Obstáculos para el desarrollo": {'col': "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?", 'type': 'list'}})
            agregar_seccion("6. Proyección y Crecimiento", {"¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional?": {'col': "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa?", 'type': 'checkbox', 'options': ["Sí", "No"]}, "¿Estaría dispuesto a asumir nuevos desafíos/responsabilidades?": {'col': "¿Estaría dispuesto a asumir nuevos desafíos/responsabilidades para avanzar en su carrera dentro de la empresa?", 'type': 'checkbox', 'options': ["Sí", "No", "No lo sé"]}, "Comentarios adicionales": {'col': "Si desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"}})
            
            story.append(PageBreak())
            story.append(Paragraph("7. Síntesis de la entrevista", styles['TituloSeccion']))
            story.append(Paragraph("(Para completar por el responsable de RRHH o desarrollo)", styles['Italic']))
            story.append(Spacer(1, 24))
            story.append(Paragraph("<b>Percepción del entrevistado:</b>", styles['Normal']))
            story.append(Spacer(1, 48))
            story.append(Paragraph("<b>Expectativas y motivaciones:</b>", styles['Normal']))
            story.append(Spacer(1, 48))
            story.append(Paragraph("<b>Potencial detectado:</b>", styles['Normal']))
            story.append(Spacer(1, 48))
            story.append(Spacer(1, 24))
            story.append(Paragraph("8. Plan de Acción", styles['TituloSeccion']))
            story.append(Spacer(1, 12))
            header_style = ParagraphStyle(name='HeaderStyle', parent=styles['Normal'], fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=TA_CENTER)
            headers = [Paragraph(h, header_style) for h in ["Objetivo de Desarrollo", "Acción a realizar", "Responsable", "Fecha de inicio", "Fecha de revisión", "Estado"]]
            data_tabla = [headers]
            for i in range(4): data_tabla.append(["", "", "", "", "", ""])
            tabla = Table(data_tabla, colWidths=[1.5*inch, 1.5*inch, 1*inch, 1*inch, 1*inch, 1*inch])
            tabla.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), color_azul), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,0), 12), ('TOPPADDING', (0,0), (-1,0), 6), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#f0f0f0")), ('GRID', (0,0), (-1,-1), 1, colors.black), ('ROWHEIGHTS', (1, -1), [30] * 4)]))
            story.append(tabla)
            
            doc.build(story)
            buffer.seek(0)
            return buffer

        # --- INTERFAZ PRINCIPAL ---
        columna_nombre = "Apellido y Nombre"
        if columna_nombre in df.columns:
            
            # --- SECCIÓN PARA GENERAR PDF INDIVIDUAL ---
            st.header("Generar PDF Individual")
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
            
            st.divider() # Línea divisoria

            # --- NUEVA SECCIÓN PARA GENERAR TODOS LOS PDFS EN UN ZIP ---
            st.header("Generar Todos los Formularios en un ZIP")
            if st.button("🚀 Generar y Descargar ZIP con Todos los PDI"):
                zip_buffer = BytesIO()
                
                # Barra de progreso
                progress_bar = st.progress(0, text="Iniciando generación de PDFs...")
                total_empleados = len(df)

                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for index, row in df.iterrows():
                        datos_empleado = row.to_dict()
                        nombre_empleado_raw = datos_empleado.get(columna_nombre, f"Empleado_{index+1}")
                        
                        # Genera el PDF para el empleado actual
                        pdf_buffer = generar_pdf(datos_empleado)
                        
                        # Crea un nombre de archivo seguro
                        nombre_archivo = f"PDI_{nombre_empleado_raw.replace(' ', '_').replace(',', '')}.pdf"
                        
                        # Añade el PDF al archivo ZIP
                        zipf.writestr(nombre_archivo, pdf_buffer.getvalue())

                        # Actualiza la barra de progreso
                        progreso_actual = (index + 1) / total_empleados
                        progress_bar.progress(progreso_actual, text=f"Generando PDF para: {nombre_empleado_raw} ({index+1}/{total_empleados})")

                progress_bar.empty() # Limpia la barra de progreso al finalizar
                
                st.download_button(
                    label="📥 Descargar Archivo ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="Todos_los_PDI.zip",
                    mime="application/zip"
                )

        else:
            st.error(f"Error: No se encontró la columna '{columna_nombre}' en tu archivo Excel.")
            st.write("Columnas encontradas:", df.columns.tolist())

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo Excel: {e}")
        st.error("Sugerencia: Revisa que los nombres de las columnas en tu Excel coincidan con los del código.")

