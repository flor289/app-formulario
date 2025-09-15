import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from fpdf import FPDF

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador PDI Estable", page_icon="✅", layout="centered")
st.title("✅ Generador de PDI (Versión Estable)")
st.write("Esta aplicación genera un PDI en PDF a partir de un archivo Excel que subas.")

# --- ESTRUCTURA DE DATOS (CON LOS NOMBRES 100% CORRECTOS DE TU EXCEL) ---
SECCIONES_PDI = {
    "1. Datos Personales y Laborales": {
        "Apellido y Nombre": "Apellido y Nombre", "DNI": "DNI", "Correo electrónico": "Correo electrónico",
        "Número de contacto": "Número de contacto", "Edad": "Edad", "Posición actual": "Posición actual",
        "Fecha de ingreso": "Fecha de ingreso a la empresa", "Lugar de trabajo": "Lugar de trabajo"
    },
    "2. Formación y Nivel Educativo": {
        "Nivel educativo": "Nivel educativo alcanzado", 
        "Carrera Cursada/En Curso": "Carrera cursada/en curso",
        "Título obtenido": "Título obtenido (si corresponde)",
        "Otras capacitaciones": "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)",
        "¿Está relacionado con su formación académica?": "¿está relacionado con su formación académica?"
    },
    "3. Interés de Desarrollo": {
        "¿Le interesaría desarrollar su carrera dentro de la empresa?": '¿Le interesaría desarrollar su carrera dentro de la empresa?',
        "Área de interés futura": "¿En qué área de la empresa le gustaría desarrollarse en el futuro?",
        "Puesto al que aspira": "¿Qué tipo de puesto aspira ocupar en el futuro?",
        "Motivaciones para cambiar": "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición  dentro de la empresa? (Seleccione hasta 3 opciones)"
    },
    "4. Necesidades de Capacitación": {
        "Competencias a capacitar": "¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo?",
        "Especificación de interés": "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse"
    },
    "5. Fortalezas y Obstáculos": {
        "Fortalezas profesionales": "¿Cuáles considera que son sus principales fortalezas profesionales?",
        "Obstáculos para el desarrollo": "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?"
    },
    "6. Proyección y Crecimiento": {
        "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional?": "¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa?",
        "¿Estaría dispuesto a asumir nuevos desafíos/responsabilidades?": "¿Estaría dispuesto a asumir nuevas responsabilidades o desafíos para avanzar en su carrera dentro de la empresa?",
        "Comentarios adicionales": "Si desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"
    }
}


# --- CARGADOR DE ARCHIVO EXCEL ---
uploaded_file = st.file_uploader(
    "Sube tu archivo Excel con los datos de los empleados",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]
        st.success("¡Archivo Excel cargado correctamente! ✅")

        # --- GENERACIÓN DE PDF con FPDF2 (Versión Estable) ---
        def generar_pdf(datos_empleado):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            
            # Título principal
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'PLAN DE DESARROLLO INDIVIDUAL (PDI)', 0, 1, 'C')
            pdf.ln(10)

            for titulo_seccion, campos in SECCIONES_PDI.items():
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, titulo_seccion, 0, 1, 'L')
                
                for etiqueta, columna in campos.items():
                    valor = str(datos_empleado.get(columna, 'N/A'))
                    
                    pdf.set_font('Arial', 'B', 10)
                    pdf.multi_cell(0, 6, etiqueta + ":")
                    
                    pdf.set_font('Arial', '', 10)
                    if "," in valor: # Formato de lista para respuestas múltiples
                        items = [item.strip() for item in valor.split(',')]
                        texto_lista = "\n".join(f"- {item}" for item in items)
                        pdf.multi_cell(0, 6, texto_lista)
                    else:
                        pdf.multi_cell(0, 6, valor)
                    pdf.ln(2)
                pdf.ln(5)

            return pdf.output(dest='S').encode('latin-1')

        # --- INTERFAZ PRINCIPAL ---
        columna_nombre = "Apellido y Nombre"
        if columna_nombre not in df.columns and "Nombre" in df.columns:
            columna_nombre = "Nombre" # Adaptación por si cambiaste el nombre
        
        if columna_nombre in df.columns:
            st.header("Generar PDF Individual")
            empleados = df[columna_nombre].dropna().unique()
            empleado_seleccionado = st.selectbox("Selecciona un empleado:", empleados)
            if empleado_seleccionado:
                datos_empleado = df[df[columna_nombre] == empleado_seleccionado].iloc[0].to_dict()
                if st.button(f"Generar PDF para {empleado_seleccionado}"):
                    pdf_buffer = generar_pdf(datos_empleado)
                    st.download_button(label="📥 Descargar PDF", data=pdf_buffer, file_name=f"PDI_{empleado_seleccionado.replace(' ', '_')}.pdf", mime="application/pdf")
            
            st.divider()
            
            st.header("Generar Todos los Formularios en un ZIP")
            if st.button("🚀 Generar y Descargar ZIP con Todos los PDI"):
                zip_buffer = BytesIO()
                progress_bar = st.progress(0, text="Iniciando generación de PDFs...")
                total_empleados = len(df)
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for index, row in df.iterrows():
                        nombre_empleado_raw = row.get(columna_nombre, f"Empleado_{index+1}")
                        pdf_buffer = generar_pdf(row.to_dict())
                        nombre_archivo = f"PDI_{str(nombre_empleado_raw).replace(' ', '_').replace(',', '')}.pdf"
                        zipf.writestr(nombre_archivo, pdf_buffer.getvalue())
                        progreso_actual = (index + 1) / total_empleados
                        progress_bar.progress(progreso_actual, text=f"Generando PDF: {nombre_empleado_raw} ({index+1}/{total_empleados})")
                progress_bar.empty()
                st.download_button(label="📥 Descargar Archivo ZIP", data=zip_buffer.getvalue(), file_name="Todos_los_PDI.zip", mime="application/zip")
        else:
            st.error(f"Error Crítico: No se encontró la columna '{columna_nombre}' o 'Nombre' en tu archivo Excel.")
            st.write("Columnas encontradas en tu archivo:")
            st.write(df.columns.tolist())

    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")


