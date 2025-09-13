import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Diagnóstico de Columnas", page_icon="ախ", layout="centered")
st.title("ախ Herramienta de Diagnóstico de Columnas")
st.write("Esta herramienta te ayudará a identificar inconsistencias entre tu archivo Excel y el código.")

# --- ESTRUCTURA DE DATOS (Columnas esperadas por el código) ---
# Esta es la lista de nombres de columna que el programa necesita encontrar
COLUMNAS_ESPERADAS = [
    "Apellido y Nombre", "DNI", "Correo electrónico", "Número de contacto", "Edad", "Posición actual", 
    "Fecha de ingreso a la empresa", "Lugar de trabajo", 
    "Nivel educativo alcanzado", "Título obtenido (si corresponde)", 
    "Otras capacitaciones realizadas fuera de la empresa finalizadas (Mencionar)",
    "Su puesto actual ¿está relacionado con su formación académica?",
    "Interés de desarrollo\n\n¿Le interesaría desarrollar su carrera dentro de la empresa?",
    "¿En qué área de la empresa le gustaría desarrollarse en el futuro?",
    "¿Cuáles son los principales factores que lo motivarían en su decisión de cambiar de posición  dentro de la empresa? (Seleccione hasta 3 opciones)",
    "¿Qué tipo de puesto aspira ocupar en el futuro?",
    "Capacitación y necesidades de aprendizaje\n¿En qué competencias o conocimientos le gustaría capacitarse para mejorar sus oportunidades de desarrollo? ",
    "A partir de su respuesta anterior, por favor, especifique en qué competencia o conocimiento le gustaría capacitarse",
    "Fortalezas y oportunidades de mejora\n¿Cuáles considera que son sus principales fortalezas profesionales? ",
    "¿Qué obstáculos encuentra para su desarrollo profesional dentro de la empresa?",
    "Proyección y crecimiento en la empresa\n¿Le gustaría recibir asesoramiento sobre su plan de desarrollo profesional dentro de la empresa? ",
    "¿Estaría dispuesto a asumir nuevas responsabilidades o desafíos para avanzar en su carrera dentro de la empresa? ",
    "Deje su comentario (opcional)\nSi desea agregar algún comentario sobre su desarrollo profesional en la empresa, puede hacerlo aquí:"
]

# --- CARGADOR DE ARCHIVO EXCEL ---
uploaded_file = st.file_uploader(
    "Sube tu archivo Excel para iniciar el diagnóstico",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("¡Archivo Excel cargado! Iniciando comparación...")
        st.divider()

        # --- COMPARACIÓN DE COLUMNAS ---
        columnas_excel = list(df.columns)
        
        # Creamos un DataFrame para la comparación visual
        max_len = max(len(COLUMNAS_ESPERADAS), len(columnas_excel))
        
        # Rellenamos las listas para que tengan la misma longitud
        esperadas_padded = COLUMNAS_ESPERADAS + [None] * (max_len - len(COLUMNAS_ESPERADAS))
        excel_padded = columnas_excel + [None] * (max_len - len(columnas_excel))
        
        df_comparacion = pd.DataFrame({
            'Nombres de Columna que el Código ESPERA': esperadas_padded,
            'Nombres de Columna que tu Excel TIENE': excel_padded
        })

        st.header("Resultado del Diagnóstico")
        st.write("Compara las dos columnas. Si un nombre no es **idéntico** (incluyendo espacios, acentos y saltos de línea), el programa fallará.")
        
        # Mostramos la tabla de comparación
        st.dataframe(df_comparacion, height=600)

        # Verificación final
        columnas_faltantes = [col for col in COLUMNAS_ESPERADAS if col.strip() not in [c.strip() for c in columnas_excel]]

        if not columnas_faltantes:
            st.success("¡PARECE QUE TODAS LAS COLUMNAS COINCIDEN! Si el error persiste, puede ser un problema de caché.")
        else:
            st.error("SE ENCONTRARON DIFERENCIAS. Las siguientes columnas esperadas no se encontraron:")
            st.dataframe(pd.DataFrame(columnas_faltantes, columns=["Columnas con problemas"]))
            st.warning("Por favor, copia la tabla de arriba o los nombres de la columna '...que tu Excel TIENE' y envíamelos para la corrección final.")

    except Exception as e:
        st.error(f"Ocurrió un error al leer el archivo Excel: {e}")

