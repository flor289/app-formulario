import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Diagnóstico de Columnas", page_icon="ախ", layout="centered")
st.title("ախ Herramienta de Diagnóstico Final")
st.write("Sube tu archivo Excel para ver exactamente cómo Python está leyendo los nombres de las columnas.")

# --- CARGADOR DE ARCHIVO EXCEL ---
uploaded_file = st.file_uploader(
    "Sube tu archivo Excel para iniciar el diagnóstico",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("¡Archivo Excel cargado! Estos son los nombres de columna que se encontraron:")

        # --- MUESTRA DE NOMBRES DE COLUMNA ---
        # Mostramos la lista de nombres tal cual los lee Python
        st.subheader("Columnas encontradas en tu archivo:")
        
        columnas_encontradas = df.columns.tolist()
        st.write(columnas_encontradas)
        
        st.divider()

        # --- VERIFICACIÓN ---
        st.subheader("Verificación")
        columna_buscada = "Apellido y Nombre"
        st.write(f"El código está buscando exactamente este texto: `{columna_buscada}`")

        # Limpiamos los nombres de columna como lo haría el programa real
        df.columns = [col.strip() for col in df.columns]
        
        if columna_buscada in df.columns:
            st.success(f"¡Éxito! Después de limpiar espacios, la columna '{columna_buscada}' fue encontrada.")
            st.info("Esto significa que el problema era un espacio en blanco al inicio o al final del nombre en Excel.")
        else:
            st.error(f"¡Fallo! La columna '{columna_buscada}' NO se encuentra en la lista de arriba, incluso después de limpiar espacios.")
            st.warning("Esto indica que hay una diferencia (un acento, una letra, etc.) entre el nombre en el Excel y el que busca el código.")
            st.info("Por favor, copia el nombre correcto de la lista de arriba y envíamelo.")

    except Exception as e:
        st.error(f"Ocurrió un error al leer el archivo Excel: {e}")
