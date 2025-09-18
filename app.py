import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- DEFINICIÓN DE LA PALETA DE COLORES ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)
COLOR_GRIS_LINEA = (220, 220, 220)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)

# --- CLASE MEJORADA PARA CREAR EL PDF EJECUTIVO ---
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = "Resumen de Dotación"
        self.table_header_data = None 

    def header(self):
        self.set_font("Arial", "B", 18)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, self.report_title, 0, 0, "C")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def draw_section_title(self, title):
        self.set_font("Arial", "B", 14)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, title, ln=True, align="L")
        self.set_draw_color(*COLOR_AZUL_INSTITUCIONAL)
        self.set_line_width(0.5)
        self.line(self.get_x(), self.get_y(), self.get_x() + self.page_width, self.get_y())
        self.ln(5)

    def draw_kpi_box(self, title, value, color, x, y, width=80):
        kpi_height = 16
        self.set_xy(x, y)
        
        self.set_fill_color(*color)
        self.cell(width, 1.5, "", fill=True, ln=False, border=0)
        
        self.set_xy(x, y + 1.5)
        self.set_fill_color(255, 255, 255)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.cell(width, kpi_height - 1.5, "", border=1, fill=True)
        
        self.set_xy(x, y + 3)
        self.set_font('Arial', '', 10)
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.cell(width, 8, title, align='C')
        
        self.set_xy(x, y + 8)
        self.set_font('Arial', 'B', 16)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(width, 10, str(value), align='C')

    def draw_table(self, title, df_original, is_crosstab=False):
        if df_original.empty or (is_crosstab and len(df_original) <= 1 and not (len(df_original) == 1 and df_original.index[0] != "Total")):
             return
        
        df = df_original.copy()
        if is_crosstab: df = df.replace(0, '-')
        if df.index.name: df.reset_index(inplace=True)

        if self.get_y() + (8 * (len(df) + 1) + 10) > self.h - self.b_margin: self.add_page(orientation=self.cur_orientation)

        self.draw_section_title(title)

        df_formatted = df.copy()
        for col in df_formatted.columns:
             if pd.api.types.is_numeric_dtype(df_formatted[col]) and col not in ['Nº pers.', 'Antigüedad']:
                  df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) else x)

        widths = {col: max(self.get_string_width(str(col)) + 10, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x)).max() + 10) for col in df_formatted.columns}
        total_width = sum(widths.values())
        if total_width > self.page_width:
            scaling_factor = self.page_width / total_width
            widths = {k: v * scaling_factor for k, v in widths.items()}

        self.set_font("Arial", "B", 9)
        self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
        self.set_text_color(255, 255, 255)
        for col in df_formatted.columns:
            self.cell(widths[col], 8, str(col), 0, 0, "C", True)
        self.ln()

        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.set_line_width(0.2)
        
        for i, (_, row) in enumerate(df_formatted.iterrows()):
            if self.get_y() + 8 > self.h - self.b_margin:
                self.add_page(orientation=self.cur_orientation)
                self.set_font("Arial", "B", 9)
                self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
                self.set_text_color(255, 255, 255)
                for col in df_formatted.columns:
                    self.cell(widths[col], 8, str(col), 0, 0, "C", True)
                self.ln()
                self.set_text_color(*COLOR_TEXTO_CUERPO)

            fill = i % 2 == 1
            if "Total" in str(row.iloc[0]):
                self.set_font("Arial", "B", 9)
                fill = False
            else:
                self.set_font("Arial", "", 9)
            
            self.set_fill_color(*COLOR_GRIS_FONDO_FILA)
            
            for col in df_formatted.columns:
                self.cell(widths[col], 8, str(row[col]), 'T', 0, "C", fill)
            self.ln()
        
        self.ln(10)

def crear_pdf_reporte(titulo_reporte, rango_fechas_str, df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos, df_desaparecidos=None):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = titulo_reporte
    pdf.add_page()
    
    pdf.draw_section_title(f"Indicadores del Período: {rango_fechas_str}")
    total_activos_val = f"{resumen_activos.loc['Total', 'Total']:,}".replace(',', '.') if not resumen_activos.empty else "0"
    
    y = pdf.get_y()

    if df_desaparecidos is not None and not df_desaparecidos.empty:
        kpi_width = 65; spacing = (pdf.page_width - (kpi_width * 4)) / 3
        x1 = pdf.l_margin; x2 = x1 + kpi_width + spacing; x3 = x2 + kpi_width + spacing; x4 = x3 + kpi_width + spacing
        pdf.draw_kpi_box("Dotación Activa", total_activos_val, (200, 200, 200), x1, y, width=kpi_width)
        pdf.draw_kpi_box("Altas del Período", str(len(df_altas)), (200, 200, 200), x2, y, width=kpi_width)
        pdf.draw_kpi_box("Bajas del Período", str(len(df_bajas)), (200, 200, 200), x3, y, width=kpi_width)
        pdf.draw_kpi_box("Cambio Organizacional", str(len(df_desaparecidos)), (255, 165, 0), x4, y, width=kpi_width)
    else:
        kpi_width = 80
        x1 = pdf.l_margin; x2 = x1 + kpi_width + 10; x3 = x2 + kpi_width + 10
        pdf.draw_kpi_box("Dotación Activa", total_activos_val, (200, 200, 200), x1, y, width=kpi_width)
        pdf.draw_kpi_box("Altas del Período", str(len(df_altas)), (200, 200, 200), x2, y, width=kpi_width)
        pdf.draw_kpi_box("Bajas del Período", str(len(df_bajas)), (200, 200, 200), x3, y, width=kpi_width)
    
    pdf.ln(22)
    
    fecha_final = rango_fechas_str.split(' - ')[-1]
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", resumen_bajas, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", resumen_altas, is_crosstab=True)
    pdf.draw_table(f"Composición de la Dotación Activa (Al {fecha_final})", resumen_activos, is_crosstab=True)

    if not df_altas.empty or not df_bajas.empty or not bajas_por_motivo.empty or (df_desaparecidos is not None and not df_desaparecidos.empty):
        if not df_altas.empty: pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
        if not df_bajas.empty: pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
        if not bajas_por_motivo.empty: pdf.draw_table("Bajas por Motivo", bajas_por_motivo)
        if df_desaparecidos is not None and not df_desaparecidos.empty:
            pdf.draw_table("Cambios Organizacionales (Legajos no encontrados)", df_desaparecidos[['Nº pers.']])

    return bytes(pdf.output())
    
def procesar_archivo_base(archivo_cargado, sheet_name='BaseQuery'):
    df_base = pd.read_excel(archivo_cargado, sheet_name=sheet_name, engine='openpyxl')
    df_base.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}, inplace=True)
    for col in ['Fecha', 'Desde', 'Fecha nac.']:
        if col in df_base.columns: df_base[col] = pd.to_datetime(df_base[col], errors='coerce')
    
    orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
    orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
    df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
    df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)
    return df_base

def formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw, df_desaparecidos_raw=None):
    df_bajas = df_bajas_raw.copy()
    if not df_bajas.empty:
        # Intentar calcular antigüedad solo si la columna 'Fecha' existe
        if 'Fecha' in df_bajas.columns:
            df_bajas['Antigüedad'] = ((datetime.now() - df_bajas['Fecha']) / pd.Timedelta(days=365.25)).fillna(0).astype(int)
        else:
            df_bajas['Antigüedad'] = 0 # Valor por defecto si no hay fecha de alta
        df_bajas['Fecha nac.'] = df_bajas['Fecha nac.'].dt.strftime('%d/%m/%Y')
        df_bajas['Desde'] = df_bajas['Desde'].dt.strftime('%d/%m/%Y')
    else:
        df_bajas = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría'])
    
    df_altas = df_altas_raw.copy()
    if not df_altas.empty:
        df_altas['Fecha'] = df_altas['Fecha'].dt.strftime('%d/%m/%Y')
        df_altas['Fecha nac.'] = df_altas['Fecha nac.'].dt.strftime('%d/%m/%Y')
    else:
        df_altas = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría'])
    
    df_desaparecidos = df_desaparecidos_raw.copy() if df_desaparecidos_raw is not None else pd.DataFrame(columns=['Nº pers.'])
    if not df_desaparecidos.empty:
        if 'Fecha' in df_desaparecidos.columns:
             df_desaparecidos['Antigüedad'] = ((datetime.now() - df_desaparecidos['Fecha']) / pd.Timedelta(days=365.25)).fillna(0).astype(int)
        else:
             df_desaparecidos['Antigüedad'] = 0
    else:
        df_desaparecidos = pd.DataFrame(columns=['Nº pers.'])

    return df_altas, df_bajas, df_desaparecidos

def filtrar_novedades_por_fecha(df_base_para_filtrar, fecha_inicio, fecha_fin):
    df = df_base_para_filtrar.copy()
    altas_filtradas = df[(df['Fecha'] >= fecha_inicio) & (df['Fecha'] <= fecha_fin)].copy()
    df_bajas_potenciales = df[df['Status ocupación'] == 'Dado de baja'].copy()
    if not df_bajas_potenciales.empty:
        df_bajas_potenciales['fecha_baja_corregida'] = df_bajas_potenciales['Desde'] - pd.Timedelta(days=1)
        bajas_filtradas = df_bajas_potenciales[(df_bajas_potenciales['fecha_baja_corregida'] >= fecha_inicio) & (df_bajas_potenciales['fecha_baja_corregida'] <= fecha_fin)].copy()
        if not bajas_filtradas.empty:
            bajas_filtradas['Desde'] = bajas_filtradas['fecha_baja_corregida']
    else:
        bajas_filtradas = pd.DataFrame()
    return altas_filtradas, bajas_filtradas

def calcular_activos_a_fecha(df_base, fecha_fin):
    df = df_base.copy()
    df = df[df['Fecha'] <= fecha_fin]
    
    df_bajas = df[df['Status ocupación'] == 'Dado de baja'].copy()
    if not df_bajas.empty:
        df_bajas['fecha_baja_corregida'] = df_bajas['Desde'] - pd.Timedelta(days=1)
        legajos_baja_despues_de_fecha = df_bajas[df_bajas['fecha_baja_corregida'] > fecha_fin]['Nº pers.']
    else:
        legajos_baja_despues_de_fecha = []

    activos_en_fecha = df[
        (df['Status ocupación'] == 'Activo') | 
        (df['Nº pers.'].isin(legajos_baja_despues_de_fecha))
    ]
    return activos_en_fecha

# --- INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")
st.markdown("""<style>.main .block-container { padding-top: 2rem; padding-bottom: 2rem; background-color: #f0f2f6; } h1, h2, h3 { color: #003366; } div.stDownloadButton > button { background-color: #28a745; color: white; border-radius: 5px; font-weight: bold; }</style>""", unsafe_allow_html=True)
st.title("📊 Dashboard de Control de Dotación")

tab1, tab2, tab3, tab4 = st.tabs(["▶️ Novedades (General)", "📈 Resúmenes (General)", "📅 Reporte Semanal", "📅 Reporte Mensual"])

with tab1:
    st.header("Análisis General por Comparación de Archivos")
    st.info("Sube tu archivo Excel con las pestañas 'BaseQuery' y 'Activos' para ver las novedades generales.")
    uploaded_file_general = st.file_uploader("Sube tu archivo Excel aquí", type=['xlsx'], key="main_uploader")

    if uploaded_file_general:
        try:
            st.session_state.uploaded_file_general = uploaded_file_general
            df_base_general = procesar_archivo_base(uploaded_file_general, sheet_name='BaseQuery')
            df_activos_general_raw = pd.read_excel(uploaded_file_general, sheet_name='Activos')
            st.session_state.df_base_general = df_base_general
            st.session_state.df_activos_general_raw = df_activos_general_raw
            st.success("Archivo general cargado y procesado.")

            activos_legajos_viejos = set(df_activos_general_raw['Nº pers.'])
            df_bajas_general_raw = df_base_general[df_base_general['Nº pers.'].isin(activos_legajos_viejos) & (df_base_general['Status ocupación'] == 'Dado de baja')].copy()
            df_altas_general_raw = df_base_general[~df_base_general['Nº pers.'].isin(activos_legajos_viejos) & (df_base_general['Status ocupación'] == 'Activo')].copy()
            
            todos_legajos_nuevos = set(df_base_general['Nº pers.'])
            legajos_desaparecidos = activos_legajos_viejos - todos_legajos_nuevos
            df_desaparecidos_raw = df_activos_general_raw[df_activos_general_raw['Nº pers.'].isin(legajos_desaparecidos)].copy()
            
            if not df_bajas_general_raw.empty: df_bajas_general_raw['Desde'] = df_bajas_general_raw['Desde'] - pd.Timedelta(days=1)
            
            df_altas_general, df_bajas_general, df_desaparecidos = formatear_y_procesar_novedades(df_altas_general_raw, df_bajas_general_raw, df_desaparecidos_raw)
            st.session_state.df_altas_general, st.session_state.df_bajas_general, st.session_state.df_desaparecidos = df_altas_general, df_bajas_general, df_desaparecidos
            
            resumen_activos_full = pd.crosstab(df_base_general[df_base_general['Status ocupación'] == 'Activo']['Categoría'], df_base_general[df_base_general['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
            resumen_bajas_full = pd.crosstab(df_bajas_general_raw['Categoría'], df_bajas_general_raw['Línea'], margins=True, margins_name="Total")
            resumen_altas_full = pd.crosstab(df_altas_general_raw['Categoría'], df_altas_general_raw['Línea'], margins=True, margins_name="Total")
            bajas_por_motivo_full = df_bajas_general_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
            if not bajas_por_motivo_full.empty: bajas_por_motivo_full.loc['Total'] = bajas_por_motivo_full.sum()

            pdf_bytes_general = crear_pdf_reporte("Resumen de Dotación", datetime.now().strftime('%d/%m/%Y'), df_altas_general, df_bajas_general, bajas_por_motivo_full.reset_index(), resumen_altas_full, resumen_bajas_full, resumen_activos_full, df_desaparecidos=df_desaparecidos)
            st.download_button(label="📄 Descargar Reporte General (PDF)", data=pdf_bytes_general, file_name=f"Reporte_General_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf")
            st.markdown("---")

            st.subheader(f"Altas ({len(df_altas_general)})"); st.dataframe(df_altas_general, hide_index=True)
            st.subheader(f"Bajas ({len(df_bajas_general)})"); st.dataframe(df_bajas_general, hide_index=True)
            if not df_desaparecidos.empty:
                st.subheader(f"Cambios Organizacionales ({len(df_desaparecidos)})")
                st.dataframe(df_desaparecidos, hide_index=True)

        except Exception as e:
            st.error(f"Ocurrió un error en el archivo general: {e}")
            st.warning("Verifica que el archivo contenga las pestañas 'Activos' y 'BaseQuery'.")

with tab2:
    st.header("Dashboard de Resúmenes (General)")
    if 'df_base_general' in st.session_state:
        df_base_general = st.session_state.df_base_general; df_activos_general_raw = st.session_state.df_activos_general_raw
        df_altas_general = st.session_state.df_altas_general; df_bajas_general = st.session_state.df_bajas_general
        df_desaparecidos = st.session_state.df_desaparecidos
        
        activos_legajos = set(df_activos_general_raw['Nº pers.'])
        df_bajas_general_raw = df_base_general[df_base_general['Nº pers.'].isin(activos_legajos) & (df_base_general['Status ocupación'] == 'Dado de baja')].copy()
        df_altas_general_raw = df_base_general[~df_base_general['Nº pers.'].isin(activos_legajos) & (df_base_general['Status ocupación'] == 'Activo')].copy()

        resumen_activos_full = pd.crosstab(df_base_general[df_base_general['Status ocupación'] == 'Activo']['Categoría'], df_base_general[df_base_general['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
        resumen_bajas_full = pd.crosstab(df_bajas_general_raw['Categoría'], df_bajas_general_raw['Línea'], margins=True, margins_name="Total")
        resumen_altas_full = pd.crosstab(df_altas_general_raw['Categoría'], df_altas_general_raw['Línea'], margins=True, margins_name="Total")
        bajas_por_motivo_full = df_bajas_general_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
        if not bajas_por_motivo_full.empty: bajas_por_motivo_full.loc['Total'] = bajas_por_motivo_full.sum()

        st.subheader("Indicadores Principales")
        kpi_cols = st.columns(4 if not df_desaparecidos.empty else 3)
        kpi_cols[0].metric("Dotación Activa", f"{resumen_activos_full.loc['Total', 'Total']:,}".replace(',', '.'))
        kpi_cols[1].metric("Altas del Período", len(df_altas_general))
        kpi_cols[2].metric("Bajas del Período", len(df_bajas_general))
        if not df_desaparecidos.empty:
            kpi_cols[3].metric("Cambio Organizacional", len(df_desaparecidos))

        st.markdown("---")
        
        formatter = lambda x: f'{x:,.0f}'.replace(',', '.') if isinstance(x, (int, float)) else x
        st.subheader("Composición de la Dotación Activa"); st.dataframe(resumen_activos_full.replace(0, '-').style.format(formatter))
        st.subheader("Resumen de Novedades")
        colA, colB = st.columns(2)
        with colA: st.write("**Bajas por Categoría y Línea:**"); st.dataframe(resumen_bajas_full.replace(0, '-').style.format(formatter))
        with colB: st.write("**Altas por Categoría y Línea:**"); st.dataframe(resumen_altas_full.replace(0, '-').style.format(formatter))
        st.write("**Bajas por Motivo:**"); st.dataframe(bajas_por_motivo_full.style.format(formatter))
    else:
        st.info("Sube un archivo en la pestaña 'Novedades (General)' para ver los resúmenes.")

def run_period_report(report_type, df_base_main, df_activos_main):
    st.header(f"Generador de Reportes {report_type}s")
    st.info(f"Para el reporte {report_type}, sube un archivo con las pestañas 'BaseQuery' y 'Activos' que representen el inicio y fin del período.")
    
    uploader_period = st.file_uploader(f"Sube un archivo para el reporte {report_type}", type=['xlsx'], key=f"upload_{report_type}")
    archivo_para_periodo = uploader_period or st.session_state.get('uploaded_file_general')

    if archivo_para_periodo:
        try:
            df_base_period = procesar_archivo_base(archivo_para_periodo, sheet_name='BaseQuery')
            df_activos_period_raw = pd.read_excel(archivo_para_periodo, sheet_name='Activos')
            
            today = datetime.now()
            if report_type == 'Semanal':
                dflt_start = today - timedelta(days=7); dflt_end = today
            else: # Mensual
                dflt_start = today.replace(day=1); dflt_end = (dflt_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

            col1, col2 = st.columns(2)
            start_date = col1.date_input("Fecha de inicio", dflt_start, key=f"{report_type}_inicio")
            end_date = col2.date_input("Fecha de fin", dflt_end, key=f"{report_type}_fin")

            if start_date and end_date and start_date <= end_date:
                rango_str = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
                st.write(f"**Período a analizar:** {rango_str}")

                # Lógica Híbrida
                df_altas_raw, df_bajas_raw = filtrar_novedades_por_fecha(df_base_period, pd.to_datetime(start_date), pd.to_datetime(end_date))
                
                activos_viejos = set(df_activos_period_raw['Nº pers.'])
                todos_nuevos = set(df_base_period['Nº pers.'])
                legajos_desaparecidos = activos_viejos - todos_nuevos
                df_desaparecidos_raw = df_activos_period_raw[df_activos_period_raw['Nº pers.'].isin(legajos_desaparecidos)].copy()

                df_altas, df_bajas, df_desaparecidos = formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw, df_desaparecidos_raw)
                
                df_activos_periodo = calcular_activos_a_fecha(df_base_period, pd.to_datetime(end_date))
                resumen_activos = pd.crosstab(df_activos_periodo['Categoría'], df_activos_periodo['Línea'], margins=True, margins_name="Total")
                resumen_bajas = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")
                resumen_altas = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
                bajas_motivo = df_bajas_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
                if not bajas_motivo.empty: bajas_motivo.loc['Total'] = bajas_motivo.sum()

                pdf_bytes = crear_pdf_reporte(f"Resumen {report_type} de Dotación", rango_str, df_altas, df_bajas, bajas_motivo.reset_index(), resumen_altas, resumen_bajas, resumen_activos, df_desaparecidos)
                st.download_button(f"📄 Descargar Reporte {report_type} en PDF", pdf_bytes, f"Reporte_{report_type}_{start_date.strftime('%Y%m')}.pdf", "application/pdf", key=f"btn_{report_type}")
            elif start_date > end_date:
                st.error("La fecha de inicio no puede ser posterior a la fecha de fin.")
        except Exception as e:
            st.error(f"Ocurrió un error en el archivo para el reporte {report_type}: {e}")
            st.warning("Verifica que el archivo contenga las pestañas 'Activos' y 'BaseQuery'.")
    else:
        st.info("Sube un archivo en la pestaña 'Novedades (General)' o aquí mismo para generar un reporte.")

with tab3:
    run_period_report('Semanal', st.session_state.get('df_base_general'), st.session_state.get('df_activos_general_raw'))
    
with tab4:
    run_period_report('Mensual', st.session_state.get('df_base_general'), st.session_state.get('df_activos_general_raw'))
