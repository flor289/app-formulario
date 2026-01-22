import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)
COLOR_GRIS_LINEA = (220, 220, 220)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = "Resumen de Dotación"

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

    def draw_kpi_box(self, title, value, color, x, y, width=65):
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
        if df_original.empty: return
        df = df_original.copy()
        if is_crosstab: 
            df = df.replace(0, '-')
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
            self.set_font("Arial", "B" if "Total" in str(row.iloc[0]) else "", 9)
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
    
    # KPIs dinámicos
    has_co = df_desaparecidos is not None and not df_desaparecidos.empty
    kpi_width = 65 if has_co else 80
    spacing = (pdf.page_width - (kpi_width * (4 if has_co else 3))) / 3
    
    pdf.draw_kpi_box("Dotación Activa", total_activos_val, (200, 200, 200), pdf.l_margin, y, width=kpi_width)
    pdf.draw_kpi_box("Altas del Período", '-' if len(df_altas) == 0 else str(len(df_altas)), (200, 200, 200), pdf.l_margin + kpi_width + spacing, y, width=kpi_width)
    pdf.draw_kpi_box("Bajas del Período", '-' if len(df_bajas) == 0 else str(len(df_bajas)), (200, 200, 200), pdf.l_margin + (kpi_width + spacing)*2, y, width=kpi_width)
    
    if has_co:
        pdf.draw_kpi_box("Cambio Organizativo", str(len(df_desaparecidos)), (255, 165, 0), pdf.l_margin + (kpi_width + spacing)*3, y, width=kpi_width)
    
    pdf.ln(22)
    fecha_final = rango_fechas_str.split(' - ')[-1]
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", resumen_bajas, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", resumen_altas, is_crosstab=True)
    pdf.draw_table(f"Composición de la Dotación Activa (Al {fecha_final})", resumen_activos, is_crosstab=True)
    
    if not df_altas.empty: pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    if not df_bajas.empty: pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not bajas_por_motivo.empty: pdf.draw_table("Bajas por Motivo", bajas_por_motivo)
    
    if has_co:
        cols_presentes = [c for c in ['Nº pers.', 'Apellido', 'Nombre de pila', 'Desde', 'Línea', 'Categoría'] if c in df_desaparecidos.columns]
        pdf.draw_table("Detalle Cambios Organizativos", df_desaparecidos[cols_presentes])
        
    return bytes(pdf.output())

def procesar_archivo_base(archivo_cargado, sheet_name='BaseQuery'):
    try:
        df = pd.read_excel(archivo_cargado, sheet_name=sheet_name, engine='openpyxl')
        df.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}, inplace=True)
        for col in ['Fecha', 'Desde', 'Fecha nac.']:
            if col in df.columns: df[col] = pd.to_datetime(df[col], errors='coerce')
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
        df['Línea'] = pd.Categorical(df['Línea'], categories=orden_lineas, ordered=True)
        df['Categoría'] = pd.Categorical(df['Categoría'], categories=orden_categorias, ordered=True)
        return df
    except: return pd.DataFrame()

def detectar_y_completar_co(legajos_desaparecidos, df_co_respaldo):
    """Cruza legajos desaparecidos con la pestaña CO."""
    df_desaparecidos = pd.DataFrame({'Nº pers.': list(legajos_desaparecidos)})
    
    if not df_co_respaldo.empty:
        df_desaparecidos = pd.merge(df_desaparecidos, df_co_respaldo, on='Nº pers.', how='left')
        sin_datos = df_desaparecidos[df_desaparecidos['Apellido'].isna()]['Nº pers.'].tolist()
        if sin_datos:
            st.warning(f"⚠️ **Aviso de App:** Se detectaron Cambios Organizativos pero **faltan datos en la pestaña 'CO'** para los legajos: {sin_datos}")
            
        for col in ['Desde', 'Fecha nac.']:
            if col in df_desaparecidos.columns:
                df_desaparecidos[col] = df_desaparecidos[col].dt.strftime('%d/%m/%Y').fillna('-')
    else:
        if legajos_desaparecidos:
            st.warning(f"⚠️ **Aviso de App:** Hay Cambios Organizativos pero la pestaña 'CO' no fue encontrada.")
            
    return df_desaparecidos

def filtrar_novedades_por_fecha(df_base_para_filtrar, fecha_inicio, fecha_fin):
    df = df_base_para_filtrar.copy()
    altas_filtradas = df[(df['Fecha'] >= fecha_inicio) & (df['Fecha'] <= fecha_fin)].copy()
    df_bajas_potenciales = df[df['Status ocupación'] == 'Dado de baja'].copy()
    if not df_bajas_potenciales.empty:
        df_bajas_potenciales['fecha_baja_corregida'] = df_bajas_potenciales['Desde'] - pd.Timedelta(days=1)
        bajas_filtradas = df_bajas_potenciales[(df_bajas_potenciales['fecha_baja_corregida'] >= fecha_inicio) & (df_bajas_potenciales['fecha_baja_corregida'] <= fecha_fin)].copy()
        if not bajas_filtradas.empty: bajas_filtradas['Desde'] = bajas_filtradas['fecha_baja_corregida']
    else:
        bajas_filtradas = pd.DataFrame()
    return altas_filtradas, bajas_filtradas

# --- INTERFAZ ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")
st.title("📊 Dashboard de Control de Dotación")

tabs = st.tabs(["📅 Reporte Diario", "📈 Resúmenes", "📅 Semanal", "📅 Mensual", "📅 Anual"])

with tabs[0]:
    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx'])
    if uploaded_file:
        try:
            df_base = procesar_archivo_base(uploaded_file, 'BaseQuery')
            df_activos_prev = pd.read_excel(uploaded_file, sheet_name='Activos')
            df_co_respaldo = procesar_archivo_base(uploaded_file, 'CO') # Corregido a CO
            
            legajos_viejos = set(df_activos_prev['Nº pers.'])
            desaparecidos = legajos_viejos - set(df_base['Nº pers.'])
            
            df_altas_raw = df_base[~df_base['Nº pers.'].isin(legajos_viejos) & (df_base['Status ocupación'] == 'Activo')].copy()
            df_bajas_raw = df_base[df_base['Nº pers.'].isin(legajos_viejos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
            if not df_bajas_raw.empty: df_bajas_raw['Desde'] = df_bajas_raw['Desde'] - pd.Timedelta(days=1)
            
            df_desaparecidos = detectar_y_completar_co(desaparecidos, df_co_respaldo)
            
            df_altas_vis = df_altas_raw.copy()
            if not df_altas_vis.empty:
                df_altas_vis['Fecha'] = df_altas_vis['Fecha'].dt.strftime('%d/%m/%Y')
                df_altas_vis['Fecha nac.'] = df_altas_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')
            
            resumen_activos = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
            resumen_bajas = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")
            resumen_altas = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
            bajas_motivo = df_bajas_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
            
            pdf_bytes = crear_pdf_reporte("Reporte Diario", datetime.now().strftime('%d/%m/%Y'), df_altas_vis, df_bajas_raw, bajas_motivo.reset_index(), resumen_altas, resumen_bajas, resumen_activos, df_desaparecidos)
            
            st.download_button("📄 Descargar Reporte Diario", pdf_bytes, f"Reporte_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf")
            
            st.subheader(f"Altas ({len(df_altas_vis)})"); st.dataframe(df_altas_vis, hide_index=True)
            st.subheader(f"Detalle Cambios Organizativos ({len(desaparecidos)})"); st.dataframe(df_desaparecidos, hide_index=True)
            
            st.session_state.df_base = df_base
            st.session_state.df_activos_prev = df_activos_prev
            st.session_state.df_co_respaldo = df_co_respaldo
        except Exception as e: st.error(f"Error: {e}")

def render_periodico(tipo):
    if 'df_base' not in st.session_state:
        st.info("Sube un archivo en 'Reporte Diario' primero.")
        return
    
    col1, col2 = st.columns(2)
    inicio = col1.date_input(f"Inicio {tipo}", key=f"ini_{tipo}")
    fin = col2.date_input(f"Fin {tipo}", key=f"fin_{tipo}")
    
    if inicio <= fin:
        df_base = st.session_state.df_base
        df_co_respaldo = st.session_state.df_co_respaldo
        df_activos_prev = st.session_state.df_activos_prev
        
        df_altas_raw, df_bajas_raw = filtrar_novedades_por_fecha(df_base, pd.to_datetime(inicio), pd.to_datetime(fin))
        
        # Lógica Anual Normalización
        if tipo == "Anual" and not df_altas_raw.empty:
            df_altas_raw['Categoría'] = 'ASP.AY.C'
            st.info(f"💡 Info de App: Se normalizaron las altas a 'ASP.AY.C' para este reporte anual.")
            
        desaparecidos = set(df_activos_prev['Nº pers.']) - set(df_base['Nº pers.'])
        df_desaparecidos = detectar_y_completar_co(desaparecidos, df_co_respaldo)
        
        resumen_activos = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
        resumen_bajas = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")
        resumen_altas = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
        
        pdf_bytes = crear_pdf_reporte(f"Reporte {tipo}", f"{inicio.strftime('%d/%m/%Y')} - {fin.strftime('%d/%m/%Y')}", df_altas_raw, df_bajas_raw, pd.DataFrame(), resumen_altas, resumen_bajas, resumen_activos, df_desaparecidos)
        st.download_button(f"📄 Descargar Reporte {tipo}", pdf_bytes, f"Reporte_{tipo}.pdf")

with tabs[2]: render_periodico("Semanal")
with tabs[3]: render_periodico("Mensual")
with tabs[4]: render_periodico("Anual")
