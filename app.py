import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- 1. DEFINICIÓN DE LA PALETA DE COLORES ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)
COLOR_GRIS_LINEA = (220, 220, 220)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)

# --- 2. CLASE PDF ---
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

# --- 3. FUNCIONES DE PROCESAMIENTO ---
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

def formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw, df_co_raw=None):
    # Procesar Bajas y calcular Antigüedad
    df_bajas = df_bajas_raw.copy()
    if not df_bajas.empty:
        df_bajas['Antigüedad'] = ((datetime.now() - df_bajas['Fecha']).dt.days / 365.25).fillna(0).astype(int)
        df_bajas['Fecha nac.'] = df_bajas['Fecha nac.'].dt.strftime('%d/%m/%Y')
        df_bajas['Desde'] = df_bajas['Desde'].dt.strftime('%d/%m/%Y')
    
    # Procesar Altas
    df_altas = df_altas_raw.copy()
    if not df_altas.empty:
        df_altas['Fecha'] = df_altas['Fecha'].dt.strftime('%d/%m/%Y')
        df_altas['Fecha nac.'] = df_altas['Fecha nac.'].dt.strftime('%d/%m/%Y')
    
    # Procesar CO
    df_co = df_co_raw.copy() if df_co_raw is not None else pd.DataFrame()
    if not df_co.empty and 'Desde' in df_co.columns:
        if pd.api.types.is_datetime64_any_dtype(df_co['Desde']):
            df_co['Desde'] = df_co['Desde'].dt.strftime('%d/%m/%Y')
            
    return df_altas, df_bajas, df_co

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

def crear_pdf_reporte(titulo_reporte, rango_fechas_str, df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos, df_desaparecidos=None):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = titulo_reporte
    pdf.add_page()
    pdf.draw_section_title(f"Indicadores del Período: {rango_fechas_str}")
    total_activos_val = f"{resumen_activos.loc['Total', 'Total']:,}".replace(',', '.') if not resumen_activos.empty else "0"
    
    # Lógica de globos KPIs
    has_co = df_desaparecidos is not None and not df_desaparecidos.empty
    kpi_width = 65 if has_co else 80
    spacing = (pdf.page_width - (kpi_width * (4 if has_co else 3))) / 3
    
    y = pdf.get_y()
    pdf.draw_kpi_box("Dotación Activa", total_activos_val, (200, 200, 200), pdf.l_margin, y, width=kpi_width)
    pdf.draw_kpi_box("Altas del Período", '-' if len(df_altas) == 0 else str(len(df_altas)), (200, 200, 200), pdf.l_margin + kpi_width + spacing, y, width=kpi_width)
    pdf.draw_kpi_box("Bajas del Período", '-' if len(df_bajas) == 0 else str(len(df_bajas)), (200, 200, 200), pdf.l_margin + (kpi_width + spacing)*2, y, width=kpi_width)
    if has_co:
        pdf.draw_kpi_box("Cambio Organizativo", str(len(df_desaparecidos)), (255, 165, 0), pdf.l_margin + (kpi_width + spacing)*3, y, width=kpi_width)
    
    pdf.ln(22)
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", resumen_bajas, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", resumen_altas, is_crosstab=True)
    pdf.draw_table(f"Composición de la Dotación Activa", resumen_activos, is_crosstab=True)
    
    if not df_altas.empty: pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    if not df_bajas.empty: pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not bajas_por_motivo.empty: pdf.draw_table("Bajas por Motivo", bajas_por_motivo)
    if has_co:
        pdf.draw_table("Detalle Cambios Organizativos", df_desaparecidos[['Nº pers.', 'Apellido', 'Nombre de pila', 'Desde', 'Línea', 'Categoría']])
        
    return bytes(pdf.output())

# --- 4. INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")
st.title("📊 Dashboard de Control de Dotación")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📅 Reporte Diario", "📈 Resúmenes (General)", "📅 Reporte Semanal", "📅 Reporte Mensual", "📅 Reporte Anual"])

with tab1:
    st.header("Análisis Diario por Comparación de Archivos")
    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx'], key="up_main")
    if uploaded_file:
        try:
            df_base = procesar_archivo_base(uploaded_file, 'BaseQuery')
            df_activos_prev = pd.read_excel(uploaded_file, sheet_name='Activos')
            try: df_co_respaldo = procesar_archivo_base(uploaded_file, 'CO')
            except: df_co_respaldo = pd.DataFrame()

            # Detección física
            viejos = set(df_activos_prev['Nº pers.'])
            desaparecidos = viejos - set(df_base['Nº pers.'])
            
            # Cruce con CO para Diario
            if not df_co_respaldo.empty:
                df_co_diario = df_co_respaldo[df_co_respaldo['Nº pers.'].isin(desaparecidos)].copy()
                faltantes = desaparecidos - set(df_co_diario['Nº pers.'])
                if faltantes: st.warning(f"⚠️ C.O. detectados sin datos en pestaña 'CO': {faltantes}")
            else:
                df_co_diario = df_activos_prev[df_activos_prev['Nº pers.'].isin(desaparecidos)].copy()
                if desaparecidos: st.warning("⚠️ Pestaña 'CO' no encontrada.")

            df_altas_raw = df_base[~df_base['Nº pers.'].isin(viejos) & (df_base['Status ocupación'] == 'Activo')].copy()
            df_bajas_raw = df_base[df_base['Nº pers.'].isin(viejos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
            if not df_bajas_raw.empty: df_bajas_raw['Desde'] = df_bajas_raw['Desde'] - pd.Timedelta(days=1)

            df_altas, df_bajas, df_co = formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw, df_co_diario)
            
            # Tablas cruzadas
            res_act = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
            res_alt = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
            res_baj = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")
            b_motivo = df_bajas_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')

            pdf = crear_pdf_reporte("Resumen Diario de Dotación", datetime.now().strftime('%d/%m/%Y'), df_altas, df_bajas, b_motivo.reset_index(), res_alt, res_baj, res_act, df_co)
            
            # NOMBRE ARCHIVO DIARIO OK
            n_diario = f"Reporte_Diario_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf"
            st.download_button("📄 Descargar Reporte Diario", pdf, n_diario, "application/pdf")

            st.session_state.uploaded_file = uploaded_file
            st.session_state.df_base = df_base
            st.session_state.df_activos_prev = df_activos_prev
            st.session_state.df_co_respaldo = df_co_respaldo
        except Exception as e: st.error(f"Error: {e}")

with tab2:
    st.header("Dashboard de Resúmenes (General)")
    if 'df_base' in st.session_state:
        df_base = st.session_state.df_base
        res = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
        st.metric("Dotación Activa Actual", f"{res.loc['Total', 'Total']:,}".replace(',', '.'))
    else: st.info("Sube un archivo primero.")

def run_period_report(report_type):
    st.header(f"Generador de Reportes {report_type}es")
    archivo = st.session_state.get('uploaded_file')
    if archivo:
        df_base = st.session_state.df_base
        df_activos_prev = st.session_state.df_activos_prev
        df_co_respaldo = st.session_state.df_co_respaldo
        
        today = datetime.now()
        if report_type == 'Semanal': d_s = today - timedelta(days=7); d_e = today
        elif report_type == 'Mensual': d_s = today.replace(day=1); d_e = (d_s + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        else: d_s = today.replace(month=1, day=1); d_e = today.replace(month=12, day=31)

        c1, c2 = st.columns(2)
        start = c1.date_input("Inicio", d_s, key=f"s_{report_type}")
        end = c2.date_input("Fin", d_e, key=f"e_{report_type}")
        
        if start <= end:
            df_altas_raw, df_bajas_raw = filtrar_novedades_por_fecha(df_base, pd.to_datetime(start), pd.to_datetime(end))
            
            # Normalización Anual ASP.AY.C
            if report_type == 'Anual' and not df_altas_raw.empty:
                df_altas_raw['Categoría'] = 'ASP.AY.C'
                st.info("💡 Altas normalizadas a 'ASP.AY.C' para el reporte anual.")

            # FILTRADO CO POR FECHA CORRECTO
            viejos = set(df_activos_prev['Nº pers.'])
            desaparecidos = viejos - set(df_base['Nº pers.'])
            if not df_co_respaldo.empty:
                df_co_filtrado = df_co_respaldo[(df_co_respaldo['Nº pers.'].isin(desaparecidos)) & 
                                               (df_co_respaldo['Desde'] >= pd.to_datetime(start)) & 
                                               (df_co_respaldo['Desde'] <= pd.to_datetime(end))].copy()
            else:
                df_co_filtrado = pd.DataFrame()

            df_altas, df_bajas, df_co = formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw, df_co_filtrado)
            
            # Activos a la fecha fin
            df_act_per = df_base[df_base['Fecha'] <= pd.to_datetime(end)]
            res_act = pd.crosstab(df_act_per[df_act_per['Status ocupación'] == 'Activo']['Categoría'], df_act_per[df_act_per['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
            res_alt = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
            res_baj = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")

            # TÍTULOS DE REPORTE RESTAURADOS
            titulo_pdf = f"Reporte {report_type} de Dotación"
            pdf = crear_pdf_reporte(titulo_pdf, f"{start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}", df_altas, df_bajas, pd.DataFrame(), res_alt, res_baj, res_act, df_co)
            
            # NOMBRES DE ARCHIVO RESTAURADOS
            if report_type == 'Mensual': n_file = f"Reporte_Mensual_{start.strftime('%Y%m')}.pdf"
            elif report_type == 'Anual': n_file = f"Reporte_Anual_{start.strftime('%Y')}.pdf"
            else: n_file = f"Reporte_Semanal_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"

            st.download_button(f"📄 Descargar {titulo_pdf}", pdf, n_file, "application/pdf")

with tab3: run_period_report('Semanal')
with tab4: run_period_report('Mensual')
with tab5: run_period_report('Anual')
