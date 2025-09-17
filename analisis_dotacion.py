import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- CLASE MEJORADA PARA CREAR EL PDF EJECUTIVO ---
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = "Resumen de Dotación"
        self.table_header_data = None 

    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, self.report_title, 0, 0, "C")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def _draw_table_header(self):
        if self.table_header_data:
            self.set_font("Arial", "B", self.table_header_data['font_size'])
            self.set_fill_color(70, 130, 180)
            self.set_text_color(255, 255, 255)
            for col in self.table_header_data['df_columns']:
                self.cell(self.table_header_data['widths'][col], 8, str(col), 1, 0, "C", True)
            self.ln()
            self.set_text_color(0, 0, 0)

    def draw_table(self, title, df_original, is_crosstab=False):
        if df_original.empty or (is_crosstab and len(df_original) <= 1 and not (len(df_original) == 1 and df_original.index[0] != "Total")):
             return
        
        df = df_original.copy()
        if is_crosstab:
            df = df.replace(0, '-')
        
        if df.index.name:
            df.reset_index(inplace=True)
        
        table_height = 8 * (len(df) + 1) + 10
        if self.get_y() + table_height > self.h - self.b_margin:
            self.add_page(orientation=self.cur_orientation)

        self.set_font("Arial", "B", 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, title, ln=True, align="L")
        self.ln(2)

        df_formatted = df.copy()
        for col in df_formatted.columns:
             if pd.api.types.is_numeric_dtype(df_formatted[col]) and col not in ['Nº pers.', 'Antigüedad']:
                  df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) else x)

        widths = {col: max(self.get_string_width(str(col)) + 8, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x)).max() + 8) for col in df_formatted.columns}
        total_width = sum(widths.values())
        
        font_size = 9
        if total_width > self.page_width:
            scaling_factor = self.page_width / total_width
            widths = {k: v * scaling_factor for k, v in widths.items()}
            font_size = 7

        self.table_header_data = { 'widths': widths, 'font_size': font_size, 'df_columns': df_formatted.columns }
        self._draw_table_header()
        
        for _, row in df_formatted.iterrows():
            if self.get_y() + 8 > self.h - self.b_margin:
                self.add_page(orientation=self.cur_orientation)
                self._draw_table_header()

            is_total_row = "Total" in str(row.iloc[0])
            if is_total_row: self.set_font("Arial", "B", font_size)
            else: self.set_font("Arial", "", font_size)

            for col in df_formatted.columns: self.cell(widths[col], 8, str(row[col]), 1, 0, "C")
            self.ln()
        
        self.table_header_data = None
        self.ln(10)

def crear_pdf_reporte(titulo_reporte, rango_fechas_str, df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = titulo_reporte
    pdf.add_page()
    
    fecha_final = rango_fechas_str.split(' - ')[-1]
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", resumen_bajas, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", resumen_altas, is_crosstab=True)
    pdf.draw_table(f"Composición de la Dotación Activa (Al {fecha_final})", resumen_activos, is_crosstab=True)

    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 10, f"Novedades del Período: {rango_fechas_str}", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"- Cantidad de Altas: {len(df_altas)}", ln=True)
    pdf.cell(0, 8, f"- Cantidad de Bajas: {len(df_bajas)}", ln=True)
    pdf.ln(5)

    if not df_altas.empty: pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    if not df_bajas.empty: pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not bajas_por_motivo.empty: pdf.draw_table("Bajas por Motivo", bajas_por_motivo)

    return bytes(pdf.output())

def procesar_archivo_base(archivo_cargado, sheet_name='BaseQuery'):
    """Función para leer y procesar un archivo base con un nombre de pestaña específico."""
    df_base = pd.read_excel(archivo_cargado, sheet_name=sheet_name, engine='openpyxl')
    df_base.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}, inplace=True)
    for col in ['Fecha', 'Desde', 'Fecha nac.']:
        if col in df_base.columns: df_base[col] = pd.to_datetime(df_base[col], errors='coerce')
    
    orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
    orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
    df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
    df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)
    return df_base

def formatear_y_procesar_novedades(df_altas_raw, df_bajas_raw):
    df_bajas = df_bajas_raw.copy()
    if not df_bajas.empty:
        df_bajas['Antigüedad'] = ((datetime.now() - df_bajas['Fecha']) / pd.Timedelta(days=365.25)).fillna(0).astype(int)
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
    return df_altas, df_bajas

def filtrar_novedades_por_fecha(df_base_para_filtrar, fecha_inicio, fecha_fin):
    df = df_base_para_filtrar.copy()
    altas_filtradas = df[(df['Fecha'] >= fecha_inicio) & (df['Fecha'] <= fecha_fin)].copy()

    df_bajas_potenciales = df[df['Status ocupación'] == 'Dado de baja'].copy()
    if not df_bajas_potenciales.empty:
        df_bajas_potenciales['fecha_baja_corregida'] = df_bajas_potenciales['Desde'] - pd.Timedelta(days=1)
        bajas_filtradas = df_bajas_potenciales[(df_bajas_potenciales['fecha_baja_corregida'] >= fecha_inicio) & (df_bajas_potenciales['fecha_baja_corregida'] <= fecha_fin)].copy()
    else:
        bajas_filtradas = pd.DataFrame()
    
    return altas_filtradas, bajas_filtradas

# --- INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")
st.markdown("""<style>.main .block-container { padding-top: 2rem; padding-bottom: 2rem; background-color: #f0f2f6; } h1, h2, h3 { color: #003366; } div.stDownloadButton > button { background-color: #28a745; color: white; border-radius: 5px; font-weight: bold; }</style>""", unsafe_allow_html=True)
st.title("📊 Dashboard de Control de Dotación")
st.write("Sube tus archivos para el reporte general o usa las pestañas para reportes por período.")

col1, col2 = st.columns(2)
with col1: uploaded_file = st.file_uploader("1. Sube tu archivo BaseQuery", type=['xlsx'], key="main_base")
with col2: uploaded_file_activos = st.file_uploader("2. Sube tu archivo de Activos anterior", type=['xlsx'], key="main_activos")

if uploaded_file and uploaded_file_activos:
    try:
        df_base = procesar_archivo_base(uploaded_file, sheet_name='BaseQuery')
        df_activos_raw = pd.read_excel(uploaded_file_activos, sheet_name='Activos', engine='openpyxl')
        
        st.success("¡Archivos generales cargados y procesados!")
        
        activos_legajos = set(df_activos_raw['Nº pers.'])
        df_bajas_general_raw = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        df_altas_general_raw = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()

        if not df_bajas_general_raw.empty: df_bajas_general_raw['Desde'] = df_bajas_general_raw['Desde'] - pd.Timedelta(days=1)
        df_altas_general, df_bajas_general = formatear_y_procesar_novedades(df_altas_general_raw, df_bajas_general_raw)
        
        resumen_activos_full = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
        resumen_bajas_full = pd.crosstab(df_bajas_general_raw['Categoría'], df_bajas_general_raw['Línea'], margins=True, margins_name="Total")
        resumen_altas_full = pd.crosstab(df_altas_general_raw['Categoría'], df_altas_general_raw['Línea'], margins=True, margins_name="Total")
        bajas_por_motivo_full = df_bajas_general_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
        if not bajas_por_motivo_full.empty: bajas_por_motivo_full.loc['Total'] = bajas_por_motivo_full.sum()
        
        pdf_bytes_general = crear_pdf_reporte("Resumen de Dotación", datetime.now().strftime('%d/%m/%Y'), df_altas_general, df_bajas_general, bajas_por_motivo_full.reset_index(), resumen_altas_full, resumen_bajas_full, resumen_activos_full)
        st.download_button(label="📄 Descargar Reporte General (PDF)", data=pdf_bytes_general, file_name=f"Reporte_General_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf")
        st.markdown("---")

        tab1, tab2, tab3, tab4 = st.tabs(["▶️ Novedades (General)", "📈 Resúmenes (General)", "📅 Reporte Semanal", "📅 Reporte Mensual"])
        formatter = lambda x: f'{x:,.0f}'.replace(',', '.') if isinstance(x, (int, float)) else x
        
        with tab1:
            st.header("Detalle de Novedades (por comparación de archivos)")
            st.subheader(f"Altas ({len(df_altas_general)})"); st.dataframe(df_altas_general[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']], hide_index=True)
            st.subheader(f"Bajas ({len(df_bajas_general)})"); st.dataframe(df_bajas_general[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']], hide_index=True)

        with tab2:
            st.header("Dashboard de Resúmenes (Completo)"); st.subheader("Composición de la Dotación Activa")
            st.dataframe(resumen_activos_full.replace(0, '-').style.format(formatter))
            st.subheader("Resumen de Novedades")
            col1, col2 = st.columns(2)
            with col1: st.write("**Bajas por Categoría y Línea:**"); st.dataframe(resumen_bajas_full.replace(0, '-').style.format(formatter))
            with col2: st.write("**Altas por Categoría y Línea:**"); st.dataframe(resumen_altas_full.replace(0, '-').style.format(formatter))
            st.write("**Bajas por Motivo:**"); st.dataframe(bajas_por_motivo_full.style.format(formatter))

        with tab3:
            st.header("Generador de Reportes Semanales (por fecha de evento)")
            uploader_sem = st.file_uploader("Opcional: Sube un archivo (con pestaña 'Sheet1')", type=['xlsx'], key="upload_sem")
            
            df_base_sem = procesar_archivo_base(uploader_sem, sheet_name='Sheet1') if uploader_sem else df_base
            
            start_date_sem = st.date_input("Fecha de inicio del reporte", datetime.now() - timedelta(days=7), key="semanal")
            if start_date_sem:
                end_date_sem = datetime.now()
                rango_str_sem = f"{start_date_sem.strftime('%d/%m/%Y')} - {end_date_sem.strftime('%d/%m/%Y')}"
                st.write(f"**Período a analizar:** {rango_str_sem}")

                df_altas_sem_raw, df_bajas_sem_raw = filtrar_novedades_por_fecha(df_base_sem, pd.to_datetime(start_date_sem), end_date_sem)
                df_altas_sem, df_bajas_sem = formatear_y_procesar_novedades(df_altas_sem_raw, df_bajas_sem_raw)
                
                resumen_bajas_sem = pd.crosstab(df_bajas_sem_raw['Categoría'], df_bajas_sem_raw['Línea'], margins=True, margins_name="Total")
                resumen_altas_sem = pd.crosstab(df_altas_sem_raw['Categoría'], df_altas_sem_raw['Línea'], margins=True, margins_name="Total")
                bajas_motivo_sem = df_bajas_sem_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
                if not bajas_motivo_sem.empty: bajas_motivo_sem.loc['Total'] = bajas_motivo_sem.sum()

                pdf_bytes_sem = crear_pdf_reporte("Resumen Semanal de Dotación", rango_str_sem, df_altas_sem, df_bajas_sem, bajas_motivo_sem.reset_index(), resumen_altas_sem, resumen_bajas_sem, resumen_activos_full)
                st.download_button("📄 Descargar Reporte Semanal en PDF", pdf_bytes_sem, f"Reporte_Semanal_{start_date_sem.strftime('%Y%m%d')}.pdf", "application/pdf", key="btn_sem")

        with tab4:
            st.header("Generador de Reportes Mensuales (por fecha de evento)")
            uploader_men = st.file_uploader("Opcional: Sube un archivo (con pestaña 'Sheet1')", type=['xlsx'], key="upload_men")
            
            df_base_men = procesar_archivo_base(uploader_men, sheet_name='Sheet1') if uploader_men else df_base

            today = datetime.now()
            dflt_start = today.replace(day=1); dflt_end = (dflt_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
            
            col1, col2 = st.columns(2)
            with col1: start_date_men = st.date_input("Fecha de inicio", dflt_start, key="mensual_inicio")
            with col2: end_date_men = st.date_input("Fecha de fin", dflt_end, key="mensual_fin")

            if start_date_men and end_date_men and start_date_men <= end_date_men:
                rango_str_men = f"{start_date_men.strftime('%d/%m/%Y')} - {end_date_men.strftime('%d/%m/%Y')}"
                st.write(f"**Período a analizar:** {rango_str_men}")

                df_altas_men_raw, df_bajas_men_raw = filtrar_novedades_por_fecha(df_base_men, pd.to_datetime(start_date_men), pd.to_datetime(end_date_men))
                df_altas_men, df_bajas_men = formatear_y_procesar_novedades(df_altas_men_raw, df_bajas_men_raw)
                
                resumen_bajas_men = pd.crosstab(df_bajas_men_raw['Categoría'], df_bajas_men_raw['Línea'], margins=True, margins_name="Total")
                resumen_altas_men = pd.crosstab(df_altas_men_raw['Categoría'], df_altas_men_raw['Línea'], margins=True, margins_name="Total")
                bajas_motivo_men = df_bajas_men_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
                if not bajas_motivo_men.empty: bajas_motivo_men.loc['Total'] = bajas_motivo_men.sum()

                pdf_bytes_men = crear_pdf_reporte("Resumen Mensual de Dotación", rango_str_men, df_altas_men, df_bajas_men, bajas_motivo_men.reset_index(), resumen_altas_men, resumen_bajas_men, resumen_activos_full)
                st.download_button("📄 Descargar Reporte Mensual en PDF", pdf_bytes_men, f"Reporte_Mensual_{start_date_men.strftime('%Y%m')}.pdf", "application/pdf", key="btn_men")
            elif start_date_men > end_date_men:
                st.error("La fecha de inicio no puede ser posterior a la fecha de fin.")
                
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.warning("Verifica que tus archivos Excel tengan el formato y las pestañas correctas ('Activos', 'BaseQuery' o 'Sheet1').")
