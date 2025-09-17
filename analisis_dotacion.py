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
        # Propiedades para guardar el encabezado de la tabla
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
        """Función interna para dibujar el encabezado de una tabla."""
        if self.table_header_data:
            self.set_font("Arial", "B", self.table_header_data['font_size'])
            self.set_fill_color(70, 130, 180)
            self.set_text_color(255, 255, 255)
            for col in self.table_header_data['df_columns']:
                self.cell(self.table_header_data['widths'][col], 8, str(col), 1, 0, "C", True)
            self.ln()
            self.set_text_color(0, 0, 0)

    def draw_table(self, title, df_original, is_crosstab=False):
        if df_original.empty or (is_crosstab and len(df_original) <= 1):
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

        # Guardar datos del encabezado para poder redibujarlo
        self.table_header_data = {
            'widths': widths,
            'font_size': font_size,
            'df_columns': df_formatted.columns
        }
        
        self._draw_table_header() # Dibujar el primer encabezado
        
        for _, row in df_formatted.iterrows():
            # Revisar si la siguiente celda cabe en la página
            if self.get_y() + 8 > self.h - self.b_margin:
                self.add_page(orientation=self.cur_orientation)
                self._draw_table_header() # Redibujar encabezado en la nueva página

            is_total_row = "Total" in str(row.iloc[0])
            if is_total_row:
                self.set_font("Arial", "B", font_size)
            else:
                self.set_font("Arial", "", font_size)

            for col in df_formatted.columns:
                self.cell(widths[col], 8, str(row[col]), 1, 0, "C")
            self.ln()
        
        self.table_header_data = None # Limpiar datos del encabezado
        self.ln(10)
        
# ... (El resto del código es idéntico al anterior)

def crear_pdf_reporte(titulo_reporte, rango_fechas_str, df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = titulo_reporte # Establecer título dinámico

    # --- PÁGINA 1: RESÚMENES ---
    pdf.add_page()
    
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", resumen_bajas, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", resumen_altas, is_crosstab=True)
    
    pdf.draw_table(f"Composición de la Dotación Activa (Al {rango_fechas_str.split(' - ')[1]})", resumen_activos, is_crosstab=True)

    # --- PÁGINA 2: DETALLES ---
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 10, f"Novedades del Período: {rango_fechas_str}", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"- Cantidad de Altas: {len(df_altas)}", ln=True)
    pdf.cell(0, 8, f"- Cantidad de Bajas: {len(df_bajas)}", ln=True)
    pdf.ln(5)

    if not df_altas.empty:
        pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    if not df_bajas.empty:
        pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not bajas_por_motivo.empty:
        pdf.draw_table("Bajas por Motivo", bajas_por_motivo)

    return bytes(pdf.output())

# --- CONFIGURACIÓN E INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")

st.markdown("""
<style>
/* Estilos CSS para un look más profesional */
.main .block-container { padding-top: 2rem; padding-bottom: 2rem; background-color: #f0f2f6; }
h1, h2, h3 { color: #003366; }
div.stDownloadButton > button { background-color: #28a745; color: white; border-radius: 5px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Control de Dotación")
st.write("Sube tu archivo Excel para analizar las novedades y ver resúmenes.")

uploaded_file = st.file_uploader("Selecciona tu archivo Excel", type=['xlsx'], key="uploader")

if uploaded_file:
    try:
        df_base_raw = pd.read_excel(uploaded_file, sheet_name='BaseQuery', engine='openpyxl')
        df_activos_raw = pd.read_excel(uploaded_file, sheet_name='Activos', engine='openpyxl')
        
        df_base = df_base_raw.copy()
        df_base.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}, inplace=True)

        for col in ['Fecha', 'Desde', 'Fecha nac.']:
            if col in df_base.columns:
                df_base[col] = pd.to_datetime(df_base[col], errors='coerce')
        
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
        
        df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
        df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)

        activos_legajos = set(df_activos_raw['Nº pers.'])
        df_bajas_raw = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        
        if not df_bajas_raw.empty:
            df_bajas_raw['Desde'] = df_bajas_raw['Desde'] - pd.Timedelta(days=1)
            
        df_altas_raw = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()
        
        st.success("¡Archivo cargado y procesado!")
        st.markdown("---")

        tab1, tab2, tab3 = st.tabs(["▶️ Novedades (General)", "📈 Resúmenes (General)", "📅 Reporte Semanal"])
        
        df_bajas_full = df_bajas_raw.copy()
        if not df_bajas_full.empty:
            df_bajas_full['Antigüedad'] = ((datetime.now() - df_bajas_full['Fecha']) / pd.Timedelta(days=365.25)).fillna(0).astype(int)
            df_bajas_full['Fecha nac.'] = df_bajas_full['Fecha nac.'].dt.strftime('%d/%m/%Y')
            df_bajas_full['Desde'] = df_bajas_full['Desde'].dt.strftime('%d/%m/%Y')
        else:
            df_bajas_full = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría'])

        df_altas_full = df_altas_raw.copy()
        if not df_altas_full.empty:
            df_altas_full['Fecha'] = df_altas_full['Fecha'].dt.strftime('%d/%m/%Y')
            df_altas_full['Fecha nac.'] = df_altas_full['Fecha nac.'].dt.strftime('%d/%m/%Y')
        else:
            df_altas_full = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría'])
        
        formatter = lambda x: f'{x:,.0f}'.replace(',', '.') if isinstance(x, (int, float)) else x
        
        with tab1:
            st.header("Detalle de Todas las Novedades del Archivo")
            st.subheader(f"Altas ({len(df_altas_full)})")
            st.dataframe(df_altas_full[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']], hide_index=True)
            st.subheader(f"Bajas ({len(df_bajas_full)})")
            st.dataframe(df_bajas_full[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']], hide_index=True)

        with tab2:
            st.header("Dashboard de Resúmenes (Completo)")
            resumen_activos = pd.crosstab(df_base[df_base['Status ocupación'] == 'Activo']['Categoría'], df_base[df_base['Status ocupación'] == 'Activo']['Línea'], margins=True, margins_name="Total")
            resumen_bajas = pd.crosstab(df_bajas_raw['Categoría'], df_bajas_raw['Línea'], margins=True, margins_name="Total")
            resumen_altas = pd.crosstab(df_altas_raw['Categoría'], df_altas_raw['Línea'], margins=True, margins_name="Total")
            bajas_por_motivo = df_bajas_raw['Motivo de la medida'].value_counts().to_frame('Cantidad')
            if not bajas_por_motivo.empty:
                bajas_por_motivo.loc['Total'] = bajas_por_motivo.sum()
            bajas_por_motivo.index.name = "Motivo"

            st.subheader("Composición de la Dotación Activa")
            st.dataframe(resumen_activos.replace(0, '-').style.format(formatter))
            st.subheader("Resumen de Novedades")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Bajas por Categoría y Línea:**")
                st.dataframe(resumen_bajas.replace(0, '-').style.format(formatter))
            with col2:
                st.write("**Altas por Categoría y Línea:**")
                st.dataframe(resumen_altas.replace(0, '-').style.format(formatter))
            st.write("**Bajas por Motivo:**")
            st.dataframe(bajas_por_motivo.style.format(formatter))

        with tab3:
            st.header("Generador de Reportes Semanales")
            st.info("Selecciona una fecha de inicio. El PDF se generará con todos los datos del archivo, pero los títulos mostrarán el rango desde la fecha seleccionada hasta hoy.")
            
            start_date = st.date_input("Fecha de inicio para los títulos del reporte", datetime.now() - timedelta(days=7))
            
            if start_date:
                end_date = datetime.now()
                rango_str = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
                st.write(f"**Rango para los títulos del reporte:** {rango_str}")

                pdf_bytes = crear_pdf_reporte(
                    titulo_reporte="Resumen Semanal de Dotación",
                    rango_fechas_str=rango_str,
                    df_altas=df_altas_full,
                    df_bajas=df_bajas_full,
                    bajas_por_motivo=bajas_por_motivo.reset_index(),
                    resumen_altas=resumen_altas,
                    resumen_bajas=resumen_bajas,
                    resumen_activos=resumen_activos
                )
                
                st.download_button(
                    label="📄 Descargar Reporte Semanal en PDF",
                    data=pdf_bytes,
                    file_name=f"Reporte_Semanal_Dotacion_{start_date.strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                )

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.warning("Verifica que tu archivo Excel contenga las pestañas 'Activos' y 'BaseQuery' con las columnas necesarias.")
