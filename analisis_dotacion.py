import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import io

# --- CLASE MEJORADA PARA CREAR EL PDF EJECUTIVO ---
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin

    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, "Resumen de Dotación", 0, 0, "C")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

    def draw_table(self, title, df_original, is_crosstab=False):
        if df_original.empty:
            return
        
        df = df_original.copy()
        if is_crosstab:
            df = df.replace(0, '-')

        table_height = 8 * (len(df) + 1) + 10
        if self.get_y() + table_height > self.h - self.b_margin:
            self.add_page(orientation=self.cur_orientation)

        self.set_font("Arial", "B", 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, title, ln=True, align="L")
        self.ln(2)

        df_for_width_calc = df.reset_index() if is_crosstab else df
        widths = {}
        for col in df_for_width_calc.columns:
            header_width = self.get_string_width(str(col)) + 8
            content_width = df_for_width_calc[col].astype(str).apply(lambda x: self.get_string_width(x)).max() + 8
            widths[col] = max(header_width, content_width)

        self.set_font("Arial", "B", 9)
        self.set_fill_color(70, 130, 180)
        self.set_text_color(255, 255, 255)
        
        # Usar los nombres de columna del dataframe que incluye el índice reseteado
        for col_name in df_for_width_calc.columns:
            self.cell(widths[col_name], 8, str(col_name), 1, 0, "C", True)
        self.ln()
        
        self.set_text_color(0, 0, 0)
        for _, row in df_for_width_calc.iterrows():
            is_total_row = str(row.iloc[0]) == "Total"
            if is_total_row:
                self.set_font("Arial", "B", 9)
            else:
                self.set_font("Arial", "", 9)

            for col_name in df_for_width_calc.columns:
                self.cell(widths[col_name], 8, str(row[col_name]), 1, 0, "C")
            self.ln()
        self.ln(10)

def crear_pdf_completo(df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 10, f"Período Analizado (Fecha: {datetime.now().strftime('%d/%m/%Y')})", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"- Cantidad de Altas: {len(df_altas)}", ln=True)
    pdf.cell(0, 8, f"- Cantidad de Bajas: {len(df_bajas)}", ln=True)
    pdf.ln(5)

    pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'Línea', 'Categoría']])
    pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'Línea', 'Categoría']])
    pdf.draw_table("Bajas por Motivo", bajas_por_motivo.rename(columns={'Motivo de la medida': 'Motivo'}))
    
    # Pasamos las tablas con el índice reseteado para que la función las dibuje correctamente
    pdf.draw_table("Resumen de Altas por Categoría y Línea", resumen_altas.reset_index(), is_crosstab=True)
    pdf.draw_table("Resumen de Bajas por Categoría y Línea", resumen_bajas.reset_index(), is_crosstab=True)
    pdf.draw_table("Composición de la Dotación Activa", resumen_activos.reset_index(), is_crosstab=True)

    return bytes(pdf.output())

# --- CONFIGURACIÓN E INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")

st.markdown("""
<style>
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
                df_base[col] = pd.to_datetime(df_base[col], errors='coerce').dt.date
        
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
        
        df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
        df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)

        activos_legajos = set(df_activos_raw['Nº pers.'])
        df_bajas = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        df_altas = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()

        # --- PREPARAR DATOS PARA DASHBOARD (Quitamos dropna=False para que no muestre vacíos) ---
        df_activos_actuales = df_base[df_base['Status ocupación'] == 'Activo']
        resumen_activos = pd.crosstab(df_activos_actuales['Categoría'], df_activos_actuales['Línea'], margins=True, margins_name="Total")
        resumen_bajas = pd.crosstab(df_bajas['Categoría'], df_bajas['Línea'], margins=True, margins_name="Total")
        resumen_altas = pd.crosstab(df_altas['Categoría'], df_altas['Línea'], margins=True, margins_name="Total")
        
        bajas_por_motivo_series = df_bajas['Motivo de la medida'].value_counts()
        bajas_por_motivo = bajas_por_motivo_series.to_frame('Cantidad')
        if not bajas_por_motivo.empty:
            bajas_por_motivo.loc['Total'] = bajas_por_motivo_series.sum()
        bajas_por_motivo.reset_index(inplace=True)
        bajas_por_motivo.rename(columns={'index': 'Motivo'}, inplace=True)

        st.success("¡Archivo cargado y procesado!")
        
        pdf_bytes = crear_pdf_completo(df_altas, df_bajas, bajas_por_motivo, resumen_altas, resumen_bajas, resumen_activos)
        st.download_button(
            label="📄 Descargar Resumen en PDF",
            data=pdf_bytes,
            file_name=f"Resumen_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
        )
        st.markdown("---")

        tab1, tab2, tab3 = st.tabs(["▶️ Novedades (Detalle)", "📈 Dashboard de Resúmenes", "🔄 Actualizar Activos"])
        
        with tab1:
            st.header("Detalle de Novedades")
            st.subheader(f"Altas ({len(df_altas)})")
            st.dataframe(df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'Línea', 'Categoría']], hide_index=True)
            st.subheader(f"Bajas ({len(df_bajas)})")
            st.dataframe(df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'Línea', 'Categoría']], hide_index=True)

        with tab2:
            st.header("Dashboard de Resúmenes")
            st.subheader("Composición de la Dotación Activa")
            st.dataframe(resumen_activos.replace(0, '-'))
            st.subheader("Resumen de Novedades")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Bajas por Categoría y Línea:**")
                st.dataframe(resumen_bajas.replace(0, '-'))
            with col2:
                st.write("**Altas por Categoría y Línea:**")
                st.dataframe(resumen_altas.replace(0, '-'))
            st.write("**Bajas por Motivo:**")
            st.dataframe(bajas_por_motivo, hide_index=True)

        with tab3:
            st.header("Actualizar Lista de Activos")
            st.info("Haz clic para descargar el archivo Excel con los legajos que quedaron activos para tu próximo análisis.")
            df_nuevos_activos = df_base[df_base['Status ocupación'] == 'Activo'][['Nº pers.']]
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_nuevos_activos.to_excel(writer, index=False, sheet_name='Activos')
            st.download_button(
                label="📥 Descargar 'Activos_actualizados.xlsx'",
                data=output.getvalue(),
                file_name='Activos_actualizados.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.warning("Verifica que tu archivo Excel contenga las pestañas 'Activos' y 'BaseQuery'.")
