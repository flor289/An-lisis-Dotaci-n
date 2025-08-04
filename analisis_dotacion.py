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
             if pd.api.types.is_numeric_dtype(df_formatted[col]):
                  df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) and x != 0 else x)


        widths = {col: max(self.get_string_width(str(col)) + 8, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x)).max() + 8) for col in df_formatted.columns}
        total_width = sum(widths.values())
        
        font_size = 9
        if total_width > self.page_width:
            scaling_factor = self.page_width / total_width
            widths = {k: v * scaling_factor for k, v in widths.items()}
            font_size = 7

        self.set_font("Arial", "B", font_size)
        self.set_fill_color(70, 130, 180)
        self.set_text_color(255, 255, 255)
        
        for col in df_formatted.columns:
            self.cell(widths[col], 8, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_text_color(0, 0, 0)
        for _, row in df_formatted.iterrows():
            is_total_row = "Total" in str(row.iloc[0])
            if is_total_row:
                self.set_font("Arial", "B", font_size)
            else:
                self.set_font("Arial", "", font_size)

            for col in df_formatted.columns:
                self.cell(widths[col], 8, str(row[col]), 1, 0, "C")
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
    pdf.draw_table("Bajas por Motivo", bajas_por_motivo, is_crosstab=True)
    
    pdf.draw_table("Resumen de Altas por Categoría y Línea", resumen_altas, is_crosstab=True)
    pdf.draw_table("Resumen de Bajas por Categoría y Línea", resumen_bajas, is_crosstab=True)
    pdf.draw_table("Composición de la Dotación Activa", resumen_activos, is_crosstab=True)

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
                df_base[col] = pd.to_datetime(df_base[col], errors='coerce').dt.date
        
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
        
        df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
        df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)

        activos_legajos = set(df_activos_raw['Nº pers.'])
        df_bajas = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        df_altas = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()

        # --- PREPARAR DATOS PARA DASHBOARD ---
        df_activos_actuales = df_base[df_base['Status ocupación'] == 'Activo']
        resumen_activos = pd.crosstab(df_activos_actuales['Categoría'], df_activos_actuales['Línea'], margins=True, margins_name="Total")
        resumen_bajas = pd.crosstab(df_bajas['Categoría'], df_bajas['Línea'], margins=True, margins_name="Total")
        resumen_altas = pd.crosstab(df_altas['Categoría'], df_altas['Línea'], margins=True, margins_name="Total")
        
        bajas_por_motivo_series = df_bajas['Motivo de la medida'].value_counts()
        bajas_por_motivo = bajas_por_motivo_series.to_frame('Cantidad')
        if not bajas_por_motivo.empty:
            bajas_por_motivo.loc['Total'] = bajas_por_motivo_series.sum()
        bajas_por_motivo.index.name = "Motivo"

        st.success("¡Archivo cargado y procesado!")
        
        pdf_bytes = crear_pdf_completo(df_altas, df_bajas, bajas_por_motivo.reset_index(), resumen_altas, resumen_bajas, resumen_activos)
        st.download_button(
            label="📄 Descargar Resumen en PDF",
            data=pdf_bytes,
            file_name=f"Resumen_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
        )
        st.markdown("---")

        # --- PESTAÑAS DE NAVEGACIÓN ---
        tab1, tab2, tab3 = st.tabs(["▶️ Novedades (Detalle)", "📈 Dashboard de Resúmenes", "🔄 Actualizar Activos"])
        
        # --- Formateador para las tablas de Streamlit ---
        formatter = lambda x: f'{x:,.0f}'.replace(',', '.') if isinstance(x, (int, float)) else x

        with tab1:
            st.header("Detalle de Novedades")
            st.subheader(f"Altas ({len(df_altas)})")
            st.dataframe(df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'Línea', 'Categoría']], hide_index=True)
            st.subheader(f"Bajas ({len(df_bajas)})")
            st.dataframe(df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'Línea', 'Categoría']], hide_index=True)

        with tab2:
            st.header("Dashboard de Resúmenes")
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
