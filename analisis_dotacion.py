import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime

# --- FUNCIÓN PARA CREAR EL PDF EJECUTIVO ---
def crear_pdf_resumen(n_altas, n_bajas, df_bajas_motivo, df_resumen_activos):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Resumen Ejecutivo de Dotación", ln=True, align="C")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y')}", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Indicadores Clave del Periodo", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 8, f"- Cantidad de Altas: {n_altas}", ln=True)
    pdf.cell(0, 8, f"- Cantidad de Bajas: {n_bajas}", ln=True)
    pdf.ln(8)

    if not df_bajas_motivo.empty:
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Desglose de Bajas por Motivo", ln=True)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(130, 8, "Motivo", 1)
        pdf.cell(40, 8, "Cantidad", 1, ln=True, align="C")
        pdf.set_font("Arial", "", 10)
        for _, row in df_bajas_motivo.iterrows():
            pdf.cell(130, 8, str(row['Motivo de la medida']), 1)
            pdf.cell(40, 8, str(row['Cantidad']), 1, ln=True, align="C")
        pdf.ln(8)

    if not df_resumen_activos.empty:
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Composición de la Dotación Activa", ln=True)
        pdf.set_font("Arial", "B", 8)
        header = ['Categoría'] + list(df_resumen_activos.columns)
        col_width = 180 / len(header)
        for item in header:
            pdf.cell(col_width, 8, str(item), 1, align="C")
        pdf.ln()
        pdf.set_font("Arial", "", 8)
        for index, row in df_resumen_activos.iterrows():
            pdf.cell(col_width, 8, str(index), 1)
            for item in row:
                pdf.cell(col_width, 8, str(item), 1, align="C")
            pdf.ln()

    return bytes(pdf.output())

# --- CONFIGURACIÓN E INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")

st.markdown("""
<style>
/* Estilos CSS para un look más profesional */
.main .block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    background-color: #f0f2f6; 
}
h1, h2, h3 {
    color: #003366; /* Azul corporativo */
}
div.stDownloadButton > button {
    background-color: #28a745; /* Verde para acción principal */
    color: white;
    border-radius: 5px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Control de Dotación")
st.write("Sube tu archivo Excel para analizar las novedades y ver resúmenes.")

uploaded_file = st.file_uploader(
    "Selecciona tu archivo Excel de dotación", 
    type=['xlsx'],
    key="uploader"
)

if uploaded_file:
    try:
        df_base_raw = pd.read_excel(uploaded_file, sheet_name='BaseQuery', engine='openpyxl')
        df_activos_raw = pd.read_excel(uploaded_file, sheet_name='Activos', engine='openpyxl')
        
        # --- LIMPIEZA Y PREPARACIÓN DE DATOS ---
        df_base = df_base_raw.copy()
        df_base.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}, inplace=True)

        # Formatear fechas a solo fecha, sin hora
        for col in ['Fecha', 'Desde', 'Fecha nac.']:
            if col in df_base.columns:
                df_base[col] = pd.to_datetime(df_base[col]).dt.date
        
        # Definir orden personalizado
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU']
        
        df_base['Línea'] = pd.Categorical(df_base['Línea'], categories=orden_lineas, ordered=True)
        df_base['Categoría'] = pd.Categorical(df_base['Categoría'], categories=orden_categorias, ordered=True)

        # --- LÓGICA DE PROCESAMIENTO ---
        activos_legajos = set(df_activos_raw['Nº pers.'])
        df_bajas = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        df_altas = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()

        # --- PREPARAR DATOS PARA DASHBOARD ---
        df_activos_actuales = df_base[df_base['Status ocupación'] == 'Activo']
        resumen_activos = pd.crosstab(df_activos_actuales['Categoría'], df_activos_actuales['Línea'], margins=True, margins_name="Total")
        bajas_por_motivo = df_bajas['Motivo de la medida'].value_counts().reset_index().rename(columns={"count": "Cantidad"})

        st.success("¡Archivo cargado y procesado con éxito!")
        
        # --- BOTÓN DE DESCARGA PDF ---
        pdf_bytes = crear_pdf_resumen(len(df_altas), len(df_bajas), bajas_por_motivo, resumen_activos.drop('Total', axis=1).drop('Total', axis=0))
        st.download_button(
            label="📄 Descargar Resumen Ejecutivo en PDF",
            data=pdf_bytes,
            file_name=f"Resumen_Ejecutivo_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
        )
        st.markdown("---")

        # --- PESTAÑAS DE NAVEGACIÓN ---
        tab1, tab2, tab3 = st.tabs(["▶️ Novedades (Detalle)", "📈 Dashboard de Resúmenes", "🔄 Actualizar Activos"])
        
        with tab1:
            st.header("Detalle de Novedades")
            st.subheader(f"Altas ({len(df_altas)})")
            if not df_altas.empty:
                st.dataframe(df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'Línea', 'Categoría']])
            
            st.subheader(f"Bajas ({len(df_bajas)})")
            if not df_bajas.empty:
                st.dataframe(df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'Línea', 'Categoría']])

        with tab2:
            st.header("Dashboard de Resúmenes")
            st.subheader("Composición de la Dotación Activa")
            st.dataframe(resumen_activos)
            
            st.subheader("Resumen de Novedades")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Bajas por Grupo y Línea:**")
                if not df_bajas.empty:
                    st.dataframe(pd.crosstab(df_bajas['Categoría'], df_bajas['Línea'], margins=True, margins_name="Total"))
                else:
                    st.info("No hay bajas para resumir.")
            with col2:
                st.write("**Altas por Grupo y Línea:**")
                if not df_altas.empty:
                    st.dataframe(pd.crosstab(df_altas['Categoría'], df_altas['Línea'], margins=True, margins_name="Total"))
                else:
                    st.info("No hay altas para resumir.")
            
            st.write("**Bajas por Motivo:**")
            if not bajas_por_motivo.empty:
                st.dataframe(bajas_por_motivo)

        with tab3:
            st.header("Actualizar Lista de Activos")
            st.info("Haz clic para descargar la lista de legajos que quedaron activos para tu próximo análisis.")
            df_nuevos_activos = df_base[df_base['Status ocupación'] == 'Activo'][['Nº pers.']]
            st.download_button(
                label="📥 Descargar 'Activos_actualizados.csv'",
                data=df_nuevos_activos.to_csv(index=False).encode('utf-8'),
                file_name='Activos_actualizados.csv',
                mime='text/csv',
            )
            
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.warning("Verifica que tu archivo Excel contenga las pestañas 'Activos' y 'BaseQuery' y los nombres de columnas correctos.")
