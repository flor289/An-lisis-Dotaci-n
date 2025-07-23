import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime

# --- FUNCIÓN PARA CREAR EL PDF EJECUTIVO ---
def crear_pdf_resumen(n_altas, n_bajas, df_bajas_motivo, df_resumen_activos):
    """
    Genera un PDF con el resumen ejecutivo de los datos.
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)

    # Título
    pdf.cell(0, 10, "Resumen Ejecutivo de Dotación", ln=True, align="C")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
    pdf.ln(15)

    # Indicadores Clave
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Indicadores Clave del Periodo", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 8, f"- Cantidad de Altas: {n_altas}", ln=True)
    pdf.cell(0, 8, f"- Cantidad de Bajas: {n_bajas}", ln=True)
    pdf.ln(10)

    # Tabla de Bajas por Motivo
    if not df_bajas_motivo.empty:
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Desglose de Bajas por Motivo", ln=True)
        pdf.set_font("Arial", "B", 10)
        # Encabezado de la tabla
        pdf.cell(130, 8, "Motivo", 1)
        pdf.cell(40, 8, "Cantidad", 1, ln=True, align="C")
        # Cuerpo de la tabla
        pdf.set_font("Arial", "", 10)
        for index, row in df_bajas_motivo.iterrows():
            pdf.cell(130, 8, str(index), 1)
            pdf.cell(40, 8, str(row.iloc[0]), 1, ln=True, align="C")
        pdf.ln(10)

    # Tabla de Composición de Dotación Activa
    if not df_resumen_activos.empty:
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Composición de la Dotación Activa", ln=True)
        pdf.set_font("Arial", "B", 10)
        # Encabezado
        header = ['Gr.prof.'] + list(df_resumen_activos.columns)
        col_width = 180 / len(header)
        for item in header:
            pdf.cell(col_width, 8, str(item), 1, align="C")
        pdf.ln()
        # Cuerpo
        pdf.set_font("Arial", "", 10)
        for index, row in df_resumen_activos.iterrows():
            pdf.cell(col_width, 8, str(index), 1)
            for item in row:
                pdf.cell(col_width, 8, str(item), 1, align="C")
            pdf.ln()

    # Retorna el PDF como bytes
    return pdf.output(dest='S')

# --- CONFIGURACIÓN E INTERFAZ DE LA APP ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")

# Estilos CSS para un look más profesional
st.markdown("""
<style>
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
div[data-testid="stTabs"] button {
    font-size: 1.1rem;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Control de Dotación")
st.write("Sube tu archivo Excel con las pestañas 'Activos' y 'BaseQuery' para analizar las novedades y ver resúmenes.")

# Widget para subir el archivo
uploaded_file = st.file_uploader(
    "Selecciona tu archivo Excel de dotación", 
    type=['xlsx'],
    key="uploader" # Key para estabilidad
)

if uploaded_file:
    try:
        # Leer los datos del Excel en memoria
        df_activos = pd.read_excel(uploaded_file, sheet_name='Activos', engine='openpyxl')
        df_base = pd.read_excel(uploaded_file, sheet_name='BaseQuery', engine='openpyxl')
        
        st.success("¡Archivo cargado y procesado con éxito!")

        # --- Lógica de Procesamiento ---
        activos_legajos = set(df_activos['Nº pers.'])
        df_bajas = df_base[df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
        df_altas = df_base[~df_base['Nº pers.'].isin(activos_legajos) & (df_base['Status ocupación'] == 'Activo')].copy()

        # --- Preparar datos para el PDF y Dashboard ---
        bajas_por_motivo = df_bajas['Motivo de la medida'].value_counts().reset_index()
        df_activos_actuales = df_base[df_base['Status ocupación'] == 'Activo']
        resumen_dotacion_activa = pd.crosstab(index=df_activos_actuales['Gr.prof.'], columns=df_activos_actuales['División de personal'])
        
        # --- Botón de Descarga del PDF ---
        pdf_bytes = crear_pdf_resumen(len(df_altas), len(df_bajas), bajas_por_motivo, resumen_dotacion_activa)
        st.download_button(
            label="📄 Descargar Resumen Ejecutivo en PDF",
            data=pdf_bytes,
            file_name=f"Resumen_Ejecutivo_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            key="pdf_downloader"
        )
        st.markdown("---")

        # --- Pestañas para organizar la vista ---
        tab1, tab2, tab3 = st.tabs(["▶️ Novedades (Detalle)", "📈 Dashboard de Resúmenes", "🔄 Actualizar Activos"])
        
        with tab1:
            st.header("Detalle de Novedades")
            st.subheader(f"Altas ({len(df_altas)})")
            if not df_altas.empty:
                st.dataframe(df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'División de personal', 'Gr.prof.']])
            else:
                st.info("No se encontraron nuevas altas.")
            
            st.subheader(f"Bajas ({len(df_bajas)})")
            if not df_bajas.empty:
                st.dataframe(df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'División de personal', 'Gr.prof.']])
            else:
                st.info("No se encontraron bajas.")

        with tab2:
            st.header("Dashboard de Resúmenes")
            st.subheader("Composición de la Dotación Activa")
            st.dataframe(resumen_dotacion_activa)
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Resumen de Bajas")
                if not df_bajas.empty:
                    st.write("**Bajas por Motivo:**")
                    st.dataframe(df_bajas['Motivo de la medida'].value_counts())
                    st.write("**Bajas por Grupo y División:**")
                    st.dataframe(pd.crosstab(index=df_bajas['Gr.prof.'], columns=df_bajas['División de personal']))
                else:
                    st.info("No hay bajas para resumir.")
            with col2:
                st.subheader("Resumen de Altas")
                if not df_altas.empty:
                    st.write("**Altas por Grupo y División:**")
                    st.dataframe(pd.crosstab(index=df_altas['Gr.prof.'], columns=df_altas['División de personal']))
                else:
                    st.info("No hay altas para resumir.")

        with tab3:
            st.header("Actualizar Lista de Activos")
            st.info("Haz clic aquí para descargar la lista de legajos que quedaron activos. Úsala como la pestaña 'Activos' en tu próximo análisis.")
            df_nuevos_activos = df_base[df_base['Status ocupación'] == 'Activo'][['Nº pers.']]
            st.download_button(
                label="📥 Descargar 'Activos_actualizados.csv'",
                data=df_nuevos_activos.to_csv(index=False).encode('utf-8'),
                file_name='Activos_actualizados.csv',
                mime='text/csv',
                key="csv_downloader"
            )
            
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.warning("Verifica que tu archivo Excel contenga las pestañas 'Activos' y 'BaseQuery'.")
