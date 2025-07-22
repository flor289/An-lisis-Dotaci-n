import streamlit as st
import pandas as pd

# --- Configuración de la página ---
st.set_page_config(page_title="Análisis de Dotación", layout="wide")
st.title("📊 Herramienta de Análisis de Dotación")
st.write("Sube tus archivos de 'Activos' y 'BaseQuery' para encontrar las altas y las bajas.")

# --- Sección para subir archivos ---
col1, col2 = st.columns(2)
with col1:
    activos_file = st.file_uploader("1. Sube tu lista de Activos (CSV o Excel)", type=['csv', 'xlsx'])

with col2:
    base_file = st.file_uploader("2. Sube tu archivo BaseQuery (CSV o Excel)", type=['csv', 'xlsx'])

# --- Lógica de procesamiento ---
if activos_file and base_file:
    try:
        # Leer los archivos
        df_activos = pd.read_csv(activos_file) if activos_file.name.endswith('csv') else pd.read_excel(activos_file)
        df_base = pd.read_csv(base_file) if base_file.name.endswith('csv') else pd.read_excel(base_file)

        # Extraer los legajos de activos
        activos_legajos = set(df_activos['Nº pers.'])

        # --- Identificar las BAJAS ---
        df_bajas = df_base[
            df_base['Nº pers.'].isin(activos_legajos) &
            (df_base['Status ocupación'] == 'Dado de baja')
        ].copy()

        # --- Identificar las ALTAS ---
        df_altas = df_base[
            ~df_base['Nº pers.'].isin(activos_legajos) &
            (df_base['Status ocupación'] == 'Activo')
        ].copy()

        st.success("¡Análisis completado!")

        # --- Mostrar resultados ---
        st.subheader("Altas Detectadas")
        if not df_altas.empty:
            st.dataframe(df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha', 'División de personal', 'Gr.prof.']])
        else:
            st.info("No se encontraron nuevas altas.")

        st.subheader("Bajas Detectadas")
        if not df_bajas.empty:
            st.dataframe(df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de la medida', 'Desde', 'División de personal', 'Gr.prof.']])
        else:
            st.info("No se encontraron bajas.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
        st.warning("Asegúrate de que las columnas ('Nº pers.', 'Status ocupación', etc.) existan en tus archivos.")