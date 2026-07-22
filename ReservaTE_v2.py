import streamlit as st
import pandas as pd
import io

# Configuración de página ampliada para el dashboard
st.set_page_config(page_title="QA Report - Bookings TE", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# INYECCIÓN DE CSS (ESTILO OSCURO Y TARJETAS)
# ==========================================
st.markdown("""
<style>
    /* Estilo para las tarjetas de los KPIs */
    div[data-testid="metric-container"] {
        background-color: #1E293B;
        border: 1px solid #334155;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    div[data-testid="metric-container"] label {
        color: #94A3B8 !important;
        font-weight: 600;
        font-size: 0.9rem;
    }
    div[data-testid="metric-container"] div {
        color: #F8FAFC !important;
        font-size: 1.8rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# CABECERA CON LOGO
# ==========================================
col_logo, col_titulo = st.columns([1, 5])

with col_logo:
    # Asegúrate de que el nombre del archivo coincida exactamente con el que subas a GitHub
    try:
        st.image("Logo Exporinter AI_Letra Blanca.png", width=180)
    except:
        st.write("*(Logo no encontrado)*")

with col_titulo:
    st.markdown("## RECLAMOS PRESENTADOS • QA REPORT")
    st.caption("ORIGEN • EAI • CONSIGNEE • MAWB • HAWB • FBE • PIECES")

st.divider()

# ==========================================
# ZONA DE CARGA DE ARCHIVOS
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    f1_file = st.file_uploader("1. F1_Reserva", key="f1")
with col2:
    f2_file = st.file_uploader("2. F2_bookingsweb", key="f2")
with col3:
    f3_file = st.file_uploader("3. F3_DesignerBookings", key="f3")

# (Aquí van tus funciones normalizar_key, cargar_excel y limpiar_columnas exactamente igual que en app.py)
def normalizar_key(serie):
    return serie.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

def cargar_excel(archivo, **kwargs):
    if archivo.name.lower().endswith('.xlsx'):
        return pd.read_excel(archivo, engine='openpyxl', **kwargs)
    else:
        return pd.read_excel(archivo, engine='xlrd', **kwargs)

def limpiar_columnas(df):
    df.columns = df.columns.astype(str).str.strip()
    return df

# ==========================================
# PROCESAMIENTO Y DASHBOARD
# ==========================================
if f1_file and f2_file and f3_file:
    if st.button("🚀 Procesar y Generar Dashboard", use_container_width=True):
        try:
            with st.spinner("Procesando datos..."):
                
                # =====================================================================
                # PEGA AQUÍ TODO EL BLOQUE DE PROCESAMIENTO DE F1, F2 y F3 
                # (Exactamente el mismo código que ya tienes desde "1. PROCESAR F1..." 
                # hasta justo antes de "MOSTRAR Y DESCARGAR")
                # =====================================================================
                
                # SIMULACIÓN RÁPIDA DEL CRUCE PARA ESTE EJEMPLO (Omite esto al copiar tu código real)
                df_f1 = cargar_excel(f1_file, sheet_name="Booking", skiprows=4)
                df_final = df_f1 # Reemplaza esta línea con todo tu código de cruce real.

                # Asegurar tipos de datos para las sumatorias
                if 'FEB' in df_final.columns:
                    df_final['FEB'] = pd.to_numeric(df_final['FEB'], errors='coerce').fillna(0)
                if 'PIEZAS' in df_final.columns:
                    df_final['PIEZAS'] = pd.to_numeric(df_final['PIEZAS'], errors='coerce').fillna(0)

            # ==========================================
            # RENDERIZAR KPIs (NUEVO)
            # ==========================================
            st.markdown("### 📊 Indicadores Clave de Rendimiento (KPIs)")
            
            # Fila 1: 4 KPIs
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.metric("Origin (ORG)", df_final['ORG'].nunique() if 'ORG' in df_final.columns else 0)
            with k2:
                st.metric("EAI (UP)", df_final['UP'].nunique() if 'UP' in df_final.columns else 0)
            with k3:
                st.metric("Consignee (UC)", df_final['UC'].nunique() if 'UC' in df_final.columns else 0)
            with k4:
                st.metric("MAWB", df_final['MAWB'].nunique() if 'MAWB' in df_final.columns else 0)
            
            st.write("") # Espaciador
            
            # Fila 2: 3 KPIs
            k5, k6, k7 = st.columns(3)
            with k5:
                st.metric("HAWB", df_final['HAWB'].nunique() if 'HAWB' in df_final.columns else 0)
            with k6:
                st.metric("Total FBE (Cajas)", f"{df_final['FEB'].sum():,.0f}")
            with k7:
                st.metric("Total Pieces", f"{df_final['PIEZAS'].sum():,.0f}")

            st.divider()

            # ==========================================
            # MOSTRAR Y DESCARGAR (Igual que antes)
            # ==========================================
            st.success(f"¡Cruce finalizado! Se procesaron {len(df_final)} registros.")
            st.dataframe(df_final, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Merge_Final')
            
            st.download_button(
                label="📥 Descargar Resultado Excel",
                data=buffer.getvalue(),
                file_name="Merge_WebBookings_DesignerBookings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Se ha producido un error durante el cruce: {str(e)}")