import streamlit as st
import pandas as pd
import io
import plotly.express as px

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
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
    # Asegúrate de que el nombre del archivo coincida exactamente con el que subiste a GitHub
    try:
        st.image("Logo.png", width=180)
    except:
        st.write("*(Logo no encontrado)*")

with col_titulo:
    st.markdown("## BOOKINGS TE")
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

# ==========================================
# FUNCIONES DE LIMPIEZA
# ==========================================
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
            with st.spinner("Procesando datos y generando visualizaciones..."):
                
                # ==========================================
                # 1. PROCESAR F1_RESERVA
                # ==========================================
                df_f1 = cargar_excel(f1_file, sheet_name="Booking", skiprows=4)
                df_f1 = limpiar_columnas(df_f1)
                
                df_f1['INDICE'] = pd.to_numeric(df_f1['MAWB'], errors='coerce').apply(
                    lambda x: 1 if pd.notnull(x) and x >= 1 else None
                )
                df_f1 = df_f1.dropna(subset=['MAWB'])
                
                df_f1 = df_f1.rename(columns={
                    "FECHA RESERVA": "FECHA", 
                    "CAJAS RESERVA": "FEB", 
                    "PIEZAS RESERVA": "PIEZAS"
                })
                df_f1['HAWB'] = normalizar_key(df_f1['HAWB'])

                # ==========================================
                # 2. PROCESAR F2_BOOKINGSWEB
                # ==========================================
                try:
                    df_f2 = cargar_excel(f2_file, sheet_name="bookingsweb")
                except Exception:
                    df_f2 = cargar_excel(f2_file)
                
                df_f2 = limpiar_columnas(df_f2)
                
                df_f2 = df_f2.rename(columns={
                    "awbno": "MAWB", 
                    "reference": "HAWB", 
                    "packing": "TIPO EMPAQUE", 
                    "pieces": "PIEZAS", 
                    "fulleq": "FEB"
                })
                
                reemplazos = {
                    "1 CUARTO": "Q", "1 FULL": "F", 
                    "1 OCTAVO": "E", "1 TABACO": "H", "1TERCIO": "R"
                }
                if 'TIPO EMPAQUE' in df_f2.columns:
                    df_f2['TIPO EMPAQUE'] = df_f2['TIPO EMPAQUE'].replace(reemplazos)
                
                df_f2['HAWB'] = normalizar_key(df_f2['HAWB'])

                # Merge F2 con F1
                columnas_f1 = ['HAWB', 'FECHA', 'UP', 'CD', 'ORG', 'VUELO', 'INDICE']
                cols_extraer = [c for c in columnas_f1 if c in df_f1.columns]
                
                df_f2 = pd.merge(df_f2, df_f1[cols_extraer], on="HAWB", how="left")
                
                if 'INDICE' in df_f2.columns:
                    df_f2 = df_f2[df_f2['INDICE'] == 1]
                if 'CD' in df_f2.columns:
                    df_f2 = df_f2.rename(columns={"CD": "UC"})
                if 'FEB' in df_f2.columns:
                    df_f2 = df_f2[df_f2['FEB'].astype(str) != "0"]

                # ==========================================
                # 3. PROCESAR F3_DESIGNERBOOKINGS (HTML/Excel)
                # ==========================================
                try:
                    try:
                        df_f3 = cargar_excel(f3_file, sheet_name="F3_FlowerOrderShipmentLines")
                    except Exception:
                        df_f3 = cargar_excel(f3_file)
                except Exception as e:
                    f3_file.seek(0)
                    content = f3_file.getvalue().decode('utf-8', errors='replace')
                    df_f3_list = pd.read_html(io.StringIO(content))
                    df_f3 = df_f3_list[0]
                    
                    if not any("HouseAWB" in str(c) for c in df_f3.columns):
                        for idx, row in df_f3.head(10).iterrows():
                            if any("HouseAWB" in str(val) for val in row.values):
                                df_f3.columns = row.values
                                df_f3 = df_f3.iloc[idx+1:].reset_index(drop=True)
                                break

                df_f3 = limpiar_columnas(df_f3)
                
                df_f3 = df_f3.rename(columns={
                    "HouseAWB #": "HAWB", 
                    "Box": "TIPO EMPAQUE", 
                    "Total Boxes": "PIEZAS", 
                    "Total Full Boxes": "FEB"
                })
                
                if 'MAWB' in df_f3.columns:
                    df_f3['MAWB'] = df_f3['MAWB'].astype(str).str.replace("-", "", regex=False)
                
                if 'HAWB' in df_f3.columns:
                    df_f3['HAWB'] = df_f3['HAWB'].astype(str).str.replace("-", "", regex=False)
                    df_f3['HAWB'] = normalizar_key(df_f3['HAWB'])
                else:
                    raise KeyError(f"No se encontró la columna 'HouseAWB #' en F3. Columnas detectadas: {list(df_f3.columns)}")

                df_f3 = pd.merge(df_f3, df_f1[cols_extraer], on="HAWB", how="left")
                
                if 'INDICE' in df_f3.columns:
                    df_f3 = df_f3[df_f3['INDICE'] == 1]
                if 'CD' in df_f3.columns:
                    df_f3 = df_f3.rename(columns={"CD": "UC"})

                # ==========================================
                # 4. APILAR TABLAS Y LIMPIAR
                # ==========================================
                df_final = pd.concat([df_f2, df_f3], ignore_index=True)
                
                if 'FEB' in df_final.columns:
                    df_final['FEB'] = pd.to_numeric(df_final['FEB'], errors='coerce').fillna(0)
                if 'PIEZAS' in df_final.columns:
                    df_final['PIEZAS'] = pd.to_numeric(df_final['PIEZAS'], errors='coerce').fillna(0)

                columnas_plantilla = ['FLOR', 'TALLOS', 'CONDUCTOR', 'LICENCIA', 'PLACA', 'PRECINTO']
                for col in columnas_plantilla:
                    df_final[col] = None 

                if 'INDICE' in df_final.columns:
                    df_final = df_final.drop(columns=['INDICE'])

                if 'FEB' in df_final.columns:
                    df_final = df_final[
                        df_final['FEB'].notnull() & 
                        (df_final['FEB'] != 0) & 
                        (df_final['FEB'].astype(str).str.strip() != "")
                    ]

                orden_columnas = [
                    "ORG", "FECHA", "UP", "UC", "MAWB", "HAWB", "FLOR", "FEB", 
                    "PIEZAS", "TIPO EMPAQUE", "TALLOS", "CONDUCTOR", "LICENCIA", 
                    "PLACA", "PRECINTO", "VUELO"
                ]
                orden_definitivo = [col for col in orden_columnas if col in df_final.columns]
                df_final = df_final[orden_definitivo]

            # ==========================================
            # 5. RENDERIZAR KPIs
            # ==========================================
            st.markdown("### 📊 Indicadores Clave de Rendimiento (KPIs)")
            
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.metric("Origin (ORG)", df_final['ORG'].nunique() if 'ORG' in df_final.columns else 0)
            with k2:
                st.metric("EAI (UP)", df_final['UP'].nunique() if 'UP' in df_final.columns else 0)
            with k3:
                st.metric("Consignee (UC)", df_final['UC'].nunique() if 'UC' in df_final.columns else 0)
            with k4:
                st.metric("MAWB", df_final['MAWB'].nunique() if 'MAWB' in df_final.columns else 0)
            
            st.write("") 
            
            k5, k6, k7 = st.columns(3)
            with k5:
                st.metric("HAWB", df_final['HAWB'].nunique() if 'HAWB' in df_final.columns else 0)
            with k6:
                st.metric("Total FBE (Cajas)", f"{df_final['FEB'].sum():,.0f}")
            with k7:
                st.metric("Total Pieces", f"{df_final['PIEZAS'].sum():,.0f}")

            st.divider()

            # ==========================================
            # 6. ZONA DE GRÁFICOS (PLOTLY)
            # ==========================================
            st.markdown("### 📈 Análisis de EAIs (UP)")
            
            if 'UP' in df_final.columns:
                top_eais = df_final['UP'].value_counts().reset_index()
                top_eais.columns = ['EAI (UP)', 'Cantidad']
                top_eais = top_eais.head(10)
                
                # Ordenar para que el mayor quede arriba
                top_eais = top_eais.sort_values('Cantidad', ascending=True)

                fig = px.bar(
                    top_eais, 
                    x='Cantidad', 
                    y='EAI (UP)', 
                    orientation='h',
                    text='Cantidad',
                    color_discrete_sequence=['#38BDF8'] # Color azul celeste
                )
                
                fig.update_layout(
                    template='plotly_dark',
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    xaxis_title=None,
                    yaxis_title=None,
                    xaxis=dict(showgrid=False, showticklabels=False),
                    margin=dict(l=0, r=0, t=30, b=0),
                    height=400
                )
                
                col_grafico, col_vacia = st.columns([1, 1])
                with col_grafico:
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("No se encontró la columna 'UP' para generar el gráfico.")

            st.divider()

            # ==========================================
            # 7. MOSTRAR DATAFRAME Y DESCARGAR
            # ==========================================
            st.success(f"¡Cruce finalizado! Se procesaron {len(df_final)} registros.")
            st.dataframe(df_final, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Merge_Final')
            
            st.download_button(
                label="📥 Descargar Resultado Excel",
                data=buffer.getvalue(),
                file_name="Result_BookingsTE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Se ha producido un error durante el procesamiento: {str(e)}")
else:
    st.info("Sube los 3 archivos fuente para iniciar el procesamiento y ver el dashboard.")
