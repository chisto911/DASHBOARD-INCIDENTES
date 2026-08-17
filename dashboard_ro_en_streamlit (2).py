import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Control RO - Inteligencia Operativa", 
    page_icon="📊", 
    layout="wide"
)

def clean_text(text):
    """Limpia los textos para facilitar la detección de columnas (quita tildes y pasa a minúsculas)"""
    if pd.isna(text): return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    return text

@st.cache_data
def load_and_process_data(file):
    """Procesa el Excel cargado usando el motor de detección de cabeceras. Usa caché para optimizar."""
    try:
        df_raw = pd.read_excel(file, header=None)
    except Exception:
        return pd.DataFrame()
        
    # Detección Inteligente de Cabeceras
    keywords = ['id', 'req', 'codigo', 'ticket', 'incidente', 'area', 'departamento', 'sucursal', 
               'agencia', 'espera', 'cola', 'proceso', 'atencion', 'total', 'ciclo', 'tc', 'estado', 'fase', 'estatus']
    
    best_row = 0
    max_score = 0
    
    for i in range(min(20, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).apply(clean_text)
        score = sum(1 for val in row_vals if any(kw in val for kw in keywords))
        if score > max_score:
            max_score = score
            best_row = i
            
    # Leer el DataFrame usando la fila detectada como cabecera
    file.seek(0) # Resetear el puntero del archivo
    df = pd.read_excel(file, header=best_row)
    
    # Guardar las columnas originales y limpiar las internas para búsqueda
    orig_cols = df.columns
    cleaned_cols = [clean_text(c) for c in orig_cols]
    df.columns = cleaned_cols
    
    def find_col(possible_keys, ignore_keys=['comentario', 'observacion', 'detalle', 'descripcion']):
        for pk in possible_keys:
            for col in df.columns:
                if pk in col and not any(ign in col for ign in ignore_keys):
                    return col
        return None
        
    id_col = find_col(['id', 'codigo', 'incidente', 'ticket', 'req', '#', 'numero', 'caso'])
    area_col = find_col(['area', 'departamento', 'direccion', 'unidad', 'gerencia'])
    suc_col = find_col(['sucursal', 'agencia', 'regional', 'zona', 'oficina'])
    esp_col = find_col(['espera', 'cola', 'retraso'])
    proc_col = find_col(['proceso', 'atencion', 'ejecucion'])
    tot_col = find_col(['total', 'ciclo', 'tc', 'lead time', 'dias'])
    
    # --- MEJORA: Detección Robusta de la Columna Estado ---
    estado_idx = None
    for i, col in enumerate(cleaned_cols):
        # Ampliamos los términos de búsqueda para garantizar que encuentre la columna
        if any(k in col for k in ['estado', 'estatus', 'fase', 'situacion', 'etapa']) and not any(ign in col for ign in ['comentario', 'observacion', 'detalle', 'descripcion']):
            estado_idx = i
            break
            
    # Heurística de respaldo por contenido si no se encuentra en el encabezado
    if estado_idx is None:
        estados_validos = ['en analisis', 'socializado', 'revision', 'no es riesgo', 'continuidad']
        for i in range(min(50, len(df_raw))):
            for j, val in enumerate(df_raw.iloc[i]):
                if pd.isna(val): continue
                val_str = clean_text(str(val))
                if any(ev in val_str for ev in estados_validos):
                    col_name = cleaned_cols[j]
                    if not any(ign in col_name for ign in ['comentario', 'observacion', 'detalle', 'descripcion', 'nota']):
                        estado_idx = j
                        break
            if estado_idx is not None:
                break
    
    # Construcción del DataFrame Final Estructurado
    res = pd.DataFrame()
    res['ID'] = df[id_col].astype(str) if id_col else [f"REQ-{i+1}" for i in range(len(df))]
    res['Área'] = df[area_col].astype(str).fillna("No especificada") if area_col else "No especificada"
    res['Sucursal'] = df[suc_col].astype(str).fillna("Sin Sucursal") if suc_col else "Sin Sucursal"
    
    res['Espera'] = pd.to_numeric(df[esp_col], errors='coerce').fillna(0) if esp_col else 0
    res['Proceso'] = pd.to_numeric(df[proc_col], errors='coerce').fillna(0) if proc_col else 0
    
    if tot_col:
        res['Total'] = pd.to_numeric(df[tot_col], errors='coerce').fillna(res['Espera'] + res['Proceso'])
    else:
        res['Total'] = res['Espera'] + res['Proceso']
        
    # Recuperar estado de forma segura por índice de columna
    if estado_idx is not None:
        res['Estado'] = df.iloc[:, estado_idx].astype(str)
    else:
        res['Estado'] = "Sin Clasificar"
        
    # Limpieza exhaustiva de valores vacíos y normalización de mayúsculas/minúsculas
    res['Estado'] = res['Estado'].replace(['nan', 'NaN', 'None', '', ' ', '<NA>'], 'Sin Clasificar')
    res['Estado'] = res['Estado'].fillna('Sin Clasificar')
    res['Estado'] = res['Estado'].str.strip().str.title() 
    
    # Filtrar registros válidos ampliando las condiciones para no perder casos
    res = res[(res['Total'] > 0) | (res['Espera'] > 0) | (res['Proceso'] > 0) | (res['Estado'] != 'Sin Clasificar')]
    
    # Cálculo de EC para la tabla
    res['EC (%)'] = np.where(res['Total'] > 0, (res['Proceso'] / res['Total']) * 100, 0)
    res['EC (%)'] = res['EC (%)'].round(1)
    
    # Redondear tiempos
    res['Espera'] = res['Espera'].round(1)
    res['Proceso'] = res['Proceso'].round(1)
    res['Total'] = res['Total'].round(1)
    
    return res

st.sidebar.title("Carga de Datos")
st.sidebar.markdown("Sube el archivo **DEPURACIÓN BASE INCIDENTES RO.xlsx** para visualizar las métricas dinámicas.")
uploaded_file = st.sidebar.file_uploader("Seleccionar Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👋 Bienvenido. Por favor, sube un archivo Excel en el menú lateral para comenzar a visualizar los datos.")
    st.stop()

# Cargar y procesar datos
with st.spinner('Analizando y estructurando datos...'):
    df = load_and_process_data(uploaded_file)

if df.empty:
    st.error("No se pudieron extraer datos válidos del archivo. Verifica que contenga los indicadores correctos.")
    st.stop()

st.title("Monitoreo de Eficiencia: Incidentes RO")

tab1, tab2, tab3 = st.tabs(["Dashboard General", "Alertas y Casos", "Desempeño Regional"])

# ================= TAB 1: DASHBOARD GENERAL =================
with tab1:
    # Métricas Principales (KPIs)
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="Media Total de Ciclo", 
            value=f"{df['Total'].mean():.1f} días", 
            delta=f"Mediana: {df['Total'].median():.1f} d",
            delta_color="off"
        )
        
    with col2:
        ec = (df['Proceso'].sum() / df['Total'].sum()) * 100 if df['Total'].sum() > 0 else 0
        st.metric(
            label="Eficiencia Global (EC)", 
            value=f"{ec:.1f}%",
            delta="Proceso vs Tiempo Total",
            delta_color="off"
        )
        
    with col3:
        criticos = len(df[df['Total'] > 85.0])
        st.metric(
            label="Casos Críticos (> 85d)", 
            value=f"{criticos} incidentes",
            delta="Exceden el SLA",
            delta_color="inverse"
        )
    
    st.divider()
    
    col_chart1, col_chart2 = st.columns([2, 1])
    
    with col_chart1:
        st.subheader("Tiempos Medios por Sucursal")
        df_grp = df.groupby('Sucursal').agg({'Espera':'mean', 'Proceso':'mean', 'Total':'mean'}).reset_index()
        df_grp = df_grp.sort_values('Total', ascending=False)
        
        fig_unified = px.bar(df_grp, x='Sucursal', y=['Espera', 'Proceso'], 
                     barmode='stack', color_discrete_sequence=['#cbd5e1', '#6366f1'])
        fig_unified.update_layout(legend_title_text='Fase', margin=dict(t=10, l=0, r=0, b=0))
        st.plotly_chart(fig_unified, use_container_width=True)
        
    with col_chart2:
        st.subheader("Fases y Resoluciones")
        df_st = df['Estado'].value_counts().reset_index()
        df_st.columns = ['Estado', 'Volumen']
        df_st = df_st.sort_values('Volumen', ascending=True) 
        
        # Ajuste dinámico más agresivo: 45 píxeles de alto por cada estado distinto encontrado
        altura_dinamica = max(400, len(df_st) * 45) 
        
        fig_status = px.bar(df_st, x='Volumen', y='Estado', orientation='h', color_discrete_sequence=['#10b981'], text='Volumen')
        fig_status.update_layout(
            margin=dict(t=10, l=0, r=20, b=0),
            height=altura_dinamica, 
            yaxis={'categoryorder': 'total ascending'}
        )
        
        # CRÍTICO: Esta instrucción obliga a Plotly a pintar el nombre de TODAS las filas sin saltarse ninguna
        fig_status.update_yaxes(tickmode='linear', dtick=1)
        
        st.plotly_chart(fig_status, use_container_width=True)


# ================= TAB 2: ALERTAS Y CASOS =================
with tab2:
    col_alert1, col_alert2 = st.columns([1, 2])
    
    with col_alert1:
        st.subheader("Distribución de Alertas")
        conds = [df['Total'] <= 10, df['Total'] <= 85, df['Total'] > 85]
        choices = ['Verde (≤ 10 d)', 'Amarilla (11-85 d)', 'Roja (> 85 d)']
        df['Alerta'] = np.select(conds, choices, default='Desconocido')
        
        df_alerta = df['Alerta'].value_counts().reset_index()
        color_map = {'Verde (≤ 10 d)':'#10b981', 'Amarilla (11-85 d)':'#fbbf24', 'Roja (> 85 d)':'#f43f5e'}
        
        fig_alert = px.pie(df_alerta, values='count', names='Alerta', hole=0.7, color='Alerta', color_discrete_map=color_map)
        fig_alert.update_layout(margin=dict(t=10, l=10, r=10, b=10))
        st.plotly_chart(fig_alert, use_container_width=True)
        
    with col_alert2:
        st.subheader(f"Base de Datos ({len(df)} registros)")
        # Mostramos la tabla filtrable y ordenada
        cols_to_show = ['ID', 'Área', 'Sucursal', 'Espera', 'Proceso', 'Total', 'Estado', 'EC (%)']
        st.dataframe(df[cols_to_show], use_container_width=True, hide_index=True)


# ================= TAB 3: DESEMPEÑO REGIONAL =================
with tab3:
    col_reg1, col_reg2 = st.columns([2, 1])
    
    with col_reg1:
        st.subheader("Ranking de Eficiencia por Sucursal")
        
        grp = df.groupby('Sucursal').agg(
            Count=('ID', 'count'),
            Total=('Total', 'mean'),
            Proceso=('Proceso', 'mean')
        ).reset_index()
        grp['EC'] = np.where(grp['Total'] > 0, (grp['Proceso'] / grp['Total']) * 100, 0)
        grp = grp.sort_values('Total', ascending=False)
        
        # Generar un Grid de métricas
        cols_per_row = 2
        for i in range(0, len(grp), cols_per_row):
            cols = st.columns(cols_per_row)
            for j in range(cols_per_row):
                if i + j < len(grp):
                    row = grp.iloc[i+j]
                    with cols[j]:
                        st.metric(
                            label=row['Sucursal'], 
                            value=f"{row['Total']:.1f} d (Promedio)", 
                            delta=f"{row['Count']} casos | EC: {row['EC']:.1f}%",
                            delta_color="off"
                        )
                        st.divider()

    with col_reg2:
        st.subheader("Concentración de Casos")
        df_reg = df['Sucursal'].value_counts().reset_index()
        fig_reg = px.pie(df_reg, values='count', names='Sucursal', hole=0.5)
        fig_reg.update_layout(margin=dict(t=10, l=10, r=10, b=10))
        st.plotly_chart(fig_reg, use_container_width=True)