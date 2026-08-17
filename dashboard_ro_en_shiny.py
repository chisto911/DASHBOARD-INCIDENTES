import pandas as pd
import numpy as np
import unicodedata
import plotly.express as px
import plotly.graph_objects as go
from shiny import App, ui, render, reactive
from shinywidgets import output_widget, render_widget

def clean_text(text):
    """Limpia los textos para facilitar la detección de columnas (quita tildes y pasa a minúsculas)"""
    if pd.isna(text): return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    return text

app_ui = ui.page_navbar(
    ui.nav_panel(
        "Dashboard General",
        ui.layout_columns(
            ui.value_box(
                "Media Total de Ciclo", 
                ui.output_text("kpi_media"), 
                ui.output_text("kpi_mediana"), 
                theme="primary"
            ),
            ui.value_box(
                "Eficiencia Global (EC)", 
                ui.output_text("kpi_ec"), 
                "Proceso vs Tiempo Total", 
                theme="info"
            ),
            ui.value_box(
                "Casos Críticos (> 85d)", 
                ui.output_text("kpi_criticos"), 
                "Incidentes que exceden el SLA", 
                theme="danger"
            ),
        ),
        ui.layout_columns(
            ui.card(
                ui.card_header("Tiempos Medios por Sucursal (Espera vs Proceso)"),
                output_widget("unified_chart")
            ),
            ui.card(
                ui.card_header("Fases y Resoluciones"),
                output_widget("status_chart")
            ),
            col_widths=[8, 4]
        )
    ),
    ui.nav_panel(
        "Alertas y Casos",
        ui.layout_columns(
            ui.card(
                ui.card_header("Distribución de Alertas"),
                output_widget("alert_chart")
            ),
            ui.card(
                ui.card_header("Base de Datos de Incidentes"),
                ui.output_data_frame("main_table")
            ),
            col_widths=[4, 8]
        )
    ),
    ui.nav_panel(
        "Desempeño Regional",
        ui.layout_columns(
            ui.card(
                ui.card_header("Ranking de Eficiencia por Sucursal"),
                ui.output_ui("regional_cards")
            ),
            ui.card(
                ui.card_header("Concentración de Casos"),
                output_widget("regional_pie")
            ),
            col_widths=[8, 4]
        )
    ),
    title="Control RO - Inteligencia Operativa",
    id="tabs",
    sidebar=ui.sidebar(
        ui.h4("Carga de Datos"),
        ui.input_file("file_upload", "Seleccionar Excel (.xlsx)", accept=[".xlsx", ".xls"]),
        ui.markdown("Sube el archivo **DEPURACIÓN BASE INCIDENTES RO.xlsx** para visualizar las métricas dinámicas."),
        width="300px"
    )
)

def server(input, output, session):

    @reactive.Calc
    def processed_data():
        """Procesa el Excel cargado usando el motor de detección de cabeceras"""
        f = input.file_upload()
        if not f:
            return pd.DataFrame()
            
        filepath = f[0]["datapath"]
        
        try:
            df_raw = pd.read_excel(filepath, header=None)
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
        df = pd.read_excel(filepath, header=best_row)
        
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
            
        # Mapeo de columnas
        id_col = find_col(['id', 'codigo', 'incidente', 'ticket', 'req', '#', 'numero', 'caso'])
        area_col = find_col(['area', 'departamento', 'direccion', 'unidad', 'gerencia'])
        suc_col = find_col(['sucursal', 'agencia', 'regional', 'zona', 'oficina'])
        esp_col = find_col(['espera', 'cola', 'retraso'])
        proc_col = find_col(['proceso', 'atencion', 'ejecucion'])
        tot_col = find_col(['total', 'ciclo', 'tc', 'lead time', 'dias'])
        est_col = find_col(['estado', 'estatus', 'fase', 'situacion'])
        
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
            
        # Para estado, recuperamos los valores usando el índice de la columna para mantener tildes y mayúsculas
        if est_col:
            idx = df.columns.get_loc(est_col)
            res['Estado'] = df.iloc[:, idx].astype(str).fillna("Sin Clasificar")
        else:
            res['Estado'] = "Sin Clasificar"
            
        res['Estado'] = res['Estado'].replace('nan', 'Sin Clasificar')
        
        # Filtrar registros válidos
        res = res[(res['Total'] > 0) | (res['Estado'] != 'Sin Clasificar')]
        
        # Cálculo de EC para la tabla
        res['EC (%)'] = np.where(res['Total'] > 0, (res['Proceso'] / res['Total']) * 100, 0)
        res['EC (%)'] = res['EC (%)'].round(1)
        
        # Redondear tiempos
        res['Espera'] = res['Espera'].round(1)
        res['Proceso'] = res['Proceso'].round(1)
        res['Total'] = res['Total'].round(1)
        
        return res

    @render.text
    def kpi_media():
        df = processed_data()
        if df.empty: return "-- días"
        return f"{df['Total'].mean():.1f} días"

    @render.text
    def kpi_mediana():
        df = processed_data()
        if df.empty: return "Mediana: -- días"
        return f"Mediana: {df['Total'].median():.1f} días"

    @render.text
    def kpi_ec():
        df = processed_data()
        if df.empty or df['Total'].sum() == 0: return "-- %"
        ec = (df['Proceso'].sum() / df['Total'].sum()) * 100
        return f"{ec:.1f}%"

    @render.text
    def kpi_criticos():
        df = processed_data()
        if df.empty: return "-- activos"
        criticos = len(df[df['Total'] > 85.0])
        return f"{criticos} incidentes"

    @render_widget
    def unified_chart():
        df = processed_data()
        if df.empty: return go.Figure()
        
        df_grp = df.groupby('Sucursal').agg({'Espera':'mean', 'Proceso':'mean', 'Total':'mean'}).reset_index()
        df_grp = df_grp.sort_values('Total', ascending=False)
        
        fig = px.bar(df_grp, x='Sucursal', y=['Espera', 'Proceso'], 
                     title="Tiempos Medios (Apilados)", barmode='stack',
                     color_discrete_sequence=['#cbd5e1', '#6366f1'])
        
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', legend_title_text='Fase', margin=dict(t=40, l=0, r=0, b=0))
        return fig

    @render_widget
    def status_chart():
        df = processed_data()
        if df.empty: return go.Figure()
        
        df_st = df['Estado'].value_counts().reset_index()
        df_st.columns = ['Estado', 'Volumen']
        df_st = df_st.sort_values('Volumen', ascending=True) # Para que barras mayores queden arriba en horiz
        
        fig = px.bar(df_st, x='Volumen', y='Estado', orientation='h',
                     color_discrete_sequence=['#10b981'])
        
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=10, l=0, r=10, b=0))
        return fig

    @render_widget
    def alert_chart():
        df = processed_data()
        if df.empty: return go.Figure()
        
        conds = [df['Total'] <= 10, df['Total'] <= 85, df['Total'] > 85]
        choices = ['Verde (≤ 10 d)', 'Amarilla (11-85 d)', 'Roja (> 85 d)']
        df['Alerta'] = np.select(conds, choices, default='Desconocido')
        
        df_alerta = df['Alerta'].value_counts().reset_index()
        color_map = {'Verde (≤ 10 d)':'#10b981', 'Amarilla (11-85 d)':'#fbbf24', 'Roja (> 85 d)':'#f43f5e'}
        
        fig = px.pie(df_alerta, values='count', names='Alerta', hole=0.7, color='Alerta', color_discrete_map=color_map)
        fig.update_layout(margin=dict(t=10, l=10, r=10, b=10))
        return fig

    @render_widget
    def regional_pie():
        df = processed_data()
        if df.empty: return go.Figure()
        
        df_reg = df['Sucursal'].value_counts().reset_index()
        fig = px.pie(df_reg, values='count', names='Sucursal', hole=0.5)
        fig.update_layout(margin=dict(t=10, l=10, r=10, b=10))
        return fig

    @render.data_frame
    def main_table():
        df = processed_data()
        if df.empty:
            return pd.DataFrame()
        return render.DataGrid(df, selection_mode="row", filters=True)

    @render.ui
    def regional_cards():
        df = processed_data()
        if df.empty:
            return ui.p("Sube un archivo para ver las métricas regionales.", class_="text-muted")
            
        grp = df.groupby('Sucursal').agg(
            Count=('ID', 'count'),
            Total=('Total', 'mean'),
            Proceso=('Proceso', 'mean')
        ).reset_index()
        
        grp['EC'] = np.where(grp['Total'] > 0, (grp['Proceso'] / grp['Total']) * 100, 0)
        grp = grp.sort_values('Total', ascending=False)
        
        cards = []
        for _, row in grp.iterrows():
            color = "danger" if row['Total'] > 85 else "warning" if row['Total'] > 30 else "success"
            cards.append(
                ui.value_box(
                    row['Sucursal'],
                    f"{row['Total']:.1f} d",
                    f"{row['Count']} incidentes | EC: {row['EC']:.1f}%",
                    theme=color
                )
            )
            
        return ui.layout_columns(*cards)

app = App(app_ui, server)