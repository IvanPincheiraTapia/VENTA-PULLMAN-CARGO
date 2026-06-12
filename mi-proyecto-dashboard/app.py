import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import glob
import os

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="KPI Control de Carga - Pullman Cargo",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- INYECCIÓN DE CSS AVANZADO PARA MARCA PULLMAN CARGO ---
st.markdown("""
    <style>
    .stApp { background-color: #f1f5f9; }
    h1 {
        color: #1e3a8a !important; /* Azul Marino Pullman */
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 800 !important;
        margin-bottom: 5px !important;
    }
    h2, h3 { color: #0f172a !important; font-family: 'Segoe UI', sans-serif; font-weight: 700 !important; }
    .filter-container {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        border-top: 4px solid #1e3a8a;
        margin-bottom: 25px;
    }
    .kpi-card {
        background-color: #ffffff;
        padding: 22px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        border-left: 6px solid #1e3a8a;
        margin-bottom: 15px;
    }
    .kpi-title { color: #64748b; font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; }
    .kpi-value { color: #1e293b; font-size: 1.9rem; font-weight: 800; margin-top: 5px; }
    .chart-block { background-color: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 25px; }
    .logo-container { background-color: #ffffff; padding: 15px; border-radius: 10px; margin-bottom: 20px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# 2. MOTOR DE CARGA DE DATOS (Optimizado para cabeceras reales de VENTA_AGENCIAS.xlsx)
@st.cache_data(ttl=120)
def cargar_y_procesar_datos():
    carpeta_data = "DATA"
    # Ajustado con guion bajo para que coincida exactamente con tu archivo en GitHub
    archivo_objetivo = os.path.join(carpeta_data, "VENTA_AGENCIAS.xlsx")
    
    if not os.path.exists(archivo_objetivo):
        # Fallback si no está en DATA/ pero sí en la raíz
        if os.path.exists("VENTA_AGENCIAS.xlsx"):
            archivo_objetivo = "VENTA_AGENCIAS.xlsx"
        else:
            return None
    
    # Lectura limpia sin omitir filas para que cuadren los montos al 100%
    df = pd.read_excel(archivo_objetivo)
    df.columns = df.columns.str.strip()
    
    # Conversión estricta de columnas numéricas
    columnas_numericas = ['Valor Tarifa', 'Factor Descuento', 'Valor DD', 'Valor Seguro', 'Valor Descuento', 'Valor Total', 'Peso', 'Volumen', 'Bultos', 'LATITUD', 'LONGUITUD']
    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
    # Limpieza de la columna S (Fecha Emisión) para ver SOLO fecha (sin horas)
    if 'Fecha Emisión' in df.columns:
        df['Fecha Emisión'] = pd.to_datetime(df['Fecha Emisión'], errors='coerce').dt.date
        
    return df, os.path.basename(archivo_objetivo)

data_bundle = cargar_y_procesar_datos()

if data_bundle is not None:
    df_raw, nombre_archivo = data_bundle
    
    # --- MENÚ LATERAL: LOGO CORPORATIVO ---
    with st.sidebar:
        st.markdown('<div class="logo-container">', unsafe_allow_html=True)
        if os.path.exists("logo_pullmacargo_fondo_blanco.png"):
            st.image("logo_pullmacargo_fondo_blanco.png", use_container_width=True)
        else:
            st.subheader("Pullman Cargo S.A.")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown(f"**🟢 Archivo Activo:**\n`{nombre_archivo}`")
        st.markdown("---")
        st.caption("Control de Gestión Integrado.")

    # --- ENCABEZADO DE LA PÁGINA ---
    st.title("🚚 Panel de Control Operativo y Comercial")
    st.markdown("<p style='color:#64748b; font-size:1.1rem; margin-top:-10px;'>Sistemas de Carga Integrados — Pullman Cargo S.A.</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- SECCIÓN 1: FILTROS DE NAVEGACIÓN (CONEXIÓN EN CASCADA TOTAL) ---
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)
    st.markdown("#### 🔍 Filtros de Segmentación Operativa Global")
    
    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
    
    # Identificar la columna MES de tu base de datos de manera exacta
    col_mes_nombre = 'MES' if 'MES' in df_raw.columns else ('Mes' if 'Mes' in df_raw.columns else None)
    
    with col_f1:
        # 1. Filtro de Mes (Columna nativa de tu Excel)
        if col_mes_nombre:
            meses_disp = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO"]
            meses_existentes = [m for m in meses_disp if m in df_raw[col_mes_nombre].unique()]
            mes_sel = st.multiselect("1. Mes Consulta:", meses_existentes, default=meses_existentes)
        else:
            st.warning("Falta col MES")
            mes_sel = []

    with col_f2:
        df_f1 = df_raw[df_raw[col_mes_nombre].isin(mes_sel)] if col_mes_nombre else df_raw
        zonas_disp = sorted(df_f1['ZONA'].dropna().unique()) if 'ZONA' in df_f1.columns else []
        zona_sel = st.multiselect("2. Zona:", zonas_disp, default=zonas_disp)
    
    with col_f3:
        df_f2 = df_f1[df_f1['ZONA'].isin(zona_sel)] if zonas_disp else df_f1
        regiones_disp = sorted(df_f2['REGION'].dropna().unique()) if 'REGION' in df_f2.columns else []
        region_sel = st.multiselect("3. Región:", regiones_disp, default=regiones_disp)
        
    with col_f4:
        df_f3 = df_f2[df_f2['REGION'].isin(region_sel)] if regiones_disp else df_f2
        agencias_disp = sorted(df_f3['Agencia Descripción'].dropna().unique())
        agency_sel = st.multiselect("4. Agencia(s):", agencias_disp, default=agencias_disp)

    with col_f5:
        flujos_disp = sorted(df_raw['Tipo ODT'].dropna().unique()) if 'Tipo ODT' in df_raw.columns else ['EMITIDAS', 'RECIBIDAS', 'DESCUENTOS']
        flujo_sel = st.multiselect("5. Flujo:", flujos_disp, default=flujos_disp)
    
    # FILTRO DEFINITIVO PARA TODO EL DASHBOARD
    df_f4 = df_f3[df_f3['Agencia Descripción'].isin(agency_sel)]
    df = df_f4[df_f4['Tipo ODT'].isin(flujo_sel)]
    st.markdown('</div>', unsafe_allow_html=True)

    # --- INDICADORES GENERALES DEL PERÍODO ---
    total_ingresos = df['Valor Total'].sum()
    total_kilos = df['Peso'].sum()
    total_bultos = df['Bultos'].sum()
    
    venta_emitida_tot = df[df['Tipo ODT'] == 'EMITIDAS']['Valor Total'].sum()
    venta_recibida_tot = df[df['Tipo ODT'] == 'RECIBIDAS']['Valor Total'].sum()

    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #0f172a;"><div class="kpi-title">Venta Total Consolidada Filtrada</div><div class="kpi-value">${total_ingresos:,.0f}</div></div>', unsafe_allow_html=True)
    with col_g2:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #10b981;"><div class="kpi-title">Total Kilos Movilizados</div><div class="kpi-value">{total_kilos:,.1f} Kg</div></div>', unsafe_allow_html=True)
    with col_g3:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #f59e0b;"><div class="kpi-title">Total Unidades (Bultos)</div><div class="kpi-value">{total_bultos:,.0f} Btos</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # =========================================================================
    # --- SECCIÓN 2: DESEMPEÑO OPERATIVO FINANCIERO
    # =========================================================================
    st.markdown("## 📊 Desempeño Operativo por Región/Zona")
    st.markdown("### *Auditoría de Costos, Canales de Entrega y Métodos de Pago*")
    
    total_descuentos = df['Valor Descuento'].sum() if 'Valor Descuento' in df.columns else 0
    total_seguro = df['Valor Seguro'].sum() if 'Valor Seguro' in df.columns else 0
    
    c_des1, c_des2, c_des3, c_des4 = st.columns(4)
    with c_des1:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #3b82f6;"><div class="kpi-title">Monto Total Descuentos</div><div class="kpi-value">${total_descuentos:,.0f}</div></div>', unsafe_allow_html=True)
    with c_des2:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #ef4444;"><div class="kpi-title">Recaudación Seguros</div><div class="kpi-value">${total_seguro:,.0f}</div></div>', unsafe_allow_html=True)
    with c_des3:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #2563eb;"><div class="kpi-title">Participación Emitida</div><div class="kpi-value">${venta_emitida_tot:,.0f}</div></div>', unsafe_allow_html=True)
    with c_des4:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #ec4899;"><div class="kpi-title">Participación Recibida</div><div class="kpi-value">${venta_recibida_tot:,.0f}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_analisis_left, col_analisis_right = st.columns(2)
    
    with col_analisis_left:
        st.markdown('<div class="chart-block">', unsafe_allow_html=True)
        st.markdown("#### 💰 Distribución Financiera de Formas de Pago")
        if len(df) > 0 and 'Forma Pago' in df.columns:
            df_pagos_agrup = df.groupby('Forma Pago').agg(
                Monto_Total=('Valor Total', 'sum'),
                Cantidad_Transacciones=('ODT', 'count')
            ).reset_index().sort_values(by='Monto_Total', ascending=False)
            
            fig_pie_pago = px.pie(df_pagos_agrup, names='Forma Pago', values='Monto_Total', hole=0.3, color_discrete_sequence=px.colors.sequential.YlGnBu_r)
            fig_pie_pago.update_layout(margin=dict(t=20, b=20, l=10, r=10), height=320)
            st.plotly_chart(fig_pie_pago, use_container_width=True)
            
            st.dataframe(
                df_pagos_agrup.rename(columns={'Monto_Total': 'Monto Total', 'Cantidad_Transacciones': 'N° ODTs'}), 
                use_container_width=True, hide_index=True,
                column_config={"Monto Total": st.column_config.NumberColumn("Monto Total", format="$%,.0f"), "N° ODTs": st.column_config.NumberColumn("N° ODTs", format="%,.0f")}
            )
        st.markdown('</div>', unsafe_allow_html=True)

    with col_analisis_right:
        st.markdown('<div class="chart-block">', unsafe_allow_html=True)
        st.markdown("#### 🏢 Canal de Distribución Dominante (Oficina vs Domicilio)")
        if len(df) > 0 and 'Tipo Entrega' in df.columns:
            df_entrega_agrup = df.groupby('Tipo Entrega').agg(
                Monto_Total=('Valor Total', 'sum'),
                Cantidad_ODT=('ODT', 'count'),
                Kilos_Totales=('Peso', 'sum')
            ).reset_index()
            
            fig_bar_entrega = px.bar(df_entrega_agrup, x='Tipo Entrega', y='Cantidad_ODT', color='Tipo Entrega', color_discrete_map={'DOMICILIO': '#1e3a8a', 'OFICINA': '#f43f5e'}, template='plotly_white')
            fig_bar_entrega.update_layout(margin=dict(t=20, b=20, l=10, r=10), height=320, showlegend=False)
            st.plotly_chart(fig_bar_entrega, use_container_width=True)
            
            st.dataframe(
                df_entrega_agrup.rename(columns={'Monto_Total': 'Monto Total', 'Cantidad_ODT': 'N° ODTs', 'Kilos_Totales': 'Kilos'}), 
                use_container_width=True, hide_index=True,
                column_config={"Monto Total": st.column_config.NumberColumn("Monto Total", format="$%,.0f"), "N° ODTs": st.column_config.NumberColumn("N° ODTs", format="%,.0f"), "Kilos": st.column_config.NumberColumn("Kilos", format="%,.1f Kg")}
            )
        st.markdown('</div>', unsafe_allow_html=True)

    # =========================================================================
    # --- SECCIÓN 3: MAPA GEOGRÁFICO (TOTALMENTE FILTRADO POR MES)
    # =========================================================================
    st.markdown("---")
    st.markdown("## 🗺️ Mapa de Cobertura y Concentración de Carga")
    st.markdown('<div class="chart-block">', unsafe_allow_html=True)
    if 'LATITUD' in df.columns and 'LONGUITUD' in df.columns and df['LATITUD'].sum() != 0:
        df_mapa = df[df['LATITUD'] != 0].groupby(['Agencia Descripción', 'ZONA', 'REGION', 'LATITUD', 'LONGUITUD']).agg(
            Venta_Total=('Valor Total', 'sum'), Kilos_Totales=('Peso', 'sum')
        ).reset_index()
        
        fig_mapa = px.scatter_mapbox(
            df_mapa, lat="LATITUD", lon="LONGUITUD", size="Venta_Total", color="ZONA",
            hover_name="Agencia Descripción", hover_data={"REGION": True, "Venta_Total": ":$,.0f"},
            zoom=4.5, center={"lat": -37.0, "lon": -72.0}, height=500, color_discrete_sequence=px.colors.qualitative.Dark2
        )
        fig_mapa.update_layout(mapbox_style="carto-positron", margin={"r":0,"t":0,"l":0,"b":0})
        st.plotly_chart(fig_mapa, use_container_width=True)
    else:
        st.info("ℹ️ No se registran datos geográficos válidos para los filtros seleccionados.")
    st.markdown('</div>', unsafe_allow_html=True)

    # =========================================================================
    # --- SECCIÓN 4: BALANCE OPERATIVO POR AGENCIA
    # =========================================================================
    st.markdown("---")
    st.markdown("## 🔄 Balance Operativo: Emisión vs Recepción por Agencia")
    
    col_bal1, col_bal2 = st.columns(2)
    with col_bal1:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #2563eb; background-color: #f0f7ff;"><div class="kpi-title" style="color: #1e40af;">Total Ventas Emitidas (Origen)</div><div class="kpi-value" style="color: #1e3a8a;">${venta_emitida_tot:,.0f}</div></div>', unsafe_allow_html=True)
    with col_bal2:
        st.markdown(f'<div class="kpi-card" style="border-left-color: #10b981; background-color: #f0fdf4;"><div class="kpi-title" style="color: #166534;">Total Operaciones Recibidas (Destino)</div><div class="kpi-value" style="color: #14532d;">${venta_recibida_tot:,.0f}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-block">', unsafe_allow_html=True)
    if len(df) > 0:
        df_ag_bal = df.groupby(['Agencia Descripción', 'Tipo ODT'])['Valor Total'].sum().reset_index()
        top_agencias = df.groupby('Agencia Descripción')['Valor Total'].sum().sort_values(ascending=False).head(10).index
        df_ag_bal_top = df_ag_bal[df_ag_bal['Agencia Descripción'].isin(top_agencias)]

        fig_bal_top = px.bar(
            df_ag_bal_top, x="Valor Total", y="Agencia Descripción", color="Tipo ODT",
            orientation="h", barmode="group", title="<b>Top 10 Agencias Líderes en Movimiento Comercial</b>",
            color_discrete_map={"EMITIDAS": "#2563eb", "RECIBIDAS": "#10b981", "DESCUENTOS": "#ef4444"}, template="plotly_white"
        )
        fig_bal_top.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_bal_top, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # --- SECCIÓN 5: TABLA DE AUDITORÍA BASE DE DATOS ---
    st.markdown("---")
    st.markdown("### 📋 Vista Consolidada de Registros")
    with st.expander("Ver base de datos de transacciones filtradas"):
        st.dataframe(df, use_container_width=True)
else:
    st.warning("⚠️ No se encontró el archivo 'VENTA_AGENCIAS.xlsx' en el directorio.")