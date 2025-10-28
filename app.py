import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"

archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, engine="openpyxl")
    df.columns = [col.upper() for col in df.columns]

    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]

    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notnull(fecha_max) else "Sin fecha"

    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    if "ESTADO" in df.columns:
        df = df[df["ESTADO"] == "Abierto"]

    equipo_sel = st.selectbox("🔍 Filtrar por EQUIPO", ["Todos"] + sorted(df["EQUIPO"].dropna().unique()) if "EQUIPO" in df.columns else ["Todos"])
    estado_sel = st.selectbox("🔍 Filtrar por ESTADO", ["Todos"] + sorted(df["ESTADO"].dropna().unique()) if "ESTADO" in df.columns else ["Todos"])
    usuario_sel = st.selectbox("🔍 Filtrar por USUARIO", ["Todos"] + sorted(df["USUARIO"].dropna().unique()) if "USUARIO" in df.columns else ["Todos"])

    if equipo_sel != "Todos":
        df = df[df["EQUIPO"] == equipo_sel]
    if estado_sel != "Todos":
        df = df[df["ESTADO"] == estado_sel]
    if usuario_sel != "Todos":
        df = df[df["USUARIO"] == usuario_sel]

    def crear_grafico(df, columna, titulo):
        if columna not in df.columns:
            return None
        conteo = df[columna].value_counts().reset_index()
        conteo.columns = [columna, "Cantidad"]
        fig = px.bar(conteo, y=columna, x="Cantidad", orientation="h", title=titulo, text="Cantidad", color=columna, height=400)
        fig.update_traces(textposition="outside")
        fig.update_layout(xaxis_tickformat=",d")
        return fig

    st.subheader("📈 Gráficos Generales")
    col1, col2, col3 = st.columns(3)
    with col1:
        fig1 = crear_grafico(df, "EQUIPO", "Expedientes por equipo")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
    with col2:
        fig2 = crear_grafico(df, "USUARIO", "Expedientes por usuario")
        if fig2: st.plotly_chart(fig2, use_container_width=True)
    with col3:
        fig3 = crear_grafico(df, "ESTADO", "Distribución por estado")
        if fig3: st.plotly_chart(fig3, use_container_width=True)

    st.subheader("📋 Vista general de expedientes")
    df_display = df.copy()
    for col in df_display.select_dtypes(include=["datetime64[ns]"]).columns:
        df_display[col] = df_display[col].dt.strftime("%d/%m/%Y")
    for col in df_display.select_dtypes(include=["int", "float"]).columns:
        df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}".replace(",", ".") if pd.notnull(x) else "")
    st.dataframe(df_display, use_container_width=True)
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
