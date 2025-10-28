
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

archivo = st.file_uploader("📁 Sube el archivo Excel", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, engine="openpyxl")
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    if len(df.columns) > max(columnas):
        df = df.iloc[:, columnas]
    else:
        st.error("El archivo no contiene suficientes columnas.")
        st.stop()

    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    if pd.isna(fecha_max):
        st.error("No se encontró una fecha válida.")
        st.stop()

    fecha_formateada = fecha_max.strftime("%d/%m/%Y")
    dias_transcurridos = (fecha_max - datetime(2022, 11, 1)).days
    num_semana = dias_transcurridos // 7 + 1
    st.subheader(f"📅 Semana {num_semana} a {fecha_formateada}")

    # Segmentación
    equipo_sel = st.selectbox("🔍 Filtrar por EQUIPO", ["Todos"] + sorted(df["EQUIPO"].dropna().unique()))
    estado_sel = st.selectbox("🔍 Filtrar por ESTADO", ["Todos"] + sorted(df["ESTADO"].dropna().unique()))
    usuario_sel = st.selectbox("🔍 Filtrar por USUARIO", ["Todos"] + sorted(df["USUARIO"].dropna().unique()))

    df_filtrado = df.copy()
    if equipo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["EQUIPO"] == equipo_sel]
    if estado_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["ESTADO"] == estado_sel]
    if usuario_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["USUARIO"] == usuario_sel]

    # Gráficos
    st.subheader("📈 Gráficos Generales")
    def crear_grafico(df, columna, titulo):
        if columna not in df.columns:
            return None
        conteo = df[columna].value_counts().reset_index()
        conteo.columns = [columna, "Cantidad"]
        conteo["Cantidad"] = conteo["Cantidad"].apply(lambda x: f"{x:,}".replace(",", "."))
        fig = px.bar(conteo, x=columna, y="Cantidad", title=titulo, text="Cantidad", color=columna, height=400)
        fig.update_traces(textposition="outside")
        return fig

    for col, titulo in [
        ("EQUIPO", "Expedientes por equipo"),
        ("USUARIO", "Expedientes por usuario"),
        ("ESTADO", "Distribución por estado"),
        ("NOTIFICADO", "Expedientes notificados"),
    ]:
        fig = crear_grafico(df_filtrado, col, titulo)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # Tabla general
    st.subheader("📋 Vista general de expedientes")
    df_mostrar = df_filtrado.copy()
    for col in df_mostrar.select_dtypes(include="number").columns:
        df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{int(x):,}".replace(",", "."))
    for col in df_mostrar.select_dtypes(include="datetime").columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%Y")
    st.dataframe(df_mostrar, use_container_width=True)
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
