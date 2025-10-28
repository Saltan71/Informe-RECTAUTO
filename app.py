import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
from jinja2 import Template
import tempfile

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0)

    columnas = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    df = df.iloc[:, columnas]

    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')

    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Fecha no disponible"

    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    def crear_grafico(df, columna, titulo):
        if columna not in df.columns:
            return None
        conteo = df[columna].value_counts().reset_index()
        conteo.columns = [columna, "Cantidad"]
        conteo["Cantidad"] = conteo["Cantidad"].apply(lambda x: f"{x:,}".replace(",", "."))
        fig = px.bar(conteo, x=columna, y="Cantidad", title=titulo, text="Cantidad", color=columna, height=400)
        fig.update_traces(textposition="outside")
        return fig

    st.subheader("📈 Gráficos Generales")
    cols = st.columns(3)
    graficos = [
        ("EQUIPO", "Expedientes por equipo"),
        ("USUARIO", "Expedientes por usuario"),
        ("NOTIFICADO", "Expedientes notificados")
    ]
    for i, (col, titulo) in enumerate(graficos):
        if col in df.columns:
            fig = crear_grafico(df, col, titulo)
            if fig:
                cols[i].plotly_chart(fig, use_container_width=True)
        ("EQUIPO", "Expedientes por equipo"),
        ("USUARIO", "Expedientes por usuario"),
        ("ESTADO", "Distribución por estado"),
        ("NOTIFICADO", "Expedientes notificados"),
        if col in df.columns:
            fig = crear_grafico(df, col, titulo)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

    if "USUARIO" in df.columns:
        usuarios = df["USUARIO"].dropna().unique()
        st.subheader("👤 Informes individuales por usuario")

        with tempfile.TemporaryDirectory() as tmpdir:
            carpeta_usuarios = os.path.join(tmpdir, "usuarios")
            os.makedirs(carpeta_usuarios, exist_ok=True)

            for usuario in usuarios:
                df_user = df[df["USUARIO"] == usuario].copy()
                for col in df_user.select_dtypes(include=["float", "int"]).columns:
                    df_user[col] = df_user[col].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                for col in df_user.select_dtypes(include=["datetime"]).columns:
                    df_user[col] = df_user[col].dt.strftime("%d/%m/%Y")
                tabla_user = df_user.to_html(index=False, classes="display", border=0, justify="center")
                plantilla_user = Template("""
                <html>
                <head><meta charset="utf-8"><title>Informe {{ usuario }}</title></head>
                <body>
                    <h1>Informe individual: {{ usuario }}</h1>
                    {{ tabla|safe }}
                </body>
                </html>
                """)
                ruta_html = os.path.join(carpeta_usuarios, f"{usuario}.html")
                with open(ruta_html, "w", encoding="utf-8") as f:
                    f.write(plantilla_user.render(usuario=usuario, tabla=tabla_user))

            st.success("✅ Informes individuales generados.")
            st.download_button("📥 Descargar informes individuales (ZIP)", data=None, disabled=True)

    st.subheader("📋 Vista general de expedientes")
    df_display = df.copy()
    for col in df_display.select_dtypes(include=["float", "int"]).columns:
        df_display[col] = df_display[col].apply(lambda x: f"{x:,.0f}".replace(",", "."))
    for col in df_display.select_dtypes(include=["datetime"]).columns:
        df_display[col] = df_display[col].dt.strftime("%d/%m/%Y")
    st.dataframe(df_display, use_container_width=True)

else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
