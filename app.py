import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
from jinja2 import Template
import tempfile

# === CONFIGURACIÓN GENERAL ===
FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

# === SUBIDA DE ARCHIVO ===
archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, engine="openpyxl")

    # Seleccionar columnas específicas
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]

    # Convertir columna de fecha
    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')

    # Calcular semana
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1

    st.subheader(f"📅 Semana detectada: {num_semana}")

    # === GRÁFICOS ===
    def crear_grafico(df, columna, titulo):
        if columna not in df.columns:
            return None
        conteo = df[columna].value_counts().reset_index()
        conteo.columns = [columna, "Cantidad"]
        fig = px.bar(conteo, x=columna, y="Cantidad", title=titulo, text="Cantidad", color=columna, height=400)
        fig.update_traces(textposition="outside")
        return fig

    st.subheader("📈 Gráficos Generales")
    for col, titulo in [
        ("EQUIPO", "Expedientes por equipo"),
        ("USUARIO", "Expedientes por usuario"),
        ("ESTADO", "Distribución por estado"),
        ("NOTIFICADO", "Expedientes notificados"),
    ]:
        if col in df.columns:
            fig = crear_grafico(df, col, titulo)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

    # === INFORMES POR USUARIO ===
    if "USUARIO" in df.columns:
        usuarios = df["USUARIO"].dropna().unique()
        st.subheader("👤 Informes individuales por usuario")

        with tempfile.TemporaryDirectory() as tmpdir:
            carpeta_usuarios = os.path.join(tmpdir, "usuarios")
            os.makedirs(carpeta_usuarios, exist_ok=True)

            for usuario in usuarios:
                df_user = df[df["USUARIO"] == usuario]
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

    # === TABLA GENERAL ===
    st.subheader("📋 Vista general de expedientes")
    st.dataframe(df, use_container_width=True)

else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
