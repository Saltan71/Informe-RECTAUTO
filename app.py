import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import zipfile
from fpdf import FPDF
import matplotlib.pyplot as plt
import os

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"
ESTADOS_PENDIENTES = ["Abierto"]

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 8)
        self.cell(0, 10, 'Informe de Expedientes Pendientes', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def dataframe_to_pdf_bytes(df, title):
    """Genera un PDF desde un DataFrame, con encabezados homogéneos, color de fondo y repetidos."""
    pdf = PDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", "B", 8)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    pdf.set_font("Arial", "B", 5)
    col_widths = [43, 14, 14, 8, 24, 14, 14, 24, 14, 40, 24, 14, 26]
    df_mostrar_pdf = df.iloc[:, :len(col_widths)]

    def imprimir_encabezados():
        pdf.set_font("Arial", "B", 5)
        pdf.set_fill_color(200, 220, 255)

        # 1️⃣ Medir alturas individuales
        alturas = []
        for i, header in enumerate(df_mostrar_pdf.columns):
            x, y = pdf.get_x(), pdf.get_y()
            pdf.multi_cell(col_widths[i], 3, header, border=0)
            altura = pdf.get_y() - y
            alturas.append(altura)
            pdf.set_xy(x + col_widths[i], y)
        altura_max = max(alturas)

        # 2️⃣ Dibujar celdas con fondo y texto centrado
        y_inicio = pdf.get_y()
        for i, header in enumerate(df_mostrar_pdf.columns):
            x = pdf.get_x()

            # fondo de color
            pdf.set_fill_color(200, 220, 255)
            pdf.rect(x, y_inicio, col_widths[i], altura_max, 'F')  # solo relleno

            # borde fino (opcional)
            pdf.rect(x, y_inicio, col_widths[i], altura_max, 'D')

            # texto centrado verticalmente
            pdf.set_xy(x, y_inicio + (altura_max - alturas[i]) / 2)
            pdf.multi_cell(col_widths[i], 3, header, border=0, align='C')

            pdf.set_xy(x + col_widths[i], y_inicio)
        pdf.ln(altura_max)

    # --- PRIMER ENCABEZADO ---
    imprimir_encabezados()

    # --- DATOS ---
    pdf.set_font("Arial", "", 7)
    for _, row in df_mostrar_pdf.iterrows():
        if pdf.get_y() + 6 > 190:
            pdf.add_page()
            imprimir_encabezados()

        for i, col_data in enumerate(row):
            text = str(col_data).replace('\n', ' ')
            pdf.cell(col_widths[i], 6, text, 1, 0, 'L')
        pdf.ln()

    pdf_output = pdf.output(dest='B')
    return pdf_output

archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd")
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]
    st.session_state["df"] = df
elif "df" in st.session_state:
    df = st.session_state["df"]
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
    st.stop()

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd")
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]

    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

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

    def crear_grafico(df, columna, titulo):
        if columna not in df.columns:
            return None
        conteo = df[columna].value_counts().reset_index()
        conteo.columns = [columna, "Cantidad"]
        fig = px.bar(conteo, y=columna, x="Cantidad", title=titulo, text="Cantidad", color=columna, height=400)
        fig.update_traces(texttemplate='%{text:,}', textposition="auto")
        return fig

    st.subheader("📈 Gráficos Generales")
    columnas_graficos = st.columns(3)
    graficos = [("EQUIPO", "Expedientes por equipo"), ("USUARIO", "Expedientes por usuario"), ("ESTADO", "Distribución por estado")]

    for i, (col, titulo) in enumerate(graficos):
        if col in df_filtrado.columns:
            fig = crear_grafico(df_filtrado, col, titulo)
            if fig:
                columnas_graficos[i].plotly_chart(fig, use_container_width=True)

    if "NOTIFICADO" in df_filtrado.columns:
        fig = crear_grafico(df_filtrado, "NOTIFICADO", "Expedientes notificados")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("📋 Vista general de expedientes")
    df_mostrar = df_filtrado.copy()
    for col in df_mostrar.select_dtypes(include='number').columns:
        df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%y")
    st.dataframe(df_mostrar, use_container_width=True)

    st.markdown("---")
    st.header("Descarga de Informes")
    st.subheader("Generar Informes PDF Pendientes por Usuario")

    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()

    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF Pendientes"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()
                indices_a_incluir = list(range(df_pendientes.shape[1]))
                indices_a_excluir = {1, 10}
                indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
                NOMBRES_COLUMNAS_PDF = df_pendientes.columns[indices_finales].tolist()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for usuario in usuarios_pendientes:
                    df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
                    indice_columna_a_redondear = 4
                    nombre_columna_a_redondear = df_user.columns[indice_columna_a_redondear]

                    if nombre_columna_a_redondear in df_user.columns:
                        df_user[nombre_columna_a_redondear] = pd.to_numeric(df_user[nombre_columna_a_redondear], errors='coerce').fillna(0).round(0).astype(int)

                    df_pdf = df_user[NOMBRES_COLUMNAS_PDF].copy()
                    for col in df_pdf.select_dtypes(include='datetime').columns:
                        df_pdf[col] = df_pdf[col].dt.strftime("%d/%m/%y")

                    num_expedientes = len(df_pdf)
                    file_name = f"{num_semana}{usuario}.pdf"
                    titulo_pdf = f"{usuario} - Semana {num_semana} a {fecha_max_str} - Expedientes Pendientes ({num_expedientes})"
                    pdf_data = dataframe_to_pdf_bytes(df_pdf, titulo_pdf)
                    zip_file.writestr(file_name, pdf_data)

            zip_buffer.seek(0)
            zip_file_name = f"Informes_Pendientes_Semana_{num_semana}.zip"
            st.download_button(
                label=f"⬇️ Descargar {len(usuarios_pendientes)} Informes PDF (ZIP)",
                data=zip_buffer.read(),
                file_name=zip_file_name,
                mime="application/zip",
                help="Descarga todos los informes PDF listos para subirlos a SharePoint.",
                key='pdf_download_button'
            )
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
