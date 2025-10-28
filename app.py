
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import zipfile
from fpdf import FPDF
import matplotlib.pyplot as plt
import os # Aunque no se usa directamente en este flujo, es buena práctica mantenerla

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

# Función para generar PDF a partir de una tabla de DataFrame (para informe)
# Se usa FPDF y Matplotlib para renderizar la tabla de forma estable en PDF.
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'Informe de Expedientes Pendientes', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def dataframe_to_pdf_bytes(df, title):
    """Genera un archivo PDF a partir de un DataFrame con un título."""
    pdf = PDF('L', 'mm', 'A4') # 'L' para formato horizontal A4
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    # Convertir el DataFrame a imagen usando matplotlib
    # Esto es necesario porque FPDF puro no maneja bien tablas grandes de Pandas
    fig, ax = plt.subplots(figsize=(28/2.54, 18/2.54)) # Tamaño ajustado para A4 horizontal
    ax.axis('tight')
    ax.axis('off')
    
    # Renderizar la tabla
    tabla = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='left')
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(7) 
    tabla.scale(1, 1.1) 

    # Guardar la imagen de la tabla en un buffer
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight')
    plt.close(fig)
    img_buffer.seek(0)
    
    # Añadir la imagen al PDF (ajuste el tamaño para que quepa)
    pdf.image(img_buffer, x=5, y=25, w=287) # w=287 es casi todo el ancho A4
    
    # Guardar el PDF en bytes
    pdf_output = pdf.output(dest='S').encode('latin-1')
    return pdf_output

# --- PROCESAMIENTO DE ARCHIVO ---

archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(
        archivo,
        sheet_name=HOJA,
        header=0,
        engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd"
    )
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
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
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

    # Mostrar los tres gráficos en paralelo
    columnas_graficos = st.columns(3)
    graficos = [
        ("EQUIPO", "Expedientes por equipo"),
        ("USUARIO", "Expedientes por usuario"),
        ("ESTADO", "Distribución por estado")
    ]
    for i, (col, titulo) in enumerate(graficos):
        if col in df_filtrado.columns:
            fig = crear_grafico(df_filtrado, col, titulo)
            if fig:
                columnas_graficos[i].plotly_chart(fig, use_container_width=True)

    # Mostrar el gráfico de NOTIFICADO debajo
    if "NOTIFICADO" in df_filtrado.columns:
        fig = crear_grafico(df_filtrado, "NOTIFICADO", "Expedientes notificados")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("📋 Vista general de expedientes")
    df_mostrar = df_filtrado.copy()
    for col in df_mostrar.select_dtypes(include='number').columns:
        df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%Y")
    st.dataframe(df_mostrar, use_container_width=True)

# --- DESCARGA DE INFORMES EN EXCEL Y PDF ---
    
    st.markdown("---")
    st.header("Descarga de Informes")

    # B. Generación de Informes PDF por Usuario (ZIP)
    st.subheader("Generar Informes PDF Pendientes por Usuario")

    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()
    
    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF Pendientes"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()
                
                # Columnas a incluir en los informes PDF
                COLUMNAS_CLAVE = ["EXPEDIENTE", "CLIENTE", "EQUIPO", "ESTADO", columna_fecha]
                
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for usuario in usuarios_pendientes:
                        df_user = df_pendientes[df_pendientes["USUARIO"] == usuario][COLUMNAS_CLAVE].copy()
                        
                        # Formato de fechas para el PDF
                        for col in df_user.select_dtypes(include='datetime').columns:
                            df_user[col] = df_user[col].dt.strftime("%d/%m/%Y")
                            
                        # Sanear nombre de archivo
                        nombre_usuario_sanitizado = "".join(c for c in usuario if c.isalnum() or c in ('_',)).replace(' ', '_')
                        file_name = f"Semana_{num_semana}_{nombre_usuario_sanitizado}_PENDIENTES.pdf"
                        
                        # Generar el PDF
                        titulo_pdf = f"Expedientes Pendientes - Semana {num_semana} - {usuario}"
                        pdf_data = dataframe_to_pdf_bytes(df_user, titulo_pdf)
                        
                        # Añadir al ZIP
                        zip_file.writestr(file_name, pdf_data)

            # Botón de descarga del ZIP
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
