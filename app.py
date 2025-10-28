
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

# ***AJUSTA ESTA LISTA CON LOS ESTADOS QUE CONSIDERAS "PENDIENTES"***
ESTADOS_PENDIENTES = ["Abierto"]

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Generador de Informes Rectauto")

# Modifica la clase PDF para asegurar la correcta inicialización de FPDF
# --- CLASE PDF CORREGIDA (VERSION FINAL) ---
class PDF(FPDF):
    
    # 1. PASAR VARIABLES AL CONSTRUCTOR (para que existan al llamar header)
    def __init__(self, orientation, unit, format, title, headers, col_widths):
        super().__init__(orientation, unit, format)
        self.report_title = title
        self.headers = headers
        self.col_widths = col_widths
        
    def header(self):
        # Título principal del informe
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, self.report_title, 0, 1, 'C')
        self.ln(5)

        # 2. VERIFICACIÓN CRÍTICA: Solo dibuja si las variables tienen datos
        if self.headers and self.col_widths: 
            self.set_font("Arial", "B", 7)
            self.set_fill_color(200, 220, 255) 
            
            cell_height = 16 
            x_start = self.get_x()
            y_start = self.get_y()
            
            for i, header in enumerate(self.headers):
                self.set_xy(x_start, y_start)
                # La lista self.col_widths[i] ya está garantizada
                self.multi_cell(self.col_widths[i], 4, header, 1, 'C', 1, align='T', max_line_height=4)
                x_start += self.col_widths[i]
            
            self.set_xy(10, y_start + cell_height) 
            self.set_font("Arial", "", 8)
            self.set_auto_page_break(True, margin=20)
            
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
        
# --- Función para generar PDF a partir de una tabla de DataFrame ---
# (Se asume que la clase PDF se define antes, como en tu código)

# Ajuste el ancho de las columnas (el ancho total de A4 horizontal es ~277mm)
# Asegúrese de que la suma de los anchos sea <= 277
# --- FUNCIÓN dataframe_to_pdf_bytes CORREGIDA ---
def dataframe_to_pdf_bytes(df, title):
    """Genera un archivo PDF a partir de un DataFrame, manejando saltos de página y formatos."""
    
    # 1. Preparar las variables antes de crear el objeto
    headers = df.columns.tolist()
    
    # 2. Definir los anchos de columna (EJEMPLO para 13 columnas. AJÚSTALOS SI NECESARIO)
    col_widths = [25, 30, 20, 25, 20, 25, 20, 20, 20, 20, 20, 10, 12] 
    
    if len(col_widths) != len(df.columns):
        # Manejo de error o ajuste dinámico
        st.error(f"Error interno: La tabla final tiene {len(df.columns)} columnas, pero se asignaron {len(col_widths)} anchos. El PDF podría estar mal formateado.")
        # Aquí puedes añadir una lógica para asignar anchos de forma equitativa si el conteo falla.
        # ancho_dinamico = 280 / len(df.columns)
        # col_widths = [ancho_dinamico] * len(df.columns)
    
    # 3. INSTANCIACIÓN CORREGIDA: Pasar las variables al constructor
    pdf = PDF(
        orientation='L', 
        unit='mm', 
        format='A4', 
        title=title, 
        headers=headers, 
        col_widths=col_widths
    ) 
    
    # 4. Iniciar la generación (llama a header(), que ahora tiene las variables)
    pdf.set_auto_page_break(True, margin=20) 
    pdf.add_page() # ¡Ahora el header() debe funcionar!
    
    # 5. Imprimir datos de las filas
    pdf.set_font("Arial", "", 8)
    data_row_height = 8 

    for index, row in df.iterrows():
        for i, col_data in enumerate(row):
            text = str(col_data).replace('\n', ' ')
            pdf.cell(pdf.col_widths[i], data_row_height, text, 1, 0, 'L') 
        pdf.ln()

    # 6. Obtener el PDF como bytes
    pdf_output = pdf.output(dest='B')
    
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
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%y")
    st.dataframe(df_mostrar, use_container_width=True)

    # --- DESCARGA DE INFORMES EN EXCEL Y PDF ---
    
    st.markdown("---")
    st.header("Descarga de Informes")

    # B. Generación de Informes PDF por Usuario (ZIP)
    st.subheader("Generar Informes PDF Pendientes por Usuario")

    # Usamos el DataFrame completo (df) para la selección inicial de pendientes
    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()

    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF Pendientes"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()
            
                # --- PREPARACIÓN Y SELECCIÓN DE COLUMNAS PARA EL PDF ---
            
                # 1. Identificar las columnas a excluir y redondear
                # Indices basados en el DataFrame de 15 columnas (df_pendientes):
                # Columna 4 (índice 4) -> Redondear
                # Columna 1 (índice 1) -> Excluir
                # Columna 10 (índice 10) -> Excluir
            
                # Creamos una lista con todos los índices de columna
                indices_a_incluir = list(range(df_pendientes.shape[1])) 
            
                # Identificamos los índices a excluir
                indices_a_excluir = {1, 10} 
            
                # Filtramos para obtener solo los índices que queremos
                indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
            
                # 2. Creamos la lista final de nombres de columnas
                NOMBRES_COLUMNAS_PDF = df_pendientes.columns[indices_finales].tolist()

            # -----------------------------------------------------------
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for usuario in usuarios_pendientes:
                    # 3. Filtrar por usuario (obteniendo TODAS las columnas pendientes)
                    df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
                    
                    # 4. Redondear la Columna 4 (índice 4)
                    indice_columna_a_redondear = 4
                    nombre_columna_a_redondear = df_user.columns[indice_columna_a_redondear]
                    
                    if nombre_columna_a_redondear in df_user.columns:
                         # Forzar a numérico, redondear y convertir a entero (si no es nulo)
                        df_user[nombre_columna_a_redondear] = pd.to_numeric(
                            df_user[nombre_columna_a_redondear], errors='coerce'
                        ).fillna(0).round(0).astype(int)
                    
                    # 5. Seleccionar SÓLO las columnas deseadas para el informe final
                    df_pdf = df_user[NOMBRES_COLUMNAS_PDF].copy()
                    
                    # 6. Formato de fechas (si aplica)
                    for col in df_pdf.select_dtypes(include='datetime').columns:
                        df_pdf[col] = df_pdf[col].dt.strftime("%d/%m/%y")
                    
                    # 6. Obtener el número de expedientes abiertos
                    num_expedientes = len(df_pdf)
                    
                    # 7. Generar el PDF
                    #nombre_usuario_sanitizado = "".join(c for c in usuario if c.isalnum() or c in ('_',)).replace(' ', '_')
                    file_name = f"{num_semana}{usuario}.pdf"
                    titulo_pdf = f"{usuario} - Semana {num_semana} a {fecha_max_str} - Expedientes Pendientes ({num_expedientes})"
                    
                    # Llamada a la función de generación PDF (que maneja múltiples páginas)
                    pdf_data = dataframe_to_pdf_bytes(df_pdf, titulo_pdf)
                    
                    # 8. Añadir al ZIP
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
