
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
class PDF(FPDF):
    # Variables de clase para el título dinámico y los anchos de columna
    col_widths = []
    headers = []
    report_title = "Informe de Expedientes Pendientes"
    
    def header(self):
        # Título principal del informe (con el conteo de expedientes)
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, self.report_title, 0, 1, 'C')
        self.ln(5)

        # Encabezados de la tabla (se repiten en cada página)
        if self.headers:
            self.set_font("Arial", "B", 7)  # Fuente más pequeña para el encabezado (para ajustarse)
            self.set_fill_color(200, 220, 255) # Color de fondo
            
            # Altura de la celda: Aumentamos a 12mm para permitir dos líneas
            cell_height = 12 
            
            # Guardamos la posición X e Y antes de dibujar celdas multi-línea
            x_start = self.get_x()
            y_start = self.get_y()
            
            for i, header in enumerate(self.headers):
                self.set_xy(x_start, y_start)
                # Usamos multi_cell para envolver el texto si es largo
                self.multi_cell(self.col_widths[i], 6, header, 1, 'C', 1, align='T', max_line_height=3)
                x_start += self.col_widths[i]
            
            # Mover a la siguiente línea después de dibujar todos los encabezados
            self.set_xy(10, y_start + cell_height) # Movemos Y a la posición final
            self.set_font("Arial", "", 8) # Volvemos a la fuente de datos
            self.set_auto_page_break(True, margin=20) # Aseguramos que el auto-break esté activado
            

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# --- Función para generar PDF a partir de una tabla de DataFrame ---
# (Se asume que la clase PDF se define antes, como en tu código)

# Ajuste el ancho de las columnas (el ancho total de A4 horizontal es ~277mm)
# Asegúrese de que la suma de los anchos sea <= 277
# --- Función para generar PDF a partir de una tabla de DataFrame ---
# (La clase PDF debe estar definida antes)
def dataframe_to_pdf_bytes(df, title):
    """Genera un archivo PDF a partir de un DataFrame, manejando saltos de página."""
    pdf = PDF('L', 'mm', 'A4') 
    
    # 1. Ajustar el ancho de las columnas (el ancho total debe ser <= 287mm)
    # Ejemplo: 13 columnas. Ajuste estos valores si el informe final tiene más o menos.
    # [EXPEDIENTE, CLIENTE, EQUIPO, ... (13 en total después de exclusiones)]
    # Si su DataFrame final (df) tiene 13 columnas, asegúrese de que la lista tenga 13 valores.
    
    # Hemos excluido las columnas 1 y 10, así que quedan 13.
    # 287mm total / 13 columnas = ~22.07mm por columna.
    pdf.col_widths = [
        25, 30, 20, 25, 20, 25, 20, 20, 20, 20, 20, 10, 12  # Suma total: 267mm (aprox.)
    ] 

    # 2. Asignar encabezados y título dinámico
    pdf.headers = df.columns.tolist()
    pdf.report_title = title
    
    # 3. Iniciar la generación (llama a header() por primera vez)
    pdf.set_auto_page_break(True, margin=20) # Margen inferior de 20mm
    pdf.add_page()
    
    # 4. Imprimir datos de las filas
    pdf.set_font("Arial", "", 8) # Fuente para los datos

    for index, row in df.iterrows():
        # La comprobación de salto de página ya la hace pdf.add_page() y pdf.set_auto_page_break(True)
        
        # Imprimir las celdas de la fila
        for i, col_data in enumerate(row):
            text = str(col_data).replace('\n', ' ')
            # Usamos pdf.cell, la altura de la fila de datos es 6mm
            pdf.cell(pdf.col_widths[i], 6, text, 1, 0, 'L')
        pdf.ln() # Salto de línea después de cada fila

    # 5. Obtener el PDF como bytes
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

                    # 6. Obtener el número de expedientes abiertos
                    num_expedientes = len(df_pdf)
                    
                    # 6. Formato de fechas (si aplica)
                    for col in df_pdf.select_dtypes(include='datetime').columns:
                        df_pdf[col] = df_pdf[col].dt.strftime("%d/%m/%y")
                        
                    # 7. Generar el PDF
                    #nombre_usuario_sanitizado = "".join(c for c in usuario if c.isalnum() or c in ('_',)).replace(' ', '_')
                    file_name = f"{num_semana}{usuario}.pdf"
                    titulo_pdf = f"Expedientes Pendientes ({num_expedientes}) - Semana {num_semana} a {fecha_max_str} - {usuario}"
                    
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
