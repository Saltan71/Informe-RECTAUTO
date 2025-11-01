import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta
import io
import zipfile
from fpdf import FPDF
import matplotlib.pyplot as plt
import os
import hashlib

# Constantes
FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"
ESTADOS_PENDIENTES = ["Abierto"]
CACHE_TTL = 7200  # 2 horas en segundos

st.set_page_config(page_title="Informe Rectauto", layout="wide")

# CSS con tamaños reducidos para títulos
st.markdown("""
<style>
    [data-testid="stSidebar"] {
    background-color: #007933 !important;
    }
    
    .main .block-container {
        background-color: #C4DDCA !important;
        padding: 2rem;
        border-radius: 10px;
    }
    
    .stApp {
        background-color: #92C88F !important;
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    .main .stMarkdown, .main h1, .main h2, .main h3 {
        color: #333333 !important;
    }
    
    /* NUEVO: Reducir tamaño de títulos */
    h1 {
        font-size: 24px !important;
    }
    h2 {
        font-size: 18px !important;
    }
    h3 {
        font-size: 16px !important;
    }
    .stTitle {
        font-size: 24px !important;
    }
</style>
""", unsafe_allow_html=True)

# Título principal con tamaño reducido
st.markdown("<h1 style='font-size: 24px;'>📊 Seguimiento Equipo Regional RECTAUTO</h1>", unsafe_allow_html=True)

# Clase PDF (se mantiene igual)
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 8)
        self.cell(0, 10, 'Informe de Expedientes Pendientes', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# Funciones optimizadas con cache
@st.cache_data(ttl=CACHE_TTL, show_spinner="Procesando archivo Excel...")
def cargar_y_procesar_datos(archivo):
    """Carga y procesa el archivo Excel con cache de 2 horas"""
    df = pd.read_excel(
        archivo, 
        sheet_name=HOJA, 
        header=0, 
        thousands='.', 
        decimal=',', 
        engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd"
    )
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 6, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]
    return df

@st.cache_data(ttl=CACHE_TTL)
def dataframe_to_pdf_bytes(df, title):
    """Genera un PDF desde un DataFrame con cache"""
    pdf = PDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", "B", 6)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    col_widths = [43, 14, 14, 8, 24, 14, 14, 24, 14, 40, 24, 14, 26]
    df_mostrar_pdf = df.iloc[:, :len(col_widths)]
    ALTURA_ENCABEZADO = 11

    def imprimir_encabezados():
        pdf.set_font("Arial", "B", 5)
        pdf.set_fill_color(200, 220, 255)
        y_inicio = pdf.get_y()
        
        for i, header in enumerate(df_mostrar_pdf.columns):
            x = pdf.get_x()
            y = pdf.get_y()
            pdf.cell(col_widths[i], ALTURA_ENCABEZADO, "", 1, 0, 'C', True)
            pdf.set_xy(x, y)
            
            texto = str(header)
            ancho_texto = pdf.get_string_width(texto)
            
            if ancho_texto <= col_widths[i] - 2:
                altura_texto = 3
                y_pos = y + (ALTURA_ENCABEZADO - altura_texto) / 2
                pdf.set_xy(x, y_pos)
                pdf.cell(col_widths[i], altura_texto, texto, 0, 0, 'C')
            else:
                pdf.set_xy(x, y + 1)
                pdf.multi_cell(col_widths[i], 2.5, texto, 0, 'C')
            
            pdf.set_xy(x + col_widths[i], y)
        
        pdf.set_xy(pdf.l_margin, y_inicio + ALTURA_ENCABEZADO)

    imprimir_encabezados()

    pdf.set_font("Arial", "", 7)
    for _, row in df_mostrar_pdf.iterrows():
        if pdf.get_y() + 6 > 190:
            pdf.add_page()
            imprimir_encabezados()

        for i, col_data in enumerate(row):
            text = str(col_data).replace("\n", " ")
            pdf.cell(col_widths[i], 6, text, 1, 0, 'L')
        pdf.ln()

    pdf_output = pdf.output(dest='B')
    return pdf_output

def obtener_hash_archivo(archivo):
    """Genera un hash único del archivo para detectar cambios"""
    archivo.seek(0)
    file_hash = hashlib.md5(archivo.read()).hexdigest()
    archivo.seek(0)
    return file_hash

# Logo
st.sidebar.image("Logo Atrian.png", width=260)

# Botón para limpiar cache
with st.sidebar:
    st.markdown("---")
    if st.button("🔄 Limpiar cache", help="Limpiar toda la cache y recargar"):
        st.cache_data.clear()
        # Limpiar session state excepto lo esencial
        keys_to_keep = ['archivos_hash', 'df_combinado', 'kpi_semana_index']
        for key in list(st.session_state.keys()):
            if key not in keys_to_keep:
                del st.session_state[key]
        st.success("Cache limpiada correctamente")
        st.rerun()

# NUEVA SECCIÓN: CARGA DE TRES ARCHIVOS EN COLUMNAS PARALELAS
st.markdown("---")
st.markdown("<h2 style='font-size: 18px;'>📁 Carga de Archivos</h2>", unsafe_allow_html=True)

# Crear tres columnas
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div style="text-align: center; background-color: #1f77b4; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">📊 RECTAUTO</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_rectauto = st.file_uploader(
        "Archivo principal de expedientes",
        type=["xlsx", "xls"],
        key="rectauto_upload",
        label_visibility="collapsed",
        help="Sube el archivo principal RECTAUTO"
    )
    if archivo_rectauto:
        st.success(f"✅ {archivo_rectauto.name}")
    else:
        st.info("⏳ Esperando archivo RECTAUTO")

with col2:
    st.markdown('<div style="text-align: center; background-color: #ff7f0e; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">📨 NOTIFICA</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_notifica = st.file_uploader(
        "Archivo de notificaciones",
        type=["xlsx", "xls"],
        key="notifica_upload",
        label_visibility="collapsed",
        help="Sube el archivo NOTIFICA"
    )
    if archivo_notifica:
        st.success(f"✅ {archivo_notifica.name}")
    else:
        st.info("⏳ Esperando archivo NOTIFICA")

with col3:
    st.markdown('<div style="text-align: center; background-color: #2ca02c; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">⚡ TRIAJE</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_triaje = st.file_uploader(
        "Archivo de triaje",
        type=["xlsx", "xls"],
        key="triaje_upload",
        label_visibility="collapsed",
        help="Sube el archivo TRIAJE"
    )
    if archivo_triaje:
        st.success(f"✅ {archivo_triaje.name}")
    else:
        st.info("⏳ Esperando archivo TRIAJE")

# Estado de carga
st.markdown("---")
st.markdown("<h2 style='font-size: 18px;'>📋 Estado de Carga</h2>", unsafe_allow_html=True)

# Mostrar estado con métricas
estado_col1, estado_col2, estado_col3, estado_col4 = st.columns(4)

with estado_col1:
    rectauto_status = "✅ Cargado" if archivo_rectauto else "❌ Pendiente"
    st.metric("RECTAUTO", rectauto_status)

with estado_col2:
    notifica_status = "✅ Cargado" if archivo_notifica else "❌ Pendiente"
    st.metric("NOTIFICA", notifica_status)

with estado_col3:
    triaje_status = "✅ Cargado" if archivo_triaje else "❌ Pendiente"
    st.metric("TRIAJE", triaje_status)

with estado_col4:
    archivos_cargados = sum([1 for f in [archivo_rectauto, archivo_notifica, archivo_triaje] if f])
    st.metric("Total Cargados", f"{archivos_cargados}/3")

# Función para combinar los archivos por RUE
@st.cache_data(ttl=CACHE_TTL, show_spinner="Combinando archivos...")
def combinar_archivos(_archivo_rectauto, _archivo_notifica=None, _archivo_triaje=None):
    """Combina los archivos por el campo RUE y mantiene las columnas originales de RECTAUTO"""
    
    # Cargar archivo RECTAUTO principal
    df_rectauto = cargar_y_procesar_datos(_archivo_rectauto)
    
    # Lista para almacenar dataframes adicionales
    dataframes_adicionales = []
    
    # Cargar NOTIFICA si está disponible
    if _archivo_notifica:
        try:
            df_notifica = pd.read_excel(_archivo_notifica)
            df_notifica.columns = [col.upper() for col in df_notifica.columns]
            if 'RUE' in df_notifica.columns:
                dataframes_adicionales.append(('NOTIFICA', df_notifica))
            else:
                st.warning("⚠️ NOTIFICA no tiene columna RUE, no se puede combinar")
        except Exception as e:
            st.error(f"❌ Error cargando NOTIFICA: {e}")
    
    # Cargar TRIAJE si está disponible
    if _archivo_triaje:
        try:
            df_triaje = pd.read_excel(_archivo_triaje)
            df_triaje.columns = [col.upper() for col in df_triaje.columns]
            if 'RUE' in df_triaje.columns:
                dataframes_adicionales.append(('TRIAJE', df_triaje))
            else:
                st.warning("⚠️ TRIAJE no tiene columna RUE, no se puede combinar")
        except Exception as e:
            st.error(f"❌ Error cargando TRIAJE: {e}")
    
    # Combinar todos los dataframes por RUE
    df_combinado = df_rectauto.copy()
    
    for nombre, df_adicional in dataframes_adicionales:
        # Hacer merge manteniendo todas las filas de RECTAUTO
        df_combinado = pd.merge(
            df_combinado, 
            df_adicional, 
            on='RUE', 
            how='left',
            suffixes=('', f'_{nombre}')
        )
    
    return df_combinado

# Procesar archivos cuando estén listos
if archivo_rectauto:
    # Verificar si los archivos han cambiado
    archivos_actuales = {
        'rectauto': obtener_hash_archivo(archivo_rectauto),
        'notifica': obtener_hash_archivo(archivo_notifica) if archivo_notifica else None,
        'triaje': obtener_hash_archivo(archivo_triaje) if archivo_triaje else None
    }
    
    archivos_guardados = st.session_state.get("archivos_hash", {})
    
    # Si los archivos son nuevos o cambiaron, procesar
    if (archivos_actuales != archivos_guardados or 
        "df_combinado" not in st.session_state):
        
        with st.spinner("🔄 Combinando archivos por RUE..."):
            try:
                df_combinado = combinar_archivos(archivo_rectauto, archivo_notifica, archivo_triaje)
                
                # Guardar en session_state
                st.session_state["df_combinado"] = df_combinado
                st.session_state["archivos_hash"] = archivos_actuales
                
                st.success(f"✅ Archivos combinados correctamente")
                st.info(f"📊 Dataset final: {len(df_combinado)} registros, {len(df_combinado.columns)} columnas")
                
            except Exception as e:
                st.error(f"❌ Error combinando archivos: {e}")
                # Fallback: usar solo RECTAUTO
                with st.spinner("🔄 Cargando solo RECTAUTO..."):
                    df_rectauto = cargar_y_procesar_datos(archivo_rectauto)
                    st.session_state["df_combinado"] = df_rectauto
                    st.session_state["archivos_hash"] = archivos_actuales
                    st.warning("⚠️ Usando solo archivo RECTAUTO debido a errores en combinación")
    
    else:
        # Usar datos cacheados
        df_combinado = st.session_state["df_combinado"]
        st.sidebar.success("✅ Usando datos combinados cacheados")

elif "df_combinado" in st.session_state:
    # Usar datos previamente cargados
    df_combinado = st.session_state["df_combinado"]
    st.sidebar.info("📊 Datos combinados cargados desde cache")
else:
    st.warning("⚠️ **Carga obligatoria:** Sube al menos el archivo RECTAUTO para continuar")
    st.info("💡 **Archivos opcionales:** NOTIFICA y TRIAJE enriquecerán el análisis")
    st.stop()

# Mostrar información del dataset combinado
with st.expander("📊 Información del Dataset Combinado"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Registros", f"{len(df_combinado):,}".replace(",", "."))
    
    with col2:
        st.metric("Total Columnas", len(df_combinado.columns))
    
    with col3:
        archivos_usados = 1
        if archivo_notifica:
            try:
                df_temp = pd.read_excel(archivo_notifica)
                if 'RUE' in df_temp.columns:
                    archivos_usados += 1
            except:
                pass
        if archivo_triaje:
            try:
                df_temp = pd.read_excel(archivo_triaje)
                if 'RUE' in df_temp.columns:
                    archivos_usados += 1
            except:
                pass
        st.metric("Archivos Usados", f"{archivos_usados}/3")
    
    # Mostrar primeras filas
    st.write("**Vista previa del dataset combinado:**")
    st.dataframe(df_combinado.head(3), use_container_width=True)

# Menú principal
menu = ["Principal", "Indicadores clave (KPI)"]
eleccion = st.sidebar.selectbox("Menú", menu)

# Función para gráficos dinámicos con cache
@st.cache_data(ttl=300)  # 5 minutos para gráficos
def crear_grafico_dinamico(_conteo, columna, titulo):
    """Crea gráficos dinámicos con cache"""
    if _conteo.empty:
        return None
    
    fig = px.bar(_conteo, y=columna, x="Cantidad", title=titulo, text="Cantidad", 
                 color=columna, height=400)
    fig.update_traces(texttemplate='%{text:,}', textposition="auto")
    return fig

if eleccion == "Principal":
    # Usar df_combinado en lugar de df
    df = df_combinado
    
    # Inicializar filtros en session_state si no existen
    if 'filtro_estado' not in st.session_state:
        st.session_state.filtro_estado = ['Abierto'] if 'Abierto' in df['ESTADO'].values else []
    
    if 'filtro_equipo' not in st.session_state:
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())
    
    if 'filtro_usuario' not in st.session_state:
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())
    
    # Calcular semana actual (solo una vez)
    if 'num_semana' not in st.session_state or 'fecha_max_str' not in st.session_state:
        columna_fecha = df.columns[11]
        df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
        fecha_max = df[columna_fecha].max()
        dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
        num_semana = dias_transcurridos // 7 + 1
        fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
        st.session_state.num_semana = num_semana
        st.session_state.fecha_max_str = fecha_max_str
    else:
        num_semana = st.session_state.num_semana
        fecha_max_str = st.session_state.fecha_max_str
    
    st.markdown(f"<h2 style='font-size: 18px;'>📅 Semana {num_semana} a {fecha_max_str}</h2>", unsafe_allow_html=True)

    # Sidebar para filtros
    st.sidebar.header("Filtros")

    if st.sidebar.button("Mostrar todos / Resetear filtros"):
        st.session_state.filtro_estado = sorted(df['ESTADO'].dropna().unique())
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())
        st.rerun()

    opciones_estado = sorted(df['ESTADO'].dropna().unique())
    opciones_equipo = sorted(df['EQUIPO'].dropna().unique())
    opciones_usuario = sorted(df['USUARIO'].dropna().unique())

    estado_sel = st.sidebar.multiselect(
        "Selecciona Estado:",
        options=opciones_estado,
        default=st.session_state.filtro_estado,
        key='filtro_estado_selector'
    )

    equipo_sel = st.sidebar.multiselect(
        "Selecciona Equipo:",
        options=opciones_equipo,
        default=st.session_state.filtro_equipo,
        key='filtro_equipo_selector'
    )

    usuario_sel = st.sidebar.multiselect(
        "Selecciona Usuario:",
        options=opciones_usuario,
        default=st.session_state.filtro_usuario,
        key='filtro_usuario_selector'
    )

    # Actualizar session_state con los valores seleccionados
    st.session_state.filtro_estado = estado_sel
    st.session_state.filtro_equipo = equipo_sel
    st.session_state.filtro_usuario = usuario_sel

    # Aplicar filtros al DataFrame
    df_filtrado = df.copy()

    if estado_sel:
        df_filtrado = df_filtrado[df_filtrado['ESTADO'].isin(estado_sel)]

    if equipo_sel:
        df_filtrado = df_filtrado[df_filtrado['EQUIPO'].isin(equipo_sel)]

    if usuario_sel:
        df_filtrado = df_filtrado[df_filtrado['USUARIO'].isin(usuario_sel)]

    # Mostrar filtros activos
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros activos")
    if estado_sel:
        st.sidebar.write(f"Estados: {', '.join(estado_sel)}")
    if equipo_sel:
        st.sidebar.write(f"Equipos: {len(equipo_sel)} seleccionados")
    if usuario_sel:
        st.sidebar.write(f"Usuarios: {len(usuario_sel)} seleccionados")

    # Gráficos Generales - CON CACHE
    st.markdown("<h2 style='font-size: 18px;'>📈 Gráficos Generales</h2>", unsafe_allow_html=True)
    columnas_graficos = st.columns(3)
    graficos = [("EQUIPO", "Expedientes por equipo"), 
                ("USUARIO", "Expedientes por usuario"), 
                ("ESTADO", "Distribución por estado")]

    for i, (col, titulo) in enumerate(graficos):
        if col in df_filtrado.columns:
            # Calcular el conteo actual
            conteo_actual = df_filtrado[col].value_counts().reset_index()
            conteo_actual.columns = [col, "Cantidad"]
            
            # Crear gráfico con cache
            fig = crear_grafico_dinamico(conteo_actual, col, titulo)
            if fig:
                columnas_graficos[i].plotly_chart(fig, use_container_width=True)

    if "NOTIFICADO" in df_filtrado.columns:
        conteo_notificado = df_filtrado["NOTIFICADO"].value_counts().reset_index()
        conteo_notificado.columns = ["NOTIFICADO", "Cantidad"]
        fig = crear_grafico_dinamico(conteo_notificado, "NOTIFICADO", "Expedientes notificados")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # Vista de datos
    st.markdown("<h2 style='font-size: 18px;'>📋 Vista general de expedientes</h2>", unsafe_allow_html=True)
    df_mostrar = df_filtrado.copy()
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%y")
    st.dataframe(df_mostrar, use_container_width=True)
    
    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")

    # Descarga de informes - SOLO se ejecuta al hacer clic
    st.markdown("---")
    st.markdown("<h2 style='font-size: 18px;'>Descarga de Informes</h2>", unsafe_allow_html=True)
    st.markdown("<h3 style='font-size: 16px;'>Generar Informes PDF Pendientes por Usuario</h3>", unsafe_allow_html=True)

    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()

    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF Pendientes", key="generar_pdfs"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for usuario in usuarios_pendientes:
                        df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
                        # ... (resto del código de generación de PDFs)

            zip_buffer.seek(0)
            # ... (resto del código de descarga)

    # SECCIÓN: ENVÍO DE CORREOS MANUALES - SOLO se ejecuta al interactuar
    st.markdown("---")
    st.markdown("<h2 style='font-size: 18px;'>📧 Preparación de Correos para Envío Manual</h2>", unsafe_allow_html=True)

    # El resto del código de envío de correos se mantiene igual...
    # Solo se ejecutará cuando el usuario interactúe con los botones

elif eleccion == "Indicadores clave (KPI)":
    # Usar df_combinado en lugar de df
    df = df_combinado
    
    st.markdown("<h2 style='font-size: 18px;'>Indicadores clave (KPI)</h2>", unsafe_allow_html=True)
    
    # Obtener fecha de referencia para cálculos
    columna_fecha = df.columns[11]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    
    if pd.isna(fecha_max):
        st.error("No se pudo encontrar la fecha máxima en los datos")
        st.stop()
    
    # Crear rango de semanas disponibles (cached)
    @st.cache_data(ttl=CACHE_TTL)
    def obtener_semanas_disponibles(_fecha_max):
        fecha_inicio = pd.to_datetime("2022-11-04")
        semanas_disponibles = pd.date_range(
            start=fecha_inicio,
            end=_fecha_max,
            freq='W-FRI'
        ).tolist()
        return semanas_disponibles
    
    semanas_disponibles = obtener_semanas_disponibles(fecha_max)
    
    if not semanas_disponibles:
        st.error("No hay semanas disponibles para mostrar")
        st.stop()
    
    # Inicialización del estado para KPI
    if 'kpi_semana_index' not in st.session_state:
        st.session_state.kpi_semana_index = len(semanas_disponibles) - 1
    
    # Obtener la semana seleccionada actual
    semana_seleccionada = semanas_disponibles[st.session_state.kpi_semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Selector de semana en el área principal
    st.markdown("---")
    st.markdown("<h2 style='font-size: 18px;'>🗓️ Selector de Semana</h2>", unsafe_allow_html=True)
    
    # Crear etiquetas formateadas para el slider
    opciones_slider = []
    for i, fecha in enumerate(semanas_disponibles):
        num_semana = ((fecha - FECHA_REFERENCIA).days) // 7 + 1
        fecha_str_opcion = fecha.strftime('%d/%m/%Y')
        opciones_slider.append(f"Semana {num_semana} ({fecha_str_opcion})")
    
    # Slider - solo actualiza el índice
    semana_index_slider = st.select_slider(
        "Selecciona la semana:",
        options=list(range(len(semanas_disponibles))),
        value=st.session_state.kpi_semana_index,
        format_func=lambda x: opciones_slider[x],
        key="kpi_semana_slider"
    )
    
    # Actualizar el índice si el slider cambió
    if semana_index_slider != st.session_state.kpi_semana_index:
        st.session_state.kpi_semana_index = semana_index_slider
        st.rerun()
    
    # RECALCULAR después de posibles cambios del slider
    semana_seleccionada = semanas_disponibles[st.session_state.kpi_semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Mostrar información de la semana seleccionada
    st.info(f"**Semana seleccionada:** {fecha_str} (Semana {num_semana_seleccionada})")
    
    # Sidebar con botones de navegación
    with st.sidebar:
        st.header("🗓️ Navegación por Semanas")
        
        st.write(f"**Semana actual:**")
        st.write(f"{fecha_str}")
        st.write(f"(Semana {num_semana_seleccionada})")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("◀️ Anterior", use_container_width=True, key="btn_anterior_kpi"):
                nuevo_indice = st.session_state.kpi_semana_index - 1
                if nuevo_indice >= 0:
                    st.session_state.kpi_semana_index = nuevo_indice
                    st.rerun()
        
        with col2:
            if st.button("Siguiente ▶️", use_container_width=True, key="btn_siguiente_kpi"):
                nuevo_indice = st.session_state.kpi_semana_index + 1
                if nuevo_indice < len(semanas_disponibles):
                    st.session_state.kpi_semana_index = nuevo_indice
                    st.rerun()
        
        st.write(f"**Posición:** {st.session_state.kpi_semana_index + 1} de {len(semanas_disponibles)}")
        
        if st.button("📅 Ir a semana actual", use_container_width=True, key="btn_actual_kpi"):
            st.session_state.kpi_semana_index = len(semanas_disponibles) - 1
            st.rerun()

    # Función para calcular KPIs por semana - CORREGIDA
    def calcular_kpis_para_semana(_df, semana_fin):
        inicio_semana = semana_fin - timedelta(days=6)
        
        # Expedientes abiertos durante la semana
        if 'FECHA APERTURA' in _df.columns:
            nuevos_expedientes = _df[
                (_df['FECHA APERTURA'] >= inicio_semana) & 
                (_df['FECHA APERTURA'] <= semana_fin)
            ].shape[0]
        else:
            nuevos_expedientes = 0
        
        # Expedientes cerrados durante la semana - CORREGIDO
        if 'FECHA CIERRE' in _df.columns:
            expedientes_cerrados = _df[
                (_df['FECHA CIERRE'] >= inicio_semana) & 
                (_df['FECHA CIERRE'] <= semana_fin)
            ].shape[0]
        else:
            expedientes_cerrados = 0

        # Total de expedientes abiertos al final de la semana - CORREGIDO
        if 'FECHA APERTURA' in _df.columns:
            total_abiertos = _df[
                (_df['FECHA APERTURA'] <= semana_fin) & 
                ((_df['FECHA CIERRE'].isna()) | (_df['FECHA CIERRE'] > semana_fin))
            ].shape[0]
        else:
            total_abiertos = 0
        
        return {
            'nuevos_expedientes': nuevos_expedientes,
            'expedientes_cerrados': expedientes_cerrados,
            'total_abiertos': total_abiertos
        }

    # CALCULAR KPIs PARA TODAS LAS SEMANAS con cache
    @st.cache_data(ttl=CACHE_TTL)
    def calcular_kpis_todas_semanas_optimizado(_df, _semanas, _fecha_referencia):
        datos_semanales = []
        
        for semana in _semanas:
            kpis = calcular_kpis_para_semana(_df, semana)
            num_semana = ((semana - _fecha_referencia).days) // 7 + 1
            
            datos_semanales.append({
                'semana_numero': num_semana,
                'semana_fin': semana,
                'semana_str': semana.strftime('%d/%m/%Y'),
                'nuevos_expedientes': kpis['nuevos_expedientes'],
                'expedientes_cerrados': kpis['expedientes_cerrados'],
                'total_abiertos': kpis['total_abiertos']
            })
        
        return pd.DataFrame(datos_semanales)

    # Calcular KPIs para todas las semanas (usando cache)
    df_kpis_semanales = calcular_kpis_todas_semanas_optimizado(df, semanas_disponibles, FECHA_REFERENCIA)

    def mostrar_kpis_principales(_df_kpis, _semana_seleccionada, _num_semana):
        kpis_semana = _df_kpis[_df_kpis['semana_numero'] == _num_semana].iloc[0]
        
        fecha_str = _semana_seleccionada.strftime('%d/%m/%Y')
        st.markdown(f"<h2 style='font-size: 18px;'>📊 KPIs de la Semana: {fecha_str} (Semana {_num_semana})</h2>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="💰 Nuevos Expedientes",
                value=f"{int(kpis_semana['nuevos_expedientes']):,}".replace(",", "."),
                delta=None
            )
        
        with col2:
            st.metric(
                label="✅ Expedientes Cerrados",
                value=f"{int(kpis_semana['expedientes_cerrados']):,}".replace(",", "."),
                delta=None
            )
        
        with col3:
            st.metric(
                label="📂 Total Abiertos",
                value=f"{int(kpis_semana['total_abiertos']):,}".replace(",", "."),
                delta=None
            )
        
        st.markdown("---")
        
        st.markdown("<h3 style='font-size: 16px;'>Detalles de la Semana</h3>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        
        with col1:
            periodo_inicio = (_semana_seleccionada - timedelta(days=6)).strftime('%d/%m/%Y')
            periodo_fin = _semana_seleccionada.strftime('%d/%m/%Y')
            st.write(f"**Período:** {periodo_inicio} a {periodo_fin}")
        
        with col2:
            if kpis_semana['nuevos_expedientes'] > 0 and kpis_semana['expedientes_cerrados'] > 0:
                ratio_cierre = kpis_semana['expedientes_cerrados'] / kpis_semana['nuevos_expedientes']
                st.write(f"**Ratio de cierre:** {ratio_cierre:.2%}")

    # Mostrar dashboard principal
    mostrar_kpis_principales(df_kpis_semanales, semana_seleccionada, num_semana_seleccionada)

    # GRÁFICO DE EVOLUCIÓN TEMPORAL (CACHEADO)
    st.markdown("---")
    st.markdown("<h2 style='font-size: 18px;'>📈 Evolución Temporal de KPIs</h2>", unsafe_allow_html=True)

    # Crear gráfico con cache
    @st.cache_data(ttl=300)  # 5 minutos para el gráfico
    def crear_grafico_evolucion(_df_kpis, _num_semana_seleccionada):
        fig = px.line(
            _df_kpis,
            x='semana_numero',
            y=['nuevos_expedientes', 'expedientes_cerrados', 'total_abiertos'],
            title='Evolución de KPIs a lo largo del tiempo',
            labels={
                'semana_numero': 'Número de Semana',
                'value': 'Cantidad de Expedientes',
                'variable': 'Tipo de KPI'
            },
            color_discrete_map={
                'nuevos_expedientes': '#1f77b4',
                'expedientes_cerrados': '#ff7f0e', 
                'total_abiertos': '#2ca02c'
            }
        )

        # Personalizar el gráfico
        fig.update_layout(
            xaxis_title='Semana',
            yaxis_title='Cantidad de Expedientes',
            legend_title='KPIs',
            hovermode='x unified',
            height=500
        )

        # Actualizar nombres de las leyendas
        fig.for_each_trace(lambda t: t.update(name='Nuevos Expedientes' if t.name == 'nuevos_expedientes' else 
                                             'Expedientes Cerrados' if t.name == 'expedientes_cerrados' else 
                                             'Total Abiertos'))

        # Añadir línea vertical para la semana seleccionada
        fig.add_vline(
            x=_num_semana_seleccionada, 
            line_width=2, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Semana Seleccionada",
            annotation_position="top left"
        )
        
        return fig

    fig = crear_grafico_evolucion(df_kpis_semanales, num_semana_seleccionada)
    st.plotly_chart(fig, use_container_width=True)
    
    # Mostrar tabla con datos históricos
    with st.expander("📋 Ver datos históricos completos"):
        st.dataframe(
            df_kpis_semanales.rename(columns={
                'semana_numero': 'Semana',
                'semana_str': 'Fecha Fin Semana',
                'nuevos_expedientes': 'Nuevos Expedientes',
                'expedientes_cerrados': 'Expedientes Cerrados',
                'total_abiertos': 'Total Abiertos'
            }),
            use_container_width=True
        )