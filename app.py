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

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"
ESTADOS_PENDIENTES = ["Abierto"]

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Seguimiento Equipo Regional RECTAUTO")

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
    """Genera un PDF desde un DataFrame, con encabezados ajustables y repetidos."""
    pdf = PDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", "B", 8)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    # --- CONFIGURACIÓN ---
    col_widths = [43, 14, 14, 8, 24, 14, 14, 24, 14, 40, 24, 14, 26]
    df_mostrar_pdf = df.iloc[:, :len(col_widths)]

    # --- CONFIGURACIÓN DE ALTURA FIJA ---
    ALTURA_ENCABEZADO = 11  # Altura fija en mm para todos los encabezados

    # --- FUNCIÓN PARA IMPRIMIR ENCABEZADOS ---
    def imprimir_encabezados():
        pdf.set_font("Arial", "B", 5)
        pdf.set_fill_color(200, 220, 255)
        y_inicio = pdf.get_y()
        
        # Dibujar todos los encabezados con altura fija
        for i, header in enumerate(df_mostrar_pdf.columns):
            x = pdf.get_x()
            y = pdf.get_y()
            
            # Dibujar el rectángulo de fondo con altura fija
            pdf.cell(col_widths[i], ALTURA_ENCABEZADO, "", 1, 0, 'C', True)
            
            # Volver a la posición para escribir el texto
            pdf.set_xy(x, y)
            
            # Calcular posición vertical para centrar el texto
            texto = str(header)
            ancho_texto = pdf.get_string_width(texto)
            
            # Si el texto cabe en una línea, centrarlo verticalmente
            if ancho_texto <= col_widths[i] - 2:  # Margen de 2mm
                # Centrar verticalmente para una línea
                altura_texto = 3  # Altura aproximada del texto
                y_pos = y + (ALTURA_ENCABEZADO - altura_texto) / 2
                pdf.set_xy(x, y_pos)
                pdf.cell(col_widths[i], altura_texto, texto, 0, 0, 'C')
            else:
                # Para texto multilínea, usar multi_cell
                pdf.set_xy(x, y + 1)  # Pequeño margen superior
                pdf.multi_cell(col_widths[i], 2.5, texto, 0, 'C')
            
            # Mover a la siguiente columna
            pdf.set_xy(x + col_widths[i], y)
        
        # Mover a la siguiente línea para los datos
        pdf.set_xy(pdf.l_margin, y_inicio + ALTURA_ENCABEZADO)

    # --- PRIMER ENCABEZADO ---
    imprimir_encabezados()

    # --- IMPRIMIR DATOS ---
    pdf.set_font("Arial", "", 7)
    for _, row in df_mostrar_pdf.iterrows():
        # Si la siguiente fila no cabe, añadir nueva página y repetir encabezados
        if pdf.get_y() + 6 > 190:
            pdf.add_page()
            imprimir_encabezados()

        for i, col_data in enumerate(row):
            text = str(col_data).replace("\n", " ")
            pdf.cell(col_widths[i], 6, text, 1, 0, 'L')
        pdf.ln()

    # --- EXPORTAR COMO BYTES ---
    pdf_output = pdf.output(dest='B')
    return pdf_output

# CSS para ambos fondos
st.markdown("""
<style>
    /* Barra lateral - Verde oscuro */
    [data-testid="stSidebar"] {
    background-color: #007933 !important;
    }
    
    /* Área principal - Verde claro */
    .main .block-container {
        background-color: #C4DDCA !important;
        padding: 2rem;
        border-radius: 10px;
    }
    
    /* Fondo general de la página */
    .stApp {
        background-color: #92C88F !important;
    }
    
    /* Texto en barra lateral */
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    /* Mejorar contraste en área principal */
    .main .stMarkdown, .main h1, .main h2, .main h3 {
        color: #333333 !important;
    }
</style>
""", unsafe_allow_html=True)
    
# Logo que funciona como enlace
st.sidebar.image("Logo Atrian.png", width=260)

    
archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, thousands='.', decimal=',', engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd")
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 6, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]
    st.session_state["df"] = df
elif "df" in st.session_state:
    df = st.session_state["df"]
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
    st.stop()

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, thousands='.', decimal=',', engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd")
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 6, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]

menu = ["Principal", "Indicadores clave (KPI)", "Envío de correos"]
eleccion = st.sidebar.selectbox("Menú", menu)

# Inicializar estado de navegación de semanas
if 'semana_index' not in st.session_state:
    st.session_state.semana_index = 0

if eleccion == "Principal":
    # ... (código de la sección Principal se mantiene igual)
    columna_fecha = df.columns[11]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    # Copiamos el DataFrame original para no modificar el cargado
    df_enriquecido = df.copy()

    # Sidebar para filtros
    st.sidebar.header("Filtros")

    # Inicializar session_state para los filtros si no existen
    if 'filtro_estado' not in st.session_state:
        st.session_state.filtro_estado = ['Abierto'] if 'Abierto' in df['ESTADO'].values else []

    if 'filtro_equipo' not in st.session_state:
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())

    if 'filtro_usuario' not in st.session_state:
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())

    # Botón para mostrar todos los elementos
    if st.sidebar.button("Mostrar todos / Resetear filtros"):
        st.session_state.filtro_estado = sorted(df['ESTADO'].dropna().unique())
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())
        st.rerun()

    # Obtener opciones ordenadas
    opciones_estado = sorted(df['ESTADO'].dropna().unique())
    opciones_equipo = sorted(df['EQUIPO'].dropna().unique())
    opciones_usuario = sorted(df['USUARIO'].dropna().unique())

    # Filtro de ESTADO
    estado_sel = st.sidebar.multiselect(
        "Selecciona Estado:",
        options=opciones_estado,
        default=st.session_state.filtro_estado,
        key='filtro_estado'
    )

    # Filtro de EQUIPO
    equipo_sel = st.sidebar.multiselect(
        "Selecciona Equipo:",
        options=opciones_equipo,
        default=st.session_state.filtro_equipo,
        key='filtro_equipo'
    )

    # Filtro de USUARIO
    usuario_sel = st.sidebar.multiselect(
        "Selecciona Usuario:",
        options=opciones_usuario,
        default=st.session_state.filtro_usuario,
        key='filtro_usuario'
    )

    # Aplicar filtros al DataFrame
    df_filtrado = df.copy()

    if estado_sel:
        df_filtrado = df_filtrado[df_filtrado['ESTADO'].isin(estado_sel)]

    if equipo_sel:
        df_filtrado = df_filtrado[df_filtrado['EQUIPO'].isin(equipo_sel)]

    if usuario_sel:
        df_filtrado = df_filtrado[df_filtrado['USUARIO'].isin(usuario_sel)]

    # Mostrar qué filtros están activos
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros activos")
    if estado_sel:
        st.sidebar.write(f"Estados: {', '.join(estado_sel)}")
    if equipo_sel:
        st.sidebar.write(f"Equipos: {len(equipo_sel)} seleccionados")
    if usuario_sel:
        st.sidebar.write(f"Usuarios: {len(usuario_sel)} seleccionados")
    

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
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%y")
    st.dataframe(df_mostrar, use_container_width=True)
    
    # Mostrar contador de registros
    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")


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
                indices_a_excluir = {1, 6, 11}
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
                help="Descarga todos los informes PDF listos.",
                key='pdf_download_button'
            )
    
elif eleccion == "Envío de correos":
    st.subheader("Envío de correos")
    
elif eleccion == "Indicadores clave (KPI)":
    st.subheader("Indicadores clave (KPI)")
    
    # Obtener fecha de referencia para cálculos
    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    
    if pd.isna(fecha_max):
        st.error("No se pudo encontrar la fecha máxima en los datos")
        st.stop()
    
    # Crear rango de semanas disponibles
    fecha_inicio = pd.to_datetime("2022-11-04")
    semanas_disponibles = pd.date_range(
        start=fecha_inicio,
        end=fecha_max,
        freq='W-FRI'
    ).tolist()
    
    # VERIFICACIÓN INICIAL CRÍTICA
    if not semanas_disponibles:
        st.error("No hay semanas disponibles para mostrar")
        st.stop()
    
    # INICIALIZACIÓN COMPLETAMENTE NUEVA - ENFOQUE DEFENSIVO
    # Reiniciar completamente el estado si hay problemas
    try:
        current_index = st.session_state.semana_index
        # Verificar que sea un entero válido
        if not isinstance(current_index, int):
            st.session_state.semana_index = len(semanas_disponibles) - 1
            current_index = st.session_state.semana_index
    except (KeyError, AttributeError):
        # Si no existe o hay error, inicializar
        st.session_state.semana_index = len(semanas_disponibles) - 1
        current_index = st.session_state.semana_index
    
    # VERIFICACIÓN EXPLÍCITA DE TIPOS Y LÍMITES
    if not isinstance(current_index, int):
        st.session_state.semana_index = len(semanas_disponibles) - 1
        current_index = st.session_state.semana_index
    
    # Ajustar índice si está fuera de los límites
    if current_index < 0:
        st.session_state.semana_index = 0
    elif current_index >= len(semanas_disponibles):
        st.session_state.semana_index = len(semanas_disponibles) - 1
    
    # Obtener la semana seleccionada actual
    semana_seleccionada = semanas_disponibles[st.session_state.semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Selector de semana en el área principal
    st.markdown("---")
    st.header("🗓️ Selector de Semana")
    
    # Crear etiquetas formateadas para el slider
    opciones_slider = []
    for i, fecha in enumerate(semanas_disponibles):
        num_semana = ((fecha - FECHA_REFERENCIA).days) // 7 + 1
        fecha_str_opcion = fecha.strftime('%d/%m/%Y')
        opciones_slider.append(f" {fecha_str_opcion} (Semana {num_semana})")
    
    # Slider corregido con formato adecuado
    semana_index_slider = st.select_slider(
        "Selecciona la semana:",
        options=list(range(len(semanas_disponibles))),
        value=st.session_state.semana_index,
        format_func=lambda x: opciones_slider[x]
    )
    
    # Actualizar el índice si el slider cambió
    if semana_index_slider != st.session_state.semana_index:
        st.session_state.semana_index = semana_index_slider
        st.rerun()
    
    # RECALCULAR después de posibles cambios del slider
    semana_seleccionada = semanas_disponibles[st.session_state.semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Mostrar información de la semana seleccionada
    st.info(f"**Semana seleccionada:** {fecha_str} (Semana {num_semana_seleccionada})")
    
    # Sidebar con botones de navegación
    with st.sidebar:
        st.header("🗓️ Navegación por Semanas")
        
        # Mostrar semana actual en sidebar
        st.write(f"**Semana actual:**")
        st.write(f"{fecha_str}")
        st.write(f"(Semana {num_semana_seleccionada})")
        
        st.markdown("---")
        
        # Botones de navegación
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("◀️ Anterior", use_container_width=True, key="btn_anterior"):
                nuevo_indice = st.session_state.semana_index - 1
                if nuevo_indice >= 0:
                    st.session_state.semana_index = nuevo_indice
                    st.rerun()
        
        with col2:
            if st.button("Siguiente ▶️", use_container_width=True, key="btn_siguiente"):
                nuevo_indice = st.session_state.semana_index + 1
                if nuevo_indice < len(semanas_disponibles):
                    st.session_state.semana_index = nuevo_indice
                    st.rerun()
        
        # Indicador de posición
        st.write(f"**Posición:** {st.session_state.semana_index + 1} de {len(semanas_disponibles)}")
        
        # Botón para ir a la semana más reciente
        if st.button("📅 Ir a semana actual", use_container_width=True, key="btn_actual"):
            st.session_state.semana_index = len(semanas_disponibles) - 1
            st.rerun()

    # El resto del código permanece igual...
    def calcular_kpis_para_semana(df, semana_fin):
        """
        Calcula KPIs específicos para una semana dada
        """
        # Definir rango de la semana (de lunes a domingo)
        inicio_semana = semana_fin - timedelta(days=6)
        
        # NUEVOS EXPEDIENTES (aperturas en la semana)
        if 'FECHA APERTURA' in df.columns:
            nuevos_expedientes = df[
                (df['FECHA APERTURA'] >= inicio_semana) & 
                (df['FECHA APERTURA'] <= semana_fin)
            ].shape[0]
        else:
            nuevos_expedientes = 0
        
        # EXPEDIENTES CERRADOS EN LA SEMANA
        if 'ESTADO' in df.columns and 'FECHA ÚLTIMO TRAM.' in df.columns:
            expedientes_cerrados = df[
                (df['ESTADO'] == 'Cerrado') & 
                (df['FECHA ÚLTIMO TRAM.'] >= inicio_semana) & 
                (df['FECHA ÚLTIMO TRAM.'] <= semana_fin)
            ].shape[0]
        else:
            expedientes_cerrados = 0

        # TOTAL EXPEDIENTES ABIERTOS AL FINAL DE LA SEMANA
        if 'FECHA CIERRE' in df.columns and 'FECHA APERTURA' in df.columns:
            total_abiertos = df[
                (df['FECHA APERTURA'] <= semana_fin) & 
                ((df['FECHA CIERRE'] > semana_fin) | (df['FECHA CIERRE'].isna()))
            ].shape[0]
        else:
            total_abiertos = 0
        
        return {
            'nuevos_expedientes': nuevos_expedientes,
            'expedientes_cerrados': expedientes_cerrados,
            'total_abiertos': total_abiertos
        }

    # CALCULAR KPIs PARA TODAS LAS SEMANAS (una sola vez)
    @st.cache_data
    def calcular_kpis_todas_semanas(_df, semanas):
        """
        Calcula KPIs para todas las semanas y retorna un DataFrame
        """
        datos_semanales = []
        
        for semana in semanas:
            kpis = calcular_kpis_para_semana(_df, semana)
            num_semana = ((semana - FECHA_REFERENCIA).days) // 7 + 1
            
            datos_semanales.append({
                'semana_numero': num_semana,
                'semana_fin': semana,
                'semana_str': semana.strftime('%d/%m/%Y'),
                'nuevos_expedientes': kpis['nuevos_expedientes'],
                'expedientes_cerrados': kpis['expedientes_cerrados'],
                'total_abiertos': kpis['total_abiertos']
            })
        
        return pd.DataFrame(datos_semanales)

    # Calcular KPIs para todas las semanas
    df_kpis_semanales = calcular_kpis_todas_semanas(df, semanas_disponibles)

    def mostrar_kpis_principales(_df_kpis, _semana_seleccionada, _num_semana):
        """
        Muestra los KPIs principales en tarjetas estilo dashboard
        """
        # Obtener KPIs para la semana seleccionada
        kpis_semana = _df_kpis[_df_kpis['semana_numero'] == _num_semana].iloc[0]
        
        fecha_str = _semana_seleccionada.strftime('%d/%m/%Y')
        st.header(f"📊 KPIs de la Semana: {fecha_str} (Semana {_num_semana})")
        
        # KPIs principales
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="💰 Nuevos Expedientes",
                value=int(kpis_semana['nuevos_expedientes']),
                delta=None
            )
        
        with col2:
            st.metric(
                label="🛒 Expedientes cerrados",
                value=int(kpis_semana['expedientes_cerrados']),
                delta=None
            )
        
        with col3:
            st.metric(
                label="👥 Total expedientes abiertos",
                value=int(kpis_semana['total_abiertos']),
                delta=None
            )
        
        st.markdown("---")
        
        # Mostrar detalles adicionales
        st.subheader("Detalles de la Semana")
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

    # GRÁFICO DE EVOLUCIÓN TEMPORAL
    st.markdown("---")
    st.subheader("📈 Evolución Temporal de KPIs")
    
    # Crear gráfico de líneas
    fig = px.line(
        df_kpis_semanales,
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
        x=num_semana_seleccionada, 
        line_width=2, 
        line_dash="dash", 
        line_color="red",
        annotation_text="Semana Seleccionada"
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Mostrar tabla con datos históricos (opcional)
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
