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
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
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
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]

menu = ["Principal", "Indicadores clave (KPI)", "Envío de correos"]
eleccion = st.sidebar.selectbox("Menú", menu)
if eleccion == "Principal":
    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    # Copiamos el DataFrame original para no modificar el cargado
    df_enriquecido = df.copy()

    # Ejemplo de creación de 12 columnas
    df_enriquecido["EQUIPO"] = df_enriquecido["EQUIPO"]
    df_enriquecido["DIAS_HASTA_MAX"] = (fecha_max - df_enriquecido[columna_fecha]).dt.days
    df_enriquecido["SEMANA_EXPEDIENTE"] = ((df_enriquecido[columna_fecha] - FECHA_REFERENCIA).dt.days // 7 + 1)

    # Mostrar las tres primeras columnas nuevas
    st.subheader("📊 Vista previa de las nuevas columnas")
    columnas_preview = ["EQUIPO"]
    st.dataframe(df[columnas_preview].head(10), use_container_width=True)
    
    #Sidebar para filtros
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
    #for col in df_mostrar.select_dtypes(include='number').columns:
    #    df_mostrar.style.format("{:,}", thousands=".", na_rep="")
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
                help="Descarga todos los informes PDF listos.",
                key='pdf_download_button'
            )
    
elif eleccion == "Envío de correos":
    st.subheader("Envío de correos")
    
    
elif eleccion == "Indicadores clave (KPI)":
    st.subheader("Indicadores clave (KPI)")
    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    
    if not fecha_max:
        st.error("No se pudo encontrar la fecha máxima")
    
    # Crear rango de semanas disponibles
    fecha_inicio = pd.to_datetime("2022-11-04")
    semanas_disponibles = pd.date_range(
        start=fecha_inicio,
        end=fecha_max,
        freq='W-FRI'
    )
    
    # Sidebar para selección
    with st.sidebar:
        st.header("🗓️ Selector de Semana")
        
        # Selector de fecha con slider
        semana_seleccionada = st.select_slider(
            "Selecciona la semana:",
            options=semanas_disponibles,
            value=semanas_disponibles[-1],  # Última semana por defecto
            format_func=lambda x: x.strftime("%d/%m/%Y")
        )
        
        st.markdown("---")
        st.info(f"**Semana seleccionada:** {semana_seleccionada.strftime('%d/%m/%Y')}")
        
        # Navegación rápida
        st.subheader("Navegación Rápida")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("⬅️ Semana Anterior"):
                idx = list(semanas_disponibles).index(semana_seleccionada)
                if idx > 0:
                    semana_seleccionada = semanas_disponibles[idx - 1]
        with col2:
            if st.button("Semana Siguiente ➡️"):
                idx = list(semanas_disponibles).index(semana_seleccionada)
                if idx < len(semanas_disponibles) - 1:
                    semana_seleccionada = semanas_disponibles[idx + 1]
    
    # Calcular KPIs para la semana seleccionada
    kpis_semana = calcular_kpis_semana(df_enriquecido, semana_seleccionada)

    # Mostrar dashboard principal
    mostrar_kpis_principales(kpis_semana, semana_seleccionada)
    mostrar_detalles_semana(df, semana_seleccionada)

def calcular_kpis_semana(df, semana_seleccionada):
    """
    Calcula KPIs específicos para la semana seleccionada
    """
    # Definir rango de la semana (de viernes a jueves)
    inicio_semana = semana_seleccionada - timedelta(days=4)  # Viernes
    fin_semana = semana_seleccionada  # Jueves
    
    # Filtrar datos de la semana
    mascara_semana = (df['fecha'] >= inicio_semana) & (df['fecha'] <= fin_semana)
    datos_semana = df[mascara_semana]
    
    # Calcular KPIs (AJUSTA SEGÚN TUS COLUMNAS)
    kpis = {
        'ventas_totales': datos_semana['ventas'].sum() if 'ventas' in df.columns else 0,
        'transacciones': len(datos_semana),
        'clientes_unicos': datos_semana['cliente_id'].nunique() if 'cliente_id' in df.columns else 0,
        'ticket_promedio': datos_semana['ventas'].mean() if 'ventas' in df.columns else 0,
        'productos_vendidos': datos_semana['producto_id'].nunique() if 'producto_id' in df.columns else 0,
        'dias_activos': datos_semana['fecha'].nunique(),
        'venta_maxima': datos_semana['ventas'].max() if 'ventas' in df.columns else 0,
        'venta_minima': datos_semana['ventas'].min() if 'ventas' in df.columns else 0,
        'empleados_activos': datos_semana['empleado_id'].nunique() if 'empleado_id' in df.columns else 0,
        'eficiencia_ventas': datos_semana['ventas'].sum() / len(datos_semana) if len(datos_semana) > 0 else 0,
        'tasa_conversion': "N/A",  # Ajusta según tu métrica
        'satisfaccion_promedio': datos_semana['satisfaccion'].mean() if 'satisfaccion' in df.columns else 0
    }
    
    return kpis

def mostrar_kpis_principales(kpis, semana_seleccionada):
    """
    Muestra los KPIs principales en tarjetas estilo dashboard
    """
    st.header(f"📊 KPIs de la Semana: {semana_seleccionada.strftime('%d/%m/%Y')}")
    
    # KPIs principales (primera fila)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="💰 Ventas Totales",
            value=f"${kpis['ventas_totales']:,.0f}",
            delta=None
        )
    
    with col2:
        st.metric(
            label="🛒 Transacciones",
            value=f"{kpis['transacciones']:,}",
            delta=None
        )
    
    with col3:
        st.metric(
            label="👥 Clientes Únicos",
            value=f"{kpis['clientes_unicos']:,}",
            delta=None
        )
    
    with col4:
        st.metric(
            label="🎫 Ticket Promedio",
            value=f"${kpis['ticket_promedio']:,.2f}",
            delta=None
        )
    
    # Segunda fila de KPIs
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        st.metric(
            label="📦 Productos Vendidos",
            value=f"{kpis['productos_vendidos']:,}",
            delta=None
        )
    
    with col6:
        st.metric(
            label="📅 Días Activos",
            value=f"{kpis['dias_activos']}",
            delta=None
        )
    
    with col7:
        st.metric(
            label="📈 Venta Máxima",
            value=f"${kpis['venta_maxima']:,.0f}",
            delta=None
        )
    
    with col8:
        st.metric(
            label="📉 Venta Mínima",
            value=f"${kpis['venta_minima']:,.0f}",
            delta=None
        )
    
    st.markdown("---")

def mostrar_detalles_semana(df, semana_seleccionada):
    """
    Muestra detalles adicionales y visualizaciones para la semana seleccionada
    """
    # Filtrar datos de la semana
    inicio_semana = semana_seleccionada - timedelta(days=6)
    fin_semana = semana_seleccionada
    datos_semana = df[(df['fecha'] >= inicio_semana) & (df['fecha'] <= fin_semana)]
    
    if datos_semana.empty:
        st.warning("No hay datos para la semana seleccionada")
        return
    
    # Pestañas para diferentes vistas
    tab1, tab2, tab3 = st.tabs(["📈 Tendencia Diaria", "📊 Análisis Detallado", "📋 Datos Crudos"])
    
    with tab1:
        st.subheader("Tendencia Diaria de Ventas")
        
        # Agrupar por día
        ventas_diarias = datos_semana.groupby('fecha')['ventas'].sum().reset_index()
        
        if not ventas_diarias.empty:
            fig = px.line(
                ventas_diarias, 
                x='fecha', 
                y='ventas',
                title=f"Ventas Diarias - Semana del {inicio_semana.strftime('%d/%m/%Y')} al {fin_semana.strftime('%d/%m/%Y')}",
                markers=True
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos de ventas para mostrar")
    
    with tab2:
        st.subheader("Análisis Detallado")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Top productos/clientes (ajusta según tus columnas)
            if 'producto_id' in datos_semana.columns:
                top_productos = datos_semana['producto_id'].value_counts().head(10)
                st.write("**Top 10 Productos:**")
                st.dataframe(top_productos)
        
        with col2:
            if 'cliente_id' in datos_semana.columns:
                top_clientes = datos_semana.groupby('cliente_id')['ventas'].sum().nlargest(5)
                st.write("**Top 5 Clientes por Ventas:**")
                st.dataframe(top_clientes)
    
    with tab3:
        st.subheader("Datos de la Semana")
        st.dataframe(
            datos_semana,
            use_container_width=True,
            height=400
        )
        
        # Estadísticas adicionales
        st.subheader("Estadísticas Adicionales")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"**Primera transacción:** {datos_semana['fecha'].min().strftime('%d/%m/%Y %H:%M')}")
            st.write(f"**Última transacción:** {datos_semana['fecha'].max().strftime('%d/%m/%Y %H:%M')}")
        
        with col2:
            st.write(f"**Total de registros:** {len(datos_semana):,}")
            st.write(f"**Días con actividad:** {datos_semana['fecha'].dt.date.nunique()}")

def cargar_datos():
    """
    Función para cargar tus datos - AJUSTA ESTO
    """
    # Ejemplo - reemplaza con tu carga real
    try:
        # return pd.read_excel('tu_archivo.xlsx')
        return df_filtrado  # Tu DataFrame existente
    except:
        return None

if __name__ == "__main__":
    main()


