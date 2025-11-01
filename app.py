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
st.title("📊 Seguimiento Equipo Regional RECTAUTO")

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
    pdf.set_font("Arial", "B", 8)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    col_widths = [35, 12, 12, 8, 20, 14, 14, 20, 12, 35, 20, 12, 20]
    df_mostrar_pdf = df.iloc[:, :len(col_widths)]
    ALTURA_ENCABEZADO = 11

    def imprimir_encabezados():
        pdf.set_font("Arial", "", 6)
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

    pdf.set_font("Arial", "", 6)
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

# CSS (se mantiene igual)
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
</style>
""", unsafe_allow_html=True)

# Logo
st.sidebar.image("Logo Atrian.png", width=260)

# Botón para limpiar cache
with st.sidebar:
    st.markdown("---")
    if st.button("🔄 Limpiar cache", help="Limpiar toda la cache y recargar"):
        st.cache_data.clear()
        # Mantener solo los datos esenciales
        keys_to_keep = ['df', 'archivo_hash', 'filtro_estado', 'filtro_equipo', 'filtro_usuario']
        for key in list(st.session_state.keys()):
            if key not in keys_to_keep:
                del st.session_state[key]
        st.success("Cache limpiada correctamente")
        st.rerun()

# Carga optimizada de archivos
archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    archivo_hash = obtener_hash_archivo(archivo)
    
    # Verificar si el archivo es nuevo o cambió
    if ('archivo_hash' not in st.session_state or 
        st.session_state.archivo_hash != archivo_hash or 
        "df" not in st.session_state):
        
        with st.spinner("🔄 Procesando datos (esto puede tomar unos momentos)..."):
            df = cargar_y_procesar_datos(archivo)
            st.session_state["df"] = df
            st.session_state["archivo_hash"] = archivo_hash
            st.success("✅ Datos procesados y cacheados por 2 horas")
    else:
        # Usar datos cacheados
        df = st.session_state["df"]
        st.sidebar.success("✅ Usando datos cacheados")
        
elif "df" in st.session_state:
    df = st.session_state["df"]
    st.sidebar.info("📊 Datos cargados desde cache")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
    st.stop()

# Menú principal (sin "Envío de correos")
menu = ["Principal", "Indicadores clave (KPI)"]
eleccion = st.sidebar.selectbox("Menú", menu)

# Función para gráficos dinámicos (SIN CACHE)
def crear_grafico_dinamico(_conteo, columna, titulo):
    """Crea gráficos dinámicos que responden a los filtros"""
    if _conteo.empty:
        return None
    
    fig = px.bar(_conteo, y=columna, x="Cantidad", title=titulo, text="Cantidad", 
                 color=columna, height=400)
    fig.update_traces(texttemplate='%{text:,}', textposition="auto")
    return fig

if eleccion == "Principal":
    columna_fecha = df.columns[11]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    # Sidebar para filtros
    st.sidebar.header("Filtros")

    if 'filtro_estado' not in st.session_state:
        st.session_state.filtro_estado = ['Abierto'] if 'Abierto' in df['ESTADO'].values else []

    if 'filtro_equipo' not in st.session_state:
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())

    if 'filtro_usuario' not in st.session_state:
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())

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
        key='filtro_estado'
    )

    equipo_sel = st.sidebar.multiselect(
        "Selecciona Equipo:",
        options=opciones_equipo,
        default=st.session_state.filtro_equipo,
        key='filtro_equipo'
    )

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

    # Mostrar filtros activos
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros activos")
    if estado_sel:
        st.sidebar.write(f"Estados: {', '.join(estado_sel)}")
    if equipo_sel:
        st.sidebar.write(f"Equipos: {len(equipo_sel)} seleccionados")
    if usuario_sel:
        st.sidebar.write(f"Usuarios: {len(usuario_sel)} seleccionados")

    # Gráficos Generales - CORREGIDOS: datos siempre frescos según filtros
    st.subheader("📈 Gráficos Generales")
    columnas_graficos = st.columns(3)
    graficos = [("EQUIPO", "Expedientes por equipo"), 
                ("USUARIO", "Expedientes por usuario"), 
                ("ESTADO", "Distribución por estado")]

    for i, (col, titulo) in enumerate(graficos):
        if col in df_filtrado.columns:
            # Calcular el conteo actual (siempre fresco según los filtros)
            conteo_actual = df_filtrado[col].value_counts().reset_index()
            conteo_actual.columns = [col, "Cantidad"]
            
            # Crear gráfico con datos actualizados (SIN CACHE)
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
    st.subheader("📋 Vista general de expedientes")
    df_mostrar = df_filtrado.copy()
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%Y")
    st.dataframe(df_mostrar, use_container_width=True)
    
    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")

    # Descarga de informes
    st.markdown("---")
    st.header("Descarga de Informes")
    st.subheader("Generar Informes PDF Pendientes por Usuario")

    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()

    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF Pendientes", key="generar_pdfs"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()
                indices_a_incluir = list(range(df_pendientes.shape[1]))
                indices_a_excluir = {1, 4, 10}
                indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
                NOMBRES_COLUMNAS_PDF = df_pendientes.columns[indices_finales].tolist()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for usuario in usuarios_pendientes:
                    df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
                    indice_columna_a_redondear = {0}
                    nombre_columna_a_redondear = df_user.columns[indice_columna_a_redondear]

                    if nombre_columna_a_redondear in df_user.columns:
                        df_user[nombre_columna_a_redondear] = pd.to_numeric(df_user[nombre_columna_a_redondear], errors='coerce').fillna(0).round(0).astype(int)

                    df_pdf = df_user[NOMBRES_COLUMNAS_PDF].copy()
                    for col in df_pdf.select_dtypes(include='datetime').columns:
                        df_pdf[col] = df_pdf[col].dt.strftime("%d/%m/%Y")

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

    # NUEVA SECCIÓN: ENVÍO DE CORREOS INTEGRADA
    st.markdown("---")
    st.header("📧 Envío de Correos")
    
    # Verificar que estamos usando la última semana
    st.info(f"**📅 Semana activa para envío:** {num_semana} (Última semana disponible - {fecha_max_str})")
    
    # Cargar configuración de usuarios desde Excel
    def cargar_configuracion_usuarios():
        """Carga la configuración de usuarios desde archivo Excel"""
        try:
            archivo_usuarios = st.file_uploader("📁 Sube el archivo USUARIOS.xlsx", type=["xlsx", "xls"], key="usuarios_upload")
            
            if archivo_usuarios:
                usuarios_df = pd.read_excel(archivo_usuarios, sheet_name="Sheet1")
                
                # Normalizar nombres de columnas
                usuarios_df.columns = [col.strip().upper() for col in usuarios_df.columns]
                
                # Verificar columnas requeridas
                columnas_requeridas = ['USUARIOS', 'ENVIAR', 'EMAIL', 'ASUNTO', 'MENSAJE']
                columnas_faltantes = [col for col in columnas_requeridas if col not in usuarios_df.columns]
                
                if columnas_faltantes:
                    st.error(f"❌ Faltan columnas en el archivo: {', '.join(columnas_faltantes)}")
                    return None
                
                st.success(f"✅ Archivo USUARIOS.xlsx cargado correctamente: {len(usuarios_df)} usuarios")
                return usuarios_df
            else:
                st.info("📝 Por favor, sube el archivo USUARIOS.xlsx para habilitar el envío de correos")
                return None
                
        except Exception as e:
            st.error(f"❌ Error al cargar USUARIOS.xlsx: {e}")
            return None
    
    # Cargar configuración
    usuarios_config = cargar_configuracion_usuarios()
    
    if usuarios_config is not None:
        # Filtrar usuarios activos
        usuarios_activos = usuarios_config[
            (usuarios_config['ENVIAR'].str.upper() == 'SÍ')
        ]
        
        if usuarios_activos.empty:
            st.warning("⚠️ No hay usuarios activos para envío (ENVIAR = 'SÍ')")
        else:
            # Función para generar el cuerpo del mensaje dinámicamente
            def generar_cuerpo_mensaje(mensaje_base):
                """Genera el cuerpo del mensaje con saludo según la hora"""
                from datetime import datetime
                
                hora_actual = datetime.now().hour
                saludo = "Buenos días" if hora_actual < 14 else "Buenas tardes"
                
                cuerpo_mensaje = f"{saludo},\n\n{mensaje_base}"
                return cuerpo_mensaje
            
            # Función para procesar el asunto con variables
            def procesar_asunto(asunto_template, num_semana, fecha_max_str):
                """Reemplaza variables en el asunto del correo"""
                asunto_procesado = asunto_template.replace("&num_semana&", str(num_semana))
                asunto_procesado = asunto_procesado.replace("&fecha_max&", fecha_max_str)
                return asunto_procesado
            
            # Función para enviar correos con Outlook (funciona con Outlook cerrado)
            def enviar_correo_outlook(destinatario, asunto, cuerpo_mensaje, archivo_pdf, nombre_archivo, cc=None, bcc=None):
                """
                Envía correo usando Outlook local (funciona con Outlook cerrado)
                """
                try:
                    import win32com.client
                    import os
                    import tempfile
                    
                    # Crear cliente Outlook
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    mail = outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # Configurar correo
                    mail.To = destinatario
                    mail.Subject = asunto
                    mail.Body = cuerpo_mensaje
                    
                    # Agregar CC si existe
                    if cc and pd.notna(cc) and str(cc).strip():
                        mail.CC = str(cc)
                    
                    # Agregar BCC si existe
                    if bcc and pd.notna(bcc) and str(bcc).strip():
                        mail.BCC = str(bcc)
                    
                    # Guardar PDF temporalmente para adjuntar
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                        temp_file.write(archivo_pdf)
                        temp_path = temp_file.name
                    
                    # Adjuntar PDF
                    mail.Attachments.Add(temp_path)
                    
                    # Enviar correo (usar Send() en lugar de Display())
                    mail.Send()
                    
                    # Limpiar archivo temporal
                    try:
                        os.unlink(temp_path)
                    except:
                        pass  # Ignorar errores al eliminar temporal
                    
                    return True
                    
                except ImportError:
                    st.error("❌ Error: win32com.client no está disponible. Instala pywin32: pip install pywin32")
                    return False
                except Exception as e:
                    st.error(f"❌ Error al enviar correo a {destinatario}: {e}")
                    return False
            
            # Preparar informes PDF para cada usuario
            def generar_informe_usuario(usuario):
                """Genera el PDF para un usuario específico"""
                df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
                
                if df_user.empty:
                    return None
                
                # Procesar datos para PDF
                indices_a_incluir = list(range(df_user.shape[1]))
                indices_a_excluir = {1, 4, 10}
                indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
                NOMBRES_COLUMNAS_PDF = df_user.columns[indices_finales].tolist()
                
                # Redondear columna numérica si existe
                indice_columna_a_redondear = 4
                if indice_columna_a_redondear < len(df_user.columns):
                    nombre_columna_a_redondear = df_user.columns[indice_columna_a_redondear]
                    if nombre_columna_a_redondear in df_user.columns:
                        df_user[nombre_columna_a_redondear] = pd.to_numeric(
                            df_user[nombre_columna_a_redondear], errors='coerce'
                        ).fillna(0).round(0).astype(int)
                
                df_pdf = df_user[NOMBRES_COLUMNAS_PDF].copy()
                for col in df_pdf.select_dtypes(include='datetime').columns:
                    df_pdf[col] = df_pdf[col].dt.strftime("%d/%m/%Y")
                
                num_expedientes = len(df_pdf)
                titulo_pdf = f"{usuario} - Semana {num_semana} a {fecha_max_str} - Expedientes Pendientes ({num_expedientes})"
                
                return dataframe_to_pdf_bytes(df_pdf, titulo_pdf)
            
            # Verificar usuarios con expedientes pendientes
            usuarios_con_pendientes = df_pendientes['USUARIO'].dropna().unique()
            usuarios_para_envio = []
            
            for _, usuario_row in usuarios_activos.iterrows():
                usuario = usuario_row['USUARIOS']
                if usuario in usuarios_con_pendientes:
                    num_expedientes = len(df_pendientes[df_pendientes['USUARIO'] == usuario])
                    
                    # Procesar asunto con variables
                    asunto_template = usuario_row['ASUNTO'] if pd.notna(usuario_row['ASUNTO']) else f"Situación RECTAUTO asignados en la semana {num_semana} a {fecha_max_str}"
                    asunto_procesado = procesar_asunto(asunto_template, num_semana, fecha_max_str)
                    
                    # Generar cuerpo del mensaje
                    mensaje_base = usuario_row['MENSAJE'] if pd.notna(usuario_row['MENSAJE']) else "Se adjunta informe de expedientes pendientes."
                    cuerpo_mensaje = generar_cuerpo_mensaje(mensaje_base)
                    
                    usuarios_para_envio.append({
                        'usuario': usuario,
                        'resumen': usuario_row.get('RESUMEN', ''),
                        'email': usuario_row['EMAIL'],
                        'cc': usuario_row.get('CC', ''),
                        'bcc': usuario_row.get('BCC', ''),
                        'expedientes': num_expedientes,
                        'asunto': asunto_procesado,
                        'mensaje': mensaje_base,
                        'cuerpo_mensaje': cuerpo_mensaje
                    })
                else:
                    st.info(f"ℹ️ Usuario {usuario} no tiene expedientes pendientes - No se enviará correo")
            
            if usuarios_para_envio:
                st.success(f"✅ {len(usuarios_para_envio)} usuarios tienen expedientes pendientes para enviar")
                
                # Mostrar resumen
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Usuarios activos", len(usuarios_activos))
                with col2:
                    st.metric("Con expedientes", len(usuarios_para_envio))
                with col3:
                    st.metric("Semana activa", f"Semana {num_semana}")
                
                # Mostrar tabla de usuarios para envío
                with st.expander("📋 Ver detalles de usuarios para envío"):
                    df_envio = pd.DataFrame(usuarios_para_envio)
                    columnas_mostrar = ['usuario', 'resumen', 'email', 'expedientes', 'asunto']
                    st.dataframe(df_envio[columnas_mostrar], use_container_width=True)
                
                # Previsualización de correo
                st.subheader("👁️ Previsualización del Correo")
                
                if usuarios_para_envio:
                    usuario_ejemplo = usuarios_para_envio[0]
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        st.write("**Destinatario:**", usuario_ejemplo['email'])
                        if usuario_ejemplo['cc']:
                            st.write("**CC:**", usuario_ejemplo['cc'])
                        if usuario_ejemplo['bcc']:
                            st.write("**BCC:**", usuario_ejemplo['bcc'])
                        st.write("**Asunto:**", usuario_ejemplo['asunto'])
                        st.write("**Expedientes:**", usuario_ejemplo['expedientes'])
                    
                    with col2:
                        st.text_area("**Cuerpo del Mensaje:**", usuario_ejemplo['cuerpo_mensaje'], height=200, key="preview_mensaje")
                
                # Botón de envío masivo
                st.markdown("---")
                st.subheader("🚀 Envío de Correos")
                
                st.warning("""
                **⚠️ Importante antes de enviar:**
                - Se usará la cuenta de Outlook predeterminada
                - No es necesario tener Outlook abierto
                - Los correos se enviarán inmediatamente
                - Se adjuntará el PDF individual de cada usuario
                - **Verifica que los datos sean correctos**
                """)
                
                if st.button("📤 Enviar Correos a Todos los Usuarios", type="primary", key="enviar_correos"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    correos_enviados = 0
                    correos_fallidos = 0
                    
                    for i, usuario_info in enumerate(usuarios_para_envio):
                        status_text.text(f"📨 Enviando a: {usuario_info['usuario']} ({usuario_info['email']})")
                        
                        # Generar PDF
                        pdf_data = generar_informe_usuario(usuario_info['usuario'])
                        
                        if pdf_data:
                            # Enviar correo con Outlook
                            nombre_archivo = f"Expedientes_Pendientes_{usuario_info['usuario']}_Semana_{num_semana}.pdf"
                            
                            exito = enviar_correo_outlook(
                                destinatario=usuario_info['email'],
                                asunto=usuario_info['asunto'],
                                cuerpo_mensaje=usuario_info['cuerpo_mensaje'],
                                archivo_pdf=pdf_data,
                                nombre_archivo=nombre_archivo,
                                cc=usuario_info.get('cc'),
                                bcc=usuario_info.get('bcc')
                            )
                            
                            if exito:
                                correos_enviados += 1
                                st.success(f"✅ Enviado a {usuario_info['usuario']}")
                            else:
                                correos_fallidos += 1
                                st.error(f"❌ Falló envío a {usuario_info['usuario']}")
                        else:
                            st.warning(f"⚠️ No se pudo generar PDF para {usuario_info['usuario']}")
                            correos_fallidos += 1
                        
                        progress_bar.progress((i + 1) / len(usuarios_para_envio))
                    
                    status_text.text("")
                    
                    # Mostrar resumen final
                    st.markdown("---")
                    st.subheader("📊 Resumen del Envío")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total procesados", len(usuarios_para_envio))
                    with col2:
                        st.metric("Correos enviados", correos_enviados, delta=f"+{correos_enviados}")
                    with col3:
                        st.metric("Correos fallidos", correos_fallidos, delta=f"-{correos_fallidos}", delta_color="inverse")
                    
                    if correos_enviados > 0:
                        st.balloons()
                        st.success("🎉 ¡Envío de correos completado!")
            
            else:
                st.warning("⚠️ No hay usuarios con expedientes pendientes para enviar")
    
    # Información de configuración
    st.markdown("---")
    with st.expander("⚙️ Configuración de Envío de Correos"):
        st.info("""
        **📋 Para que funcione el envío de correos:**
        
        1. **Outlook instalado** en el equipo
        2. **Cuenta de correo predeterminada** configurada en Outlook
        3. **Librería pywin32 instalada**: 
           ```bash
           pip install pywin32
           ```
        4. **Archivo USUARIOS.xlsx** con la estructura correcta
        
        **📋 Estructura del archivo USUARIOS.xlsx (Hoja Sheet1):**
        - USUARIOS: Código del usuario (debe coincidir con RECTAUTO)
        - ENVIAR: "SI" o "NO" 
        - EMAIL: Dirección de correo
        - ASUNTO: Puede usar &num_semana& y &fecha_max&
        - MENSAJE: Texto del mensaje
        - CC, BCC: Opcionales (separar múltiples con ;)
        - RESUMEN: Opcional (nombre del usuario)
        """)

elif eleccion == "Indicadores clave (KPI)":
    # ... (el código de la sección KPI se mantiene exactamente igual)
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
    st.header("🗓️ Selector de Semana")
    
    # Crear etiquetas formateadas para el slider
    opciones_slider = []
    for i, fecha in enumerate(semanas_disponibles):
        num_semana = ((fecha - FECHA_REFERENCIA).days) // 7 + 1
        fecha_str_opcion = fecha.strftime('%d/%m/%Y')
        opciones_slider.append(f"Semana {num_semana} ({fecha_str_opcion})")
    
    # Slider
    semana_index_slider = st.select_slider(
        "Selecciona la semana:",
        options=list(range(len(semanas_disponibles))),
        value=st.session_state.kpi_semana_index,
        format_func=lambda x: opciones_slider[x]
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

    # Función para calcular KPIs por semana
    def calcular_kpis_para_semana(_df, semana_fin):
        inicio_semana = semana_fin - timedelta(days=6)
        
        if 'FECHA APERTURA' in _df.columns:
            nuevos_expedientes = _df[
                (_df['FECHA APERTURA'] >= inicio_semana) & 
                (_df['FECHA APERTURA'] <= semana_fin)
            ].shape[0]
        else:
            nuevos_expedientes = 0
        
        if 'ESTADO' in _df.columns and 'FECHA ÚLTIMO TRAM.' in _df.columns:
            expedientes_cerrados = _df[
                (_df['ESTADO'] == 'Cerrado') & 
                (_df['FECHA ÚLTIMO TRAM.'] >= inicio_semana) & 
                (_df['FECHA ÚLTIMO TRAM.'] <= fin_semana)
            ].shape[0]
        else:
            expedientes_cerrados = 0

        if 'FECHA CIERRE' in _df.columns and 'FECHA APERTURA' in _df.columns:
            total_abiertos = _df[
                (_df['FECHA APERTURA'] <= fin_semana) & 
                ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
            ].shape[0]
        else:
            total_abiertos = 0
        
        return {
            'nuevos_expedientes': nuevos_expedientes,
            'expedientes_cerrados': expedientes_cerrados,
            'total_abiertos': total_abiertos
        }

    # CALCULAR KPIs PARA TODAS LAS SEMANAS con cache de 2 horas
    @st.cache_data(ttl=CACHE_TTL, show_spinner="📊 Calculando KPIs históricos...")
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

    # Gráfico de evolución - SOLO CACHEAR LOS DATOS, NO EL GRÁFICO COMPLETO
    @st.cache_data(ttl=CACHE_TTL)
    def obtener_datos_grafico_evolucion(_df_kpis):
        """Solo cachea los datos necesarios para el gráfico"""
        return _df_kpis.copy()

    def mostrar_kpis_principales(_df_kpis, _semana_seleccionada, _num_semana):
        kpis_semana = _df_kpis[_df_kpis['semana_numero'] == _num_semana].iloc[0]
        
        fecha_str = _semana_seleccionada.strftime('%d/%m/%Y')
        st.header(f"📊 KPIs de la Semana: {fecha_str} (Semana {_num_semana})")
        
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

    # GRÁFICO DE EVOLUCIÓN TEMPORAL (ACTUALIZADO) - CORREGIDO
    st.markdown("---")
    st.subheader("📈 Evolución Temporal de KPIs")

    # Obtener datos desde cache
    datos_grafico = obtener_datos_grafico_evolucion(df_kpis_semanales)

    # Crear gráfico completo con datos actualizados (SIEMPRE FRESCO)
    fig = px.line(
        datos_grafico,
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

    # Añadir línea vertical para la semana seleccionada (SIEMPRE ACTUALIZADA)
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fig.add_vline(
        x=num_semana_seleccionada, 
        line_width=2, 
        line_dash="dash", 
        line_color="red",
        annotation_text="Semana Seleccionada",
        annotation_position="top left"
    )

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
